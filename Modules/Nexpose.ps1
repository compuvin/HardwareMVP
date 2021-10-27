[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;

##################################
#Edit these values as you see fit
$PSSchema = "hardwaremvp"
$PSTbl = "nexpose"
$PSFN = "Rapid7 Nexpose"
$PSRunInt = 1
$PSML = 0
##################################

function Get-IniContent ($filePath)
{
    $ini = @{}
    switch -regex -file $FilePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = “Comment” + $CommentCount
            $ini[$section][$name] = $value
        }
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

$ModuleFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Source
$IniContent = Get-IniContent ((Split-Path -Parent -Path $ModuleFolder) + "\smapp.ini")
$ModuleData = New-Object System.Collections.ArrayList
$UpdateDate = Get-Date -Format "yyyy-MM-dd"
$DomainName = (Get-CimInstance Win32_ComputerSystem).Domain
$outputl = ""


[String] $Server = $IniContent[$PSTbl]["NexposeServer"]

#Check to make sure that the module info is in the INI
#We'll store this in the database at some point, don't worry
if (!$Server)
{
	$outputl = "<p>Additional information is needed for the " + $PSFN + " module to run. Please edit the smapp.ini file located at: " + (Split-Path -Parent -Path $ModuleFolder) + "</p>
		<p>Please copy and paste the lines below and edit the values as indicated:</p><p>
		[" + $PSTbl + "]<br>
		NexposeServer=#YourServerHere#<br>
		NexposeUser=#YourUserHere#<br>
		NexposePassword=#YourPasswordHere#"
	Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To ([string[]]($IniContent["Email"]["RptToEmail"]).Split(',')) -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject ("HardwareMVP: Action Required - " + $PSFN) -Body $outputl -BodyAsHtml
	Break
}

# Create credentiaal and convert to Base64String
$username = $IniContent[$PSTbl]["NexposeUser"]
$password = $IniContent[$PSTbl]["NexposePassword"]

[string]$authInfo = ([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('{0}:{1}' -f $username, $password))))


# format HTTP header
$header = @{ Authorization='Basic {0} ' -f $authInfo }

$SearchFilter = '{"match": "all","filters": [{ "field": "operating-system", "operator": "contains", "value": "Microsoft" }]}'

$ModuleData = (Invoke-RestMethod -Uri https://$($Server)/api/3/assets/search?size=10000 -Method Post -Headers $header -Body $SearchFilter -ContentType 'application/json') | select -ExpandProperty resources | select hostname, ip, os


#Open MySQL Connection
$myconnection = New-Object MySql.Data.MySqlClient.MySqlConnection
$myconnection.ConnectionString = "Database=" + $PSSchema + ";server=" + $IniContent["Database"]["DBLocation"] + ";Persist Security Info=false;user id=" + $IniContent["Database"]["DBUser"] + ";pwd=" + $IniContent["Database"]["DBPass"] + ";"
$myconnection.Open()
$command = $myconnection.CreateCommand()

#Create table for this module if it doesn't exist
$command.CommandText = "CREATE TABLE IF NOT EXISTS " + $PSSchema + "." + $PSTbl + " (ID INT PRIMARY KEY AUTO_INCREMENT, Name VARCHAR(255) UNIQUE, IPAddress text, SiteName text, OS text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL)";
$reader = $command.ExecuteNonQuery()

#Update the settings for this module using the variables set at the top of the script.
$command.CommandText = "UPDATE " + $PSSchema + ".modules SET `Name` = '" + $PSFN + "', `TableName` = '" + $PSTbl + "', `RunInterval` = '" + $PSRunInt + "', `MasterList` = '" + $PSML + "' WHERE (`FileName` = '" + $MyInvocation.MyCommand.Name + "')";
$reader = $command.ExecuteNonQuery()

#Get the Master List table name
if ($PSML -eq 0) #If this is not the master list
{
	$command.CommandText = "select TableName from " + $PSSchema + ".modules where MasterList = '1'";
	$MasterListTableName = $command.ExecuteScalar()
	#write-output $MasterListTableName
}

#Add/Update table entries
foreach ($ModuleRow in $ModuleData)
{
	if ($ModuleRow.hostname -ne "" -and $ModuleRow.hostname) #Ignore empty device names
	{
		$command.CommandText = "INSERT INTO " + $PSSchema + "." + $PSTbl + "(Name, IPAddress, SiteName, OS, FirstDiscovered, LastDiscovered) values('" + $ModuleRow.hostname.replace(("." + $DomainName),"").ToUpper() + "','" + $ModuleRow.ip + "','','" + $ModuleRow.os + "','" + $UpdateDate + "','" + $UpdateDate + "') ON DUPLICATE KEY UPDATE IPAddress='" + $ModuleRow.ip + "', OS='" + $ModuleRow.os + "', LastDiscovered='" + $UpdateDate + "'";
		$reader = $command.ExecuteNonQuery()
	}
}

#Report new devices
$command.CommandText = "select " + $PSTbl + ".*, (select count(*) from " + $PSSchema + "." + $MasterListTableName + " where " + $PSTbl + ".name = " + $MasterListTableName + ".name) MasterList from " + $PSTbl + " where FirstDiscovered='" + $UpdateDate + "'";
$reader = $command.ExecuteReader()

while ($reader.Read()) {
	if ($outputl.IndexOf('<p><b>Hardware Added:</b></p>') -eq -1) 
	{
		#Header Info
		$outputl = $outputl + "<p><b>Hardware Added:</b></p>"
		$outputl = $outputl + "<table>"
		$outputl = $outputl + "<tr>"
		$outputl = $outputl + "  <th>Computer</th>"
		$outputl = $outputl + "  <th>IP Address</th>"
		$outputl = $outputl + "  <th>Site Name</th>"
		$outputl = $outputl + "  <th>Operating System</th>"
		if ($PSML -eq 0) { $outputl = $outputl + "  <th>On Master List</th>" }
		$outputl = $outputl + "</tr>"
	}
	$outputl = $outputl + "<tr>"
	$outputl = $outputl + "  <td>" + $Reader["Name"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["IPAddress"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["SiteName"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["OS"].ToString() + "</td>"
	if ($PSML -eq 0) {
		if ($Reader["MasterList"].ToString() -eq 0)
		{
			$outputl = $outputl + "  <td bgcolor=#FF0000>No</td>"
		} else {
			$outputl = $outputl + "  <td>Yes</td>"
		}
	}
	$outputl = $outputl + "</tr>"
}
$reader.close()
if ($outputl.IndexOf('<p><b>Hardware Added:</b></p>') -ne -1) { $outputl = $outputl + "</table>" }

#Remove old entries
if ($ModuleData.length -gt 0)
{
	#Report removed devices
	$command.CommandText = "select " + $PSTbl + ".*, (select count(*) from " + $PSSchema + "." + $MasterListTableName + " where " + $PSTbl + ".name = " + $MasterListTableName + ".name) MasterList from " + $PSTbl + " where not LastDiscovered='" + $UpdateDate + "'";
	$reader = $command.ExecuteReader()
	
	while ($reader.Read()) {
		if ($outputl.IndexOf('<p><b>Hardware Removed:</b></p>') -eq -1) 
		{
			#Header Info
			$outputl = $outputl + "<p><b>Hardware Removed:</b></p>"
			$outputl = $outputl + "<table>"
			$outputl = $outputl + "<tr>"
			$outputl = $outputl + "  <th>Computer</th>"
			$outputl = $outputl + "  <th>IP Address</th>"
			$outputl = $outputl + "  <th>Site Name</th>"
			$outputl = $outputl + "  <th>Operating System</th>"
			if ($PSML -eq 0) { $outputl = $outputl + "  <th>On Master List</th>" }
			$outputl = $outputl + "</tr>"
		}
		$outputl = $outputl + "<tr>"
		$outputl = $outputl + "  <td>" + $Reader["Name"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["IPAddress"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["SiteName"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["OS"].ToString() + "</td>"
		if ($PSML -eq 0) {
			if ($Reader["MasterList"].ToString() -eq 0)
			{
				$outputl = $outputl + "  <td>No</td>"
			} else {
				$outputl = $outputl + "  <td bgcolor=#FF0000>Yes</td>"
			}
		}
		$outputl = $outputl + "</tr>"
	}
	$reader.close()
	
	#Perform Deletions
	$command.CommandText = "DELETE FROM " + $PSSchema + "." + $PSTbl + " WHERE NOT (`LastDiscovered` = '" + $UpdateDate + "')";
	$reader = $command.ExecuteNonQuery()
}

if ($outputl.IndexOf('<p><b>Hardware Removed:</b></p>') -ne -1) { $outputl = $outputl + "</table>" }

#Send email
if ($outputl -ne "")
{
	$outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" + $outputl
	Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To ([string[]]($IniContent["Email"]["RptToEmail"]).Split(',')) -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject ("HardwareMVP: " + $PSFN + " Report") -Body $outputl -BodyAsHtml
}

#Close MySQL Connection
$myconnection.Close()
