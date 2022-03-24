[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;

##################################
#Edit these values as you see fit
$PSSchema = "hardwaremvp"
$PSTbl = "TaegisXDR"
$PSFN = "Taegis XDR"
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
$outputl = ""


[String] $Server = $IniContent[$PSTbl]["TaegisServer"]

#Check to make sure that the module info is in the INI
#We'll store this in the database at some point, don't worry
if (!$Server)
{
	$outputl = "<p>Additional information is needed for the " + $PSFN + " module to run. Please edit the smapp.ini file located at: " + (Split-Path -Parent -Path $ModuleFolder) + "</p>
		<p>Please copy and paste the lines below and edit the values as indicated:</p><p>
		[" + $PSTbl + "]<br>
		TaegisServer=#YourServerHere#<br>
		TaegisClientID=#YourUserHere#<br>
		TaegisClientSecret=#YourPasswordHere#"
	Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To ([string[]]($IniContent["Email"]["RptToEmail"]).Split(',')) -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject ("HardwareMVP: Action Required - " + $PSFN) -Body $outputl -BodyAsHtml
	Break
}

$TaegisClientID = $IniContent[$PSTbl]["TaegisClientID"]
$TaegisClientSecret = $IniContent[$PSTbl]["TaegisClientSecret"]

# Get Auth token
$body = "grant_type=client_credentials&client_id=$TaegisClientID&client_secret=$TaegisClientSecret"


$authToken = (Invoke-RestMethod -Uri https://$($Server)/auth/api/v2/auth/token -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded')

# format HTTP header
$header = @{ Authorization='Bearer '+ $authToken.access_token }

$allAssetsQuery = "{
   allAssets(
     offset: 0,
     limit: 10000,
     order_by: hostname,
     filter_asset_state: All
   )
   {
     totalResults
     assets {
       id
       hostId
       sensorVersion
       hostnames {
         hostname
       }
       osVersion
     }
   }
 }"
$body = @{query= $allAssetsQuery} | ConvertTo-Json

$ModuleData = (Invoke-RestMethod -Uri https://$($Server)/graphql -method post -Headers $header -body $body -ContentType 'application/json').data.allAssets.assets | select hostId, id, osVersion, sensorVersion -ExpandProperty hostnames


#Open MySQL Connection
$myconnection = New-Object MySql.Data.MySqlClient.MySqlConnection
$myconnection.ConnectionString = "Database=" + $PSSchema + ";server=" + $IniContent["Database"]["DBLocation"] + ";Persist Security Info=false;user id=" + $IniContent["Database"]["DBUser"] + ";pwd=" + $IniContent["Database"]["DBPass"] + ";"
$myconnection.Open()
$command = $myconnection.CreateCommand()

#Create table for this module if it doesn't exist
$command.CommandText = "CREATE TABLE IF NOT EXISTS " + $PSSchema + "." + $PSTbl + " (ID INT PRIMARY KEY AUTO_INCREMENT, Name VARCHAR(255) UNIQUE, HostID text, TaegisID text, OS text, Version text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL)";
$reader = $command.ExecuteNonQuery()

#Update the settings for this module using the variables set at the top of the script
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
	$command.CommandText = "INSERT INTO " + $PSSchema + "." + $PSTbl + "(Name, HostID, TaegisID, OS, Version, FirstDiscovered, LastDiscovered) values('" + $ModuleRow.hostname.ToUpper() + "','" + $ModuleRow.hostId + "','" + $ModuleRow.id + "','" + $ModuleRow.osVersion + "','" + $ModuleRow.sensorVersion + "','" + $UpdateDate + "','" + $UpdateDate + "') ON DUPLICATE KEY UPDATE HostID='" + $ModuleRow.hostId + "', TaegisID='" + $ModuleRow.id + "', OS='" + $ModuleRow.osVersion + "', Version='" + $ModuleRow.sensorVersion + "', LastDiscovered='" + $UpdateDate + "'";
	$reader = $command.ExecuteNonQuery()
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
		$outputl = $outputl + "  <th>Host ID</th>"
		$outputl = $outputl + "  <th>Taegis ID</th>"
		$outputl = $outputl + "  <th>Operating System</th>"
		$outputl = $outputl + "  <th>Sensor Version</th>"
		if ($PSML -eq 0) { $outputl = $outputl + "  <th>On Master List</th>" }
		$outputl = $outputl + "</tr>"
	}
	$outputl = $outputl + "<tr>"
	$outputl = $outputl + "  <td>" + $Reader["Name"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["hostId"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["TaegisID"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["OS"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["Version"].ToString() + "</td>"
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
			$outputl = $outputl + "  <th>Host ID</th>"
			$outputl = $outputl + "  <th>Taegis ID</th>"
			$outputl = $outputl + "  <th>Operating System</th>"
			$outputl = $outputl + "  <th>Sensor Version</th>"
			if ($PSML -eq 0) { $outputl = $outputl + "  <th>On Master List</th>" }
			$outputl = $outputl + "</tr>"
		}
		$outputl = $outputl + "<tr>"
		$outputl = $outputl + "  <td>" + $Reader["Name"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["hostId"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["TaegisID"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["OS"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["Version"].ToString() + "</td>"
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
