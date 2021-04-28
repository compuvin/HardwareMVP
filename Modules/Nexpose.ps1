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
$NexposeData = New-Object System.Collections.ArrayList
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
	Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To $IniContent["Email"]["RptToEmail"] -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject ("HardwareMVP: Action Required - " + $PSFN) -Body $outputl -BodyAsHtml
	Break
}

# Create credentiaal and convert to Base64String
$username = $IniContent[$PSTbl]["NexposeUser"]
$password = $IniContent[$PSTbl]["NexposePassword"]

[string]$authInfo = ([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('{0}:{1}' -f $username, $password))))


# format HTTP header
$header = @{ Authorization='Basic {0} ' -f $authInfo }

$SearchFilter = '{"match": "all","filters": [{ "field": "operating-system", "operator": "contains", "value": "Microsoft" }]}'

$NexposeData = (Invoke-RestMethod -Uri https://$($Server)/api/3/assets/search?size=10000 -Method Post -Headers $header -Body $SearchFilter -ContentType 'application/json') | select -ExpandProperty resources | select hostname, ip, os


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

#Add/Update table entries
foreach ($NexposeRow in $NexposeData)
{
	if ($NexposeRow.hostname -ne "" -and $NexposeRow.hostname) #Ignore empty device names
	{
		$command.CommandText = "INSERT INTO " + $PSSchema + "." + $PSTbl + "(Name, IPAddress, SiteName, OS, FirstDiscovered, LastDiscovered) values('" + $NexposeRow.hostname.replace(("." + $DomainName),"").ToUpper() + "','" + $NexposeRow.ip + "','','" + $NexposeRow.os + "','" + $UpdateDate + "','" + $UpdateDate + "') ON DUPLICATE KEY UPDATE IPAddress='" + $NexposeRow.ip + "', OS='" + $NexposeRow.os + "', LastDiscovered='" + $UpdateDate + "'";
		$reader = $command.ExecuteNonQuery()
	}
}

#Remove old entries
if ($NexposeData.length -gt 0)
{
	$command.CommandText = "DELETE FROM " + $PSSchema + "." + $PSTbl + " WHERE NOT (`LastDiscovered` = '" + $UpdateDate + "')";
	$reader = $command.ExecuteNonQuery()
}

#$command.CommandText = "select * from " + $PSTbl;
#$dataSet = New-Object System.Data.DataSet

#$reader = $command.ExecuteReader()


#while ($reader.Read()) {
#  for ($i= 0; $i -lt $reader.FieldCount; $i++) {
#    write-output $reader.GetValue($i).ToString()
#  }
#}

#Send email
if ($outputl -ne "") { Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To $IniContent["Email"]["RptToEmail"] -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject "HardwareMVP: " + $PSFN + " Report" -Body $outputl -BodyAsHtml }

#Close MySQL Connection
$myconnection.Close()
