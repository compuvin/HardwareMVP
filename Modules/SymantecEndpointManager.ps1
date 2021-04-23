[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;

##################################
#Edit these values as you see fit
$PSSchema = "hardwaremvp"
$PSTbl = "SymantecEndpointManager"
$PSFN = "Symantec Endpoint Manager"
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
$SEPData = New-Object System.Collections.ArrayList
$UpdateDate = Get-Date -Format "yyyy-MM-dd"
$outputl = ""


[String] $Server = $IniContent[$PSTbl]["SEPServer"]

#Check to make sure the SEP info is in the INI
#We'll store this in the database at some point, don't worry
if (!$Server)
{
	$outputl = "<p>Additional information is needed for the " + $PSFN + " module to run. Please edit the smapp.ini file located at: " + (Split-Path -Parent -Path $ModuleFolder) + "</p>
		<p>Please copy and paste the lines below and edit the values as indicated:</p><p>
		[" + $PSTbl + "]<br>
		SEPServer=#YourServerHere#<br>
		SEPUser=#YourUserHere#<br>
		SEPPassword=#YourPasswordHere#<br>
		SEPDomain=#YourDomainHere#"
	Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To $IniContent["Email"]["RptToEmail"] -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject ("HardwareMVP: Action Required - " + $PSFN) -Body $outputl -BodyAsHtml
	Break
}

# Create credentiaal and convert to JSON
$cred = @{
    username = $IniContent[$PSTbl]["SEPUser"]
    password = $IniContent[$PSTbl]["SEPPassword"]
    domain   = $IniContent[$PSTbl]["SEPDomain"]
}
$auth = $cred | ConvertTo-Json


$authToken = (Invoke-RestMethod -Uri https://$($Server)/sepm/api/v1/identity/authenticate -Method Post -Body $auth -ContentType 'application/json')

# format HTTP header
$header = @{ Authorization='Bearer '+ $authToken.Token }

$SEPData = (Invoke-RestMethod -Uri https://$($Server)/sepm/api/v1/computers -Headers $header -Body @{ pageSize='1000' }) | select -ExpandProperty content | select computerName, ipAddresses


$myconnection = New-Object MySql.Data.MySqlClient.MySqlConnection
$myconnection.ConnectionString = "Database=" + $PSSchema + ";server=" + $IniContent["Database"]["DBLocation"] + ";Persist Security Info=false;user id=" + $IniContent["Database"]["DBUser"] + ";pwd=" + $IniContent["Database"]["DBPass"] + ";"
$myconnection.Open()
$command = $myconnection.CreateCommand()

$command.CommandText = "CREATE TABLE IF NOT EXISTS " + $PSSchema + "." + $PSTbl + " (ID INT PRIMARY KEY AUTO_INCREMENT, Name VARCHAR(255) UNIQUE, IPAddress text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL)";
$reader = $command.ExecuteNonQuery()

$command.CommandText = "UPDATE " + $PSSchema + ".modules SET `Name` = '" + $PSFN + "', `TableName` = '" + $PSTbl + "', `RunInterval` = '" + $PSRunInt + "', `MasterList` = '" + $PSML + "' WHERE (`FileName` = '" + $MyInvocation.MyCommand.Name + "')";
$reader = $command.ExecuteNonQuery()

foreach ($SEPRow in $SEPData)
{
	$command.CommandText = "INSERT INTO " + $PSSchema + "." + $PSTbl + "(Name, IPAddress, FirstDiscovered, LastDiscovered) values('" + $SEPRow.computerName + "','" + $SEPRow.ipAddresses + "','" + $UpdateDate + "','" + $UpdateDate + "') ON DUPLICATE KEY UPDATE IPAddress='" + $SEPRow.ipAddresses + "', LastDiscovered='" + $UpdateDate + "'";
	$reader = $command.ExecuteNonQuery()
}

if ($SEPData.length -gt 0)
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

if ($outputl -ne "") { Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To $IniContent["Email"]["RptToEmail"] -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject "HardwareMVP: " + $PSFN + " Report" -Body $outputl -BodyAsHtml }

$myconnection.Close()
