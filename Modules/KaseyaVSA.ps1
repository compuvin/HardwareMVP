[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;

##################################
#Edit these values as you see fit
$PSSchema = "hardwaremvp"
$PSTbl = "kaseya"
$PSFN = "Kaseya VSA"
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


[String] $Server = $IniContent[$PSTbl]["KaseyaServer"]

#Check to make sure that the module info is in the INI
#We'll store this in the database at some point, don't worry
if (!$Server)
{
	$outputl = "<p>Additional information is needed for the " + $PSFN + " module to run. Please edit the smapp.ini file located at: " + (Split-Path -Parent -Path $ModuleFolder) + "</p>
		<p>Please copy and paste the lines below and edit the values as indicated:</p><p>
		[" + $PSTbl + "]<br>
		KaseyaServer=#YourServerHere#<br>
		KaseyaUser=#YourUserHere#<br>
		KaseyaPassword=#YourPasswordHere#"
	Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To ([string[]]($IniContent["Email"]["RptToEmail"]).Split(',')) -SmtpServer $IniContent["Email"]["EmailSvr"] -Subject ("HardwareMVP: Action Required - " + $PSFN) -Body $outputl -BodyAsHtml
	Break
}

#C# class to create callback
$code = @"
public class SSLHandler
{
    public static System.Net.Security.RemoteCertificateValidationCallback GetSSLHandler()
    {

        return new System.Net.Security.RemoteCertificateValidationCallback((sender, certificate, chain, policyErrors) => { return true; });
    }

}
"@

#compile the class
Add-Type -TypeDefinition $code

#disable checks using new class
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = [SSLHandler]::GetSSLHandler()

function New-Hash {
	
	Param(

		[Parameter(Mandatory=$True)]
		[ValidateSet('SHA1','SHA256')]
		[string]$Algorithm,
		
		[Parameter(Mandatory=$True)]
		[string]$Text
	
	)
	try {
		$data = [system.Text.Encoding]::UTF8.GetBytes($Text)
		[string]$hash = -join ([Security.Cryptography.HashAlgorithm]::Create($Algorithm).ComputeHash($data) | ForEach-Object { "{0:x2}" -f $_ })
	}
	catch{
		$_.Exception.Message
	}
	return $hash

}

$kaseyaApiUrl = $IniContent[$PSTbl]["KaseyaServer"]
$kaseyaApiUser = $IniContent[$PSTbl]["KaseyaUser"]
$kaseyaApiPswd = $IniContent[$PSTbl]["KaseyaPassword"]

# Hash Kaseya VSA password
$RandomNumber = Get-Random -Minimum 10000000 -Maximum 99999999
$RawSHA256Hash = New-Hash -Algorithm 'SHA256' -Text "$kaseyaApiPswd"
$CoveredSHA256HashTemp = New-Hash -Algorithm 'SHA256' -Text "$kaseyaApiPswd$kaseyaApiUser"
$CoveredSHA256Hash = New-Hash -Algorithm 'SHA256' -Text "$CoveredSHA256HashTemp$RandomNumber"
$RawSHA1Hash = New-Hash -Algorithm 'SHA1' -Text "$kaseyaApiPswd"
$CoveredSHA1HashTemp = New-Hash -Algorithm 'SHA1' -Text "$kaseyaApiPswd$kaseyaApiUser"
$CoveredSHA1Hash = New-Hash -Algorithm 'SHA1' -Text "$CoveredSHA1HashTemp$RandomNumber"

# Create base64 encoded authentication for Authorization header 
$auth = 'user={0},pass2={1},pass1={2},rpass2={3},rpass1={4},rand2={5}' -f $kaseyaApiUser, $CoveredSHA256Hash, $CoveredSHA1Hash, $RawSHA256Hash, $RawSHA1Hash, $RandomNumber
$authBase64Encoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($auth)) 

# Define parameters for Invoke-WebRequest cmdlet
$params = [ordered] @{
	Uri         	= '{0}/api/v1.0/auth' -f $kaseyaApiUrl
	Method      	= 'GET'
	ContentType 	= 'application/json; charset=utf-8'
	Headers     	= @{'Authorization' = "Basic $authBase64Encoded"}
}

# Fetch new access token
try {
	$response = Invoke-RestMethod @params
}
catch{
	$_.Exception.Message
}
New-Variable -Name kaseyaApiAccessToken -value $response.result.token -Scope Script -Force

# Set local variables
$results = @()
$totalRecords = 0
$skip = 0
$top = 100

do  {
    # Add API parameters
	$params = [ordered] @{
		Uri         	= '{0}/api/v1.0{1}{2}{3}{4}' -f $kaseyaApiUrl, '/assetmgmt/agents', ('?$skip={0}&$top={1}' -f $skip, $top), $Filter, $OrderBy
		Method      	= 'GET'
		ContentType 	= 'application/json; charset=utf-8'
		Headers     	= @{'Authorization' = 'Bearer {0}' -f $kaseyaApiAccessToken}
	}

	# Add body request
	If ($ApiRequestBody) {$params.Add('Body',$ApiRequestBody)}
	
	# Invoke API request
	try { 
		$response = Invoke-RestMethod @params
	}
	catch {
		Write-Output $_.Exception.Message -f Red
		if($_.ErrorDetail) {Write-Output $_.ErrorDetail.Message -f Red}
		Write-Output $_.ScriptStackTrace -f Red
		exit 1
	}
    if($response) {$results += $response.result}

    # Get total records
    if($totalRecords -eq 0) {$totalRecords = [int]$response.totalrecords}
        
    # Update paging
    $skip += $top
}
until ($skip -ge $totalRecords) 


$ModuleData = ($results | select ComputerName, IPAddress, MachineGroup, OSType, OperatingSystem, CpuType, SystemSerialNumber)



#Open MySQL Connection
$myconnection = New-Object MySql.Data.MySqlClient.MySqlConnection
$myconnection.ConnectionString = "Database=" + $PSSchema + ";server=" + $IniContent["Database"]["DBLocation"] + ";Persist Security Info=false;user id=" + $IniContent["Database"]["DBUser"] + ";pwd=" + $IniContent["Database"]["DBPass"] + ";"
$myconnection.Open()
$command = $myconnection.CreateCommand()

#Create table for this module if it doesn't exist
$command.CommandText = "CREATE TABLE IF NOT EXISTS " + $PSSchema + "." + $PSTbl + " (ID INT PRIMARY KEY AUTO_INCREMENT, Name VARCHAR(255) UNIQUE, IPAddress text, MachineGroup text, OS text, OSDetails text, CPUName text, SerialNumber text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL)";
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
	if ($ModuleRow.ComputerName -ne "" -and $ModuleRow.ComputerName) #Ignore empty device names
	{
		$command.CommandText = "INSERT INTO " + $PSSchema + "." + $PSTbl + "(Name, IPAddress, MachineGroup, OS, OSDetails, CPUName, SerialNumber, FirstDiscovered, LastDiscovered) values('" + $ModuleRow.ComputerName.replace(("." + $DomainName),"").ToUpper() + "','" + $ModuleRow.IPAddress + "','" + $ModuleRow.MachineGroup + "','" + $ModuleRow.OSType + "','" + $ModuleRow.OperatingSystem + "','" + $ModuleRow.CpuType.Substring(0, $ModuleRow.CpuType.IndexOf(",")) + "','" + $ModuleRow.SystemSerialNumber + "','" + $UpdateDate + "','" + $UpdateDate + "') ON DUPLICATE KEY UPDATE IPAddress='" + $ModuleRow.IPAddress + "', MachineGroup='" + $ModuleRow.MachineGroup + "', OS='" + $ModuleRow.OSType + "', OSDetails='" + $ModuleRow.OperatingSystem + "', CPUName='" + $ModuleRow.CpuType.Substring(0, $ModuleRow.CpuType.IndexOf(",")) + "', SerialNumber='" + $ModuleRow.SystemSerialNumber + "', LastDiscovered='" + $UpdateDate + "'";
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
		$outputl = $outputl + "  <th>Machine Group</th>"
		$outputl = $outputl + "  <th>OS</th>"
		$outputl = $outputl + "  <th>OS Details</th>"
		$outputl = $outputl + "  <th>CPU</th>"
		$outputl = $outputl + "  <th>Serial Number</th>"
		if ($PSML -eq 0) { $outputl = $outputl + "  <th>On Master List</th>" }
		$outputl = $outputl + "</tr>"
	}
	$outputl = $outputl + "<tr>"
	$outputl = $outputl + "  <td>" + $Reader["Name"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["IPAddress"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["MachineGroup"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["OS"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["OSDetails"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["CPUName"].ToString() + "</td>"
	$outputl = $outputl + "  <td>" + $Reader["SerialNumber"].ToString() + "</td>"
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
			$outputl = $outputl + "  <th>Machine Group</th>"
			$outputl = $outputl + "  <th>OS</th>"
			$outputl = $outputl + "  <th>OS Details</th>"
			$outputl = $outputl + "  <th>CPU</th>"
			$outputl = $outputl + "  <th>Serial Number</th>"
			if ($PSML -eq 0) { $outputl = $outputl + "  <th>On Master List</th>" }
			$outputl = $outputl + "</tr>"
		}
		$outputl = $outputl + "<tr>"
		$outputl = $outputl + "  <td>" + $Reader["Name"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["IPAddress"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["MachineGroup"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["OS"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["OSDetails"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["CPUName"].ToString() + "</td>"
		$outputl = $outputl + "  <td>" + $Reader["SerialNumber"].ToString() + "</td>"
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
