# HardwareMVP - Modules Directory

This is where modules live. Anything in the directory will be inventoried and run by the HardwareMVP.vbs script. It defaults to daily but this can be changed via the specific module and/or in the database itself.
<br><br>
Here is some helpful information for each module/script:
<br><br>
KaseyaVSA.ps1 - Inventories Kaseya VSA agents. Uses Kaseya's API to collect live data when it is run using the credentials supplied. *<br>
LansweeperCSV.vbs - More on this later<br>
Nexpose.ps1 - Inventories Windows devices that have been scanned by Rapid7's Nexpose platform. Uses Nexpose's API to collect live data when it is run using the credentials supplied. *<br>
NexposeCSV.vbs - Alternate to the Powershell script above. Only one needs to be used. More on this later<br>
PDQ.vbs - More on this later<br>
SymantecEndpointManager.ps1 - Inventories devices on Symantec Endpoint Manager. Uses Symantec's API to collect live data when it is run using the credentials supplied. *<br>

<br><br>
<i>* Credentials should be supplied to the smapp.ini using the following format (substituting ModuleName for the name of the module). If data is not entered or incorrect, an email is generated.<br>
[ModuleName]<br>
ModuleNameServer=<br>
ModuleNameUser=<br>
ModuleNamePassword=<br>
</i>