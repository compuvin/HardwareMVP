dim OVersion 'Oldest Installed version
dim CVersion 'Installed version
dim outputl 'Email body
Dim AllApps 'Data from CSV
dim WPData 'Web page text
Dim yfound 'For new apps, series of tests to find similar apps
Dim mfound 'New modules, if any are found
Dim ModuleSQL 'SQL string for multiple modules
Dim UpdatePageQTH, UpdatePageQTHVarience 'Used to fix any integer values in the two fields that are actually NULL
Dim adoconn
Dim rs
Dim str
set filesys=CreateObject("Scripting.FileSystemObject")
set xmlhttp = createobject("msxml2.serverxmlhttp.3.0")
Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir = filesys.GetParentFolderName(Wscript.ScriptFullName)
Dim Response 'For answers to prompts
Dim PSSchema, PSTbl 'Define schema and table names
PSSchema = "hardwaremvp"
PSTbl = "modules"

'Gather variables from smapp.ini or prompt for them and save them for next time
If filesys.FileExists(strCurDir & "\smapp.ini") then
	'Database
	'CSVPath = ReadIni(strCurDir & "\smapp.ini", "Database", "CSVPath" )
	DBLocation = ReadIni(strCurDir & "\smapp.ini", "Database", "DBLocation" )
	DBUser = ReadIni(strCurDir & "\smapp.ini", "Database", "DBUser" )
	DBPass = ReadIni(strCurDir & "\smapp.ini", "Database", "DBPass" )
	
	'Email - Defaults to anonymous login
	RptToEmail = ReadIni(strCurDir & "\smapp.ini", "Email", "RptToEmail" )
	RptFromEmail = ReadIni(strCurDir & "\smapp.ini", "Email", "RptFromEmail" )
	EmailSvr = ReadIni(strCurDir & "\smapp.ini", "Email", "EmailSvr" )
	'Additional email settings found in Function SendMail()
	
	'WebGUI
	'BaseURL = ReadIni(strCurDir & "\smapp.ini", "WebGUI", "BaseURL" )
else
	msgbox "INI file not found at: " & strCurDir & "\smapp.ini" & vbCrlf & "You will now be prompted with questions to create it."
	
	'Database
	'CSVPath = inputbox("Enter the location where the CSV file with the software dump can be found (UNC path recommended):", "HardwareMVP", strCurDir & "\Applications.csv")
	DBLocation = inputbox("Enter the IP address or hostname for the location of the database:", "HardwareMVP", "localhost")
	DBUser = inputbox("Enter the user name to access database on " & DBLocation & ":", "HardwareMVP", "user")
	DBPass = inputbox("Enter the password to access database on " & DBLocation & ":", "HardwareMVP", "P@ssword1")
	
	'Check to see if DB exists
	CheckForTables
	
	'Email - Defaults to anonymous login
	RptToEmail = inputbox("Enter the report email's To address:", "HardwareMVP", "admin@company.com")
	RptFromEmail = inputbox("Enter the report email's From address:", "HardwareMVP", "admin@company.com")
	EmailSvr = inputbox("Enter the FQDN or IP address of email server:", "HardwareMVP", "mail.server.com")
	msgbox "Additional email settings found in Function SendMail()"
	
	'WebGUI
	'BaseURL = inputbox("Enter the base URL for the HardwareMVP GUI (Web GUI available at https://github.com/compuvin/SoftwareMatrix-GUI):", "HardwareMVP", "http://www.intranet.com")
		
	'Write the data to INI file
	'WriteIni strCurDir & "\smapp.ini", "Database", "CSVPath", CSVPath
	WriteIni strCurDir & "\smapp.ini", "Database", "DBLocation", DBLocation
	WriteIni strCurDir & "\smapp.ini", "Database", "DBUser", DBUser
	WriteIni strCurDir & "\smapp.ini", "Database", "DBPass", DBPass
	WriteIni strCurDir & "\smapp.ini", "Email", "RptToEmail", RptToEmail
	WriteIni strCurDir & "\smapp.ini", "Email", "RptFromEmail", RptFromEmail
	WriteIni strCurDir & "\smapp.ini", "Email", "EmailSvr", EmailSvr
	'WriteIni strCurDir & "\smapp.ini", "WebGUI", "BaseURL", EditURL
end if

outputl = ""

Set adoconn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & DBLocation & ";" & _
   "Database=" & PSSchema & "; User=" & DBUser & "; Password=" & DBPass & ";"

CheckForModules 'Check to see if new modules exist
CleanupModules 'Remove tables for deleted modules after a week

if outputl <> "" then
	outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" & vbcrlf & outputl
	SendMail RptToEmail, "HardwareMVP: Modules Changed"
	outputl = ""
end if

ProcessModules 'Run scheduler
ProcessMasterList

if outputl <> "" then
	outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" & vbcrlf & outputl
	SendMail RptToEmail, "HardwareMVP: Devices Removed"
	outputl = ""
end if


'Check to see if new modules exist
Function CheckForModules()
	Dim fso, f, fld, fl, Name
	Set fso = CreateObject("scripting.filesystemobject")
	Set fld = fso.GetFolder(strCurDir & "\Modules")

	mfound = ""
	ModuleSQL = ""
	For Each f In fld.Files
		if right(lcase(f.name),4) = ".vbs" or right(lcase(f.name),4) = ".ps1" or right(lcase(f.name),4) = ".bat" then 'Supported module extentions: Visual Basic Script (vbs), PowerShell (ps1), Batch file (bat)
			ModuleName = left(f.name,(len(f.name)-4))
			
			str = "Select * from modules where FileName='" & f.name & "';"
			rs.Open str, adoconn, 3, 3 'OpenType, LockType
			
			if rs.eof then
				str = "INSERT INTO modules(Name,FileName,LastDiscovered,RunInterval,NextRunDate,MasterList) values('" & ModuleName & "','" & f.name & "','" & format(date(), "YYYY-MM-DD") & "','1','" & format(date(), "YYYY-MM-DD")  & "','0');"
				adoconn.Execute(str)
				
				mfound = mfound & ModuleName & "|"
			else
				if not rs("MasterList") = "1" and not rs("TableName") & "" = "" then
					if not ModuleSQL = "" then ModuleSQL = ModuleSQL & ", "
					ModuleSQL = ModuleSQL & "(select count(Name) from " & PSSchema & "." & rs("TableName") & " where " & rs("TableName") & ".Name = %MasterList%.Name) `" & rs("Name") &"`"
				end if
				rs("LastDiscovered") = format(date(), "YYYY-MM-DD")
				rs.update
			end if
			rs.close
		end if
	Next

    If not mfound = "" Then
		'This is where we'll send an email listing the new modules found
		'msgbox "New module found: " & vbCrlf & replace(mfound, "|", vbCrlf)
		
		'Header Info
		outputl = outputl & "<p><b>The following new modules have been added:</b></p>" & vbcrlf
		outputl = outputl & "<table>" & vbcrlf
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <th>Name</th>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		outputl = outputl & "<tr><td>" &	replace(mfound, "|", "</td></tr><tr><td>")
		outputl = left(outputl,len(outputl)-8)
		outputl = outputl & "</table>" & vbcrlf
    End If

End Function

'Remove tables for deleted modules
Function CleanupModules()
	'msgbox "Here is where we would cleanup modules"
	
	str = "Select * from modules where LastDiscovered IS NOT NULL and not LastDiscovered = '" & format(date(), "YYYY-MM-DD") & "';"
	rs.Open str, adoconn, 3, 3 'OpenType, LockType
	
	if not rs.eof then
		'Header Info
		outputl = outputl & "<p><b>The following modules are no longer found in the Modules directory:</b></p>" & vbcrlf
		outputl = outputl & "<table>" & vbcrlf
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <th>Name</th>" & vbcrlf
		outputl = outputl & "  <th>File Name</th>" & vbcrlf
		outputl = outputl & "  <th>Master List</th>" & vbcrlf
		outputl = outputl & "  <th>Table Removal Date</th>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		rs.MoveFirst
	end if
	
	do while not rs.eof
		'msgbox "We would send an email here about this module being deleted on " & (rs("LastDiscovered") + 7) & ": " & rs("Name")
		
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("FileName") & "</td>" & vbcrlf
		if rs("MasterList") = "1" then
			outputl = outputl & "  <td bgcolor=#FF0000>Yes</td>" & vbcrlf
		else
			outputl = outputl & "  <td></td>" & vbcrlf
		end if
		if rs("TableName") & "" = "" then
			outputl = outputl & "  <td>N/A</td>" & vbcrlf
			rs.delete
		elseif cdate(rs("LastDiscovered")) < (Date() - 7) then
			outputl = outputl & "  <td bgcolor=#FF0000>" & (rs("LastDiscovered") + 7) & "</td>" & vbcrlf
			
			str = "DROP TABLE `" & PSSchema & "`.`" & rs("TableName") & "`;"
			adoconn.Execute(str)
			
			rs.delete
		else
			outputl = outputl & "  <td>" & (rs("LastDiscovered") + 7) & "</td>" & vbcrlf
		end if
		outputl = outputl & "</tr>" & vbcrlf

		rs.movenext
		if rs.eof then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
End Function

'Run any modules that are scheduled to run
Function ProcessModules()
	Set objShell = Wscript.CreateObject("Wscript.Shell")
	Dim MScript
	'msgbox "Here is where we would process modules"
	
	str = "Select * from modules where NextRunDate IS NOT NULL and NextRunDate <= '" & format(date(), "YYYY-MM-DD") & "' and LastDiscovered IS NOT NULL and LastDiscovered = '" & format(date(), "YYYY-MM-DD") & "';"
	rs.Open str, adoconn, 3, 3 'OpenType, LockType
	
	do while not rs.eof
		'msgbox "We would run this module now because it is set to run on or after " & (rs("NextRunDate") + 7) & ": " & rs("Name")
		
		if right(lcase(rs("FileName")),4) = ".vbs" then
			MScript = "wscript """ & strCurDir & "\modules\" & rs("FileName") & """"
		elseif right(lcase(rs("FileName")),4) = ".ps1" then
			MScript = "%systemroot%\System32\WindowsPowerShell\V1.0\PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass """ & strCurDir & "\modules\" & rs("FileName") & """"
		elseif right(lcase(rs("FileName")),4) = ".bat" then
			MScript = "cmd /c """ & strCurDir & "\modules\" & rs("FileName") & """"
		else
			MScript = rs("FileName")
		end if
		
		'Run module script
		objShell.Run MScript, 0, True ' The script will continue until it is closed.
		'objShell.Run MScript, 1, True ' Swap to this if you want the window to display on-screen
		
		'Update scheduler for the next run time
		rs("NextRunDate") = format((date() + rs("RunInterval")), "YYYY-MM-DD")
		
		rs.update
		rs.movenext
	loop
	
	rs.close
End Function

'Check for any removed devices that exist in other tables
Function ProcessMasterList()
	Dim MLTableName
	Dim NOCML 'Number of columns on MasterList
	Dim NOTC 'Number of total columns
	Dim RemoveRS 'Should we remove the record
	
	NOTC = 0
	
	str = "Select * from modules where MasterList = '1';"
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
	if not rs.eof then
		rs.movefirst
		MLTableName = rs("TableName")
		if cdate(rs("NextRunDate")) = cdate(format((date() + rs("RunInterval")), "YYYY-MM-DD")) then
			str = "select count(*)  FROM information_schema.columns where table_schema = '" & PSSchema & "' and table_name = '" & MLTableName & "';"
			NOCML = (adoconn.Execute(str))(0)
			NOCML = cint(NOCML)
			
			do while not ModuleSQL = replace(ModuleSQL,"%MasterList%",MLTableName)
				ModuleSQL = replace(ModuleSQL,"%MasterList%",MLTableName,1,1)
				NOTC = NOTC + 1
			loop
			if NOTC > 0 then ModuleSQL = ", " & ModuleSQL
			
			str = "Select " & MLTableName & ".*" & ModuleSQL & " from " & MLTableName & " where LastDiscovered IS NOT NULL and not LastDiscovered = (select max(LastDiscovered) from " & MLTableName & ");"
		else
			str = ""
		end if
	end if
	rs.close
	
	if not str = "" then
		'response = inputbox("sql:", "Test", str)
		rs.Open str, adoconn, 3, 3 'OpenType, LockType
		
		if not rs.eof then
			rs.MoveFirst
			
			'Header Info
			outputl = outputl & "<p><b>Devices removed from Master List:</b></p>" & vbcrlf
			outputl = outputl & "<table>" & vbcrlf
			outputl = outputl & "<tr>" & vbcrlf
			
			for i = 0 to (NOCML+NOTC-1)
				if not rs.Fields.Item(i).Name = "ID" then outputl = outputl & "  <th>" & rs.Fields.Item(i).Name & "</th>" & vbcrlf
			Next
			
			outputl = outputl & "</tr>" & vbcrlf
		end if
		
		do while not rs.eof
			RemoveRS = True
			outputl = outputl & "<tr>" & vbcrlf
			for i = 0 to (NOCML-1)
				if not rs.Fields.Item(i).Name = "ID" then outputl = outputl & "  <td>" & rs(i) & "</td>" & vbcrlf
			Next
			for i = NOCML to (NOCML+NOTC-1)
				if rs(i) = "0" then
					outputl = outputl & "  <td>No</td>" & vbcrlf
				else
					outputl = outputl & "  <td bgcolor=#FF0000>Yes</td>" & vbcrlf
					RemoveRS = False
				end if
			Next
			outputl = outputl & "</tr>" & vbcrlf
			
			if RemoveRS = True then
				str = "DELETE FROM `" & PSSchema & "`.`" & MLTableName & "` WHERE (`ID` = '" & rs("ID") & "');"
				adoconn.Execute(str)
			end if
			
			rs.movenext
			if rs.eof then outputl = outputl & "</table>" & vbcrlf
		loop
		rs.close
	end if
End Function


Function SendMail(TextRcv,TextSubject)
  Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
  Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

  Const cdoAnonymous = 0 'Do not authenticate
  Const cdoBasic = 1 'basic (clear-text) authentication
  Const cdoNTLM = 2 'NTLM

  Set objMessage = CreateObject("CDO.Message") 
  objMessage.Subject = TextSubject 
  objMessage.From = RptFromEmail 
  objMessage.To = TextRcv
  objMessage.HTMLBody = outputl

  '==This section provides the configuration information for the remote SMTP server.

  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

  'Name or IP of Remote SMTP Server
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailSvr

  'Type of authentication, NONE, Basic (Base64 encoded), NTLM
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoAnonymous

  'Server port (typically 25)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

  'Use SSL for the connection (False or True)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

  'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

  objMessage.Configuration.Fields.Update

  '==End remote SMTP server configuration section==

  objMessage.Send
End Function

Function Format(vExpression, sFormat)
  Dim nExpression
  nExpression = sFormat
  
  if isnull(vExpression) = False then
    if instr(1,sFormat,"Y") > 0 or instr(1,sFormat,"M") > 0 or instr(1,sFormat,"D") > 0 or instr(1,sFormat,"H") > 0 or instr(1,sFormat,"S") > 0 then 'Time/Date Format
      vExpression = cdate(vExpression)
	  if instr(1,sFormat,"AM/PM") > 0 and int(hour(vExpression)) > 12 then
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression)-12,2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)-12) '1 character hour
		nExpression = replace(nExpression,"AM/PM","PM") 'If if its afternoon, its PM
	  else
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression),2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)) '1 character hour
		if int(hour(vExpression)) = 12 then nExpression = replace(nExpression,"AM/PM","PM") '12 noon is PM while anything else in this section is AM (fixed 04/19/2019 thanks to our HR Dept.)
		nExpression = replace(nExpression,"AM/PM","AM") 'If its not PM, its AM
	  end if
	  nExpression = replace(nExpression,":MM",":" & right("00" & minute(vExpression),2)) '2 character minute
	  nExpression = replace(nExpression,"SS",right("00" & second(vExpression),2)) '2 character second
	  nExpression = replace(nExpression,"YYYY",year(vExpression)) '4 character year
	  nExpression = replace(nExpression,"YY",right(year(vExpression),2)) '2 character year
	  nExpression = replace(nExpression,"DD",right("00" & day(vExpression),2)) '2 character day
	  nExpression = replace(nExpression,"D",day(vExpression)) '(N)N format day
	  nExpression = replace(nExpression,"MMM",left(MonthName(month(vExpression)),3)) '3 character month name
	  if instr(1,sFormat,"MM") > 0 then
	    nExpression = replace(nExpression,"MM",right("00" & month(vExpression),2)) '2 character month
	  else
	    nExpression = replace(nExpression,"M",month(vExpression)) '(N)N format month
	  end if
    elseif instr(1,sFormat,"N") > 0 then 'Number format
	  nExpression = vExpression
	  if instr(1,sFormat,".") > 0 then 'Decimal format
	    if instr(1,nExpression,".") > 0 then 'Both have decimals
		  do while instr(1,sFormat,".") > instr(1,nExpression,".")
		    nExpression = "0" & nExpression
		  loop
		  if len(nExpression)-instr(1,nExpression,".") >= len(sFormat)-instr(1,sFormat,".") then
		    nExpression = left(nExpression,instr(1,nExpression,".")+len(sFormat)-instr(1,sFormat,"."))
	      else
		    do while len(nExpression)-instr(1,nExpression,".") < len(sFormat)-instr(1,sFormat,".")
			  nExpression = nExpression & "0"
			loop
	      end if
		else
		  nExpression = nExpression & "."
		  do while len(nExpression) < len(sFormat)
			nExpression = nExpression & "0"
		  loop
	    end if
	  else
		do while len(nExpression) < sFormat
		  nExpression = "0" and nExpression
		loop
	  end if
	else
      msgbox "Formating issue on page. Unrecognized format: " & sFormat
	end if
	
	Format = nExpression
  end if
End Function

'Read text file
function GetFile(FileName)
  If FileName<>"" Then
    Dim FS, FileStream
    Set FS = CreateObject("Scripting.FileSystemObject")
      on error resume Next
      Set FileStream = FS.OpenTextFile(FileName)
      GetFile = FileStream.ReadAll
  End If
End Function

Function ReadIni( myFilePath, mySection, myKey ) 'Thanks to http://www.robvanderwoude.com
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function

Sub WriteIni( myFilePath, mySection, myKey, myValue ) 'Thanks to http://www.robvanderwoude.com
    ' This subroutine writes a value to an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be written
    ' myValue     [string]  the value to be written (myKey will be
    '                       deleted if myValue is <DELETE_THIS_VALUE>)
    '
    ' Returns:
    ' N/A
    '
    ' CAVEAT:     WriteIni function needs ReadIni function to run
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre, Johan Pol and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni
    Dim strFilePath, strFolderPath, strKey, strLeftString
    Dim strLine, strSection, strTempDir, strTempFile, strValue

    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    strValue    = Trim( myValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
End Sub

'Check to see if database and tables exist
Function CheckForTables()
	Dim CreatePS2DB 'Boolean for DB creation
	CreatePS2DB = False
	
	Set adoconn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & DBLocation & ";" & _
			"User=" & DBUser & "; Password=" & DBPass & ";"
			
	str = "SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '" & PSSchema & "'"
	rs.CursorLocation = 3 'adUseClient
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
	if rs.eof then
		Response = msgbox("The database does not exist. Would you like to create it now? (Make sure the user """ & DBUser & """ has permission to do so)", vbYesNo)
		if Response = vbYes then
			CreatePS2DB = True
		else
			WScript.Quit
		end if
		rs.close
	else
		'msgbox "DB exists"
		rs.close
		
		'Double check to make sure table is also there
		str = "SELECT * FROM information_schema.tables WHERE table_schema = '" & PSSchema & "' AND table_name = '" & PSTbl & "' LIMIT 1;"
		rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
		if rs.eof then
			Response = msgbox("The database exists but the table does not exist. Would you like to create it now?", vbYesNo)
			if Response = vbYes then
				CreatePS2DB = True
			else
				WScript.Quit
			end if
			rs.close
		else
			'msgbox "Table exists"
			rs.close
		end if
	end if
	
	'Create schema and/or table if needed
	if CreatePS2DB = True then
		'Create schema if not there
		str = "CREATE DATABASE IF NOT EXISTS " & PSSchema & ";"
		adoconn.Execute(str)
		
		'Create tables
		PSTbl = "modules"
		str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, Name text, FileName text, TableName text, LastDiscovered date DEFAULT NULL, RunInterval int(11) DEFAULT '1', NextRunDate date DEFAULT NULL, MasterList int(11) DEFAULT '0');"
		adoconn.Execute(str)
		
		'PSTbl = "goldhardware"
		'str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, Name text, IPAddress text, OS text, OSVersion text, OSSP text, Manufacturer text, SerialNumber text, Memory text, CPUName text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL);"
		'adoconn.Execute(str)
		
	end if
	
	Set adoconn = Nothing
	Set rs = Nothing
End Function