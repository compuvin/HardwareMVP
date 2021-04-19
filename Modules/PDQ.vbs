Dim adoconn
Dim rs
Dim str
Dim i 'Counter
Dim outputl 'Email body
Dim CSVPath
Dim AllHW
set filesys=CreateObject("Scripting.FileSystemObject")
Dim strCurDir
strCurDir = filesys.GetParentFolderName(Wscript.ScriptFullName)
strCurDir = filesys.GetParentFolderName(strCurDir) 'Get the parent folder of the Modules folder
Dim PSSchema, PSTbl,PSFN, PSRunInt, PSML 'Define schema and table names
PSSchema = "hardwaremvp"
PSTbl = "pdq"
PSFN = "PDQ Inventory" 'Friendly Name for module
PSRunInt = 1 'Module run interval (in days)
PSML = 1 'Is this the master list?
CSVPath = strCurDir & "\source\HardwareInventory.csv"

'Gather variables from smapp.ini
If filesys.FileExists(strCurDir & "\smapp.ini") then
	'Database
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
	msgbox "INI file not found at: " & strCurDir & "\smapp.ini" & vbCrlf & "Please run HardwareMVP.vbs first before running this file."
end if


outputl = ""

If filesys.FileExists(CSVPath) then
	AllHW = GetFile(CSVPath)
	
	Set adoconn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & DBLocation & ";" & _
		"Database=" & PSSchema & "; User=" & DBUser & "; Password=" & DBPass & ";"

	'Create the table for this module if it doesn't exist
	str = "CREATE TABLE IF NOT EXISTS " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, Name text, IPAddress text, OS text, OSVersion text, OSSP text, Manufacturer text, ModelName text, SerialNumber text, Memory text, CPUName text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL);"
	adoconn.Execute(str)
	
	'Modify the module entry for this module
	str = "UPDATE " & PSSchema & ".modules SET `Name` = '" & PSFN & "', `TableName` = '" & PSTbl & "', `RunInterval` = '" & PSRunInt & "', `MasterList` = '" & PSML & "' WHERE (`FileName` = '" & Wscript.ScriptName & "');"
	adoconn.Execute(str)
	
	IngestCSV
	
	if outputl <> "" then
		outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" & vbcrlf & outputl
		SendMail RptToEmail, "HardwareMVP: " & PSFN & " Report"
		outputl = ""
	end if
	
	filesys.DeleteFile CSVPath, force
end if



Function IngestCSV()
	Dim CurrPC, IPAddress, OS, OSVersion, OSSP, Manufacturer, ModelName, SerialNumber, Memory, CPUName
	
	'Computer Name,Computer IP Address,Computer O/S,Computer O/S Version,Computer SP / Release,Computer Manufacturer,Computer Model,Computer Serial Number,Computer Memory,CPU Name,Computer Successful Scan Date
	'ID, Name, IPAddress, OS, OSVersion, OSSP, Manufacturer, ModelName, SerialNumber, Memory, CPUName, FirstDiscovered, LastDiscovered

	'PCs - Whats new/old/changed
	AllHW = right(AllHW,len(AllHW)-206)
	do while len(AllHW) > 10
		'Get PC name
		CurrPC = mid(AllHW,1,instr(1,AllHW,",",1)-1)
		AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		'msgbox CurrPC
		'Get IP Address
		if left(AllHW,1)="""" then
			IPAddress = mid(AllHW,2,instr(1,AllHW,""",",1)-2)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		else
			IPAddress = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox IPAddress
		'Get OS
		if left(AllHW,1)="""" then
			OS = mid(AllHW,2,instr(1,AllHW,""",",1)-2)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		else
			OS = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox OS
		'Get OSVersion
		if left(AllHW,1)="""" then
			OSVersion = mid(AllHW,2,instr(1,AllHW,",",1)-3)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		elseif instr(1,AllHW,",",1) - 1 =< 0 then
			OSVersion = "0"
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
			'msgbox CurrApp & " No version!"
		else
			OSVersion = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox OSVersion
		'Get OSSP
		if left(AllHW,1)="""" then
			OSSP = mid(AllHW,2,instr(1,AllHW,",",1)-3)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		elseif instr(1,AllHW,",",1) - 1 =< 0 then
			OSSP = "0"
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
			'msgbox CurrApp & " No version!"
		else
			OSSP = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'Get Manufacturer
		if left(AllHW,1)="""" then
			Manufacturer = mid(AllHW,2,instr(1,AllHW,""",",1)-2)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		else
			Manufacturer = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox Manufacturer
		'Get ModelName
		if left(AllHW,1)="""" then
			ModelName = mid(AllHW,2,instr(1,AllHW,""",",1)-2)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		else
			ModelName = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox ModelName
		'Get SerialNumber
		if left(AllHW,1)="""" then
			SerialNumber = mid(AllHW,2,instr(1,AllHW,""",",1)-2)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		else
			SerialNumber = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox SerialNumber
		'Get Memory
		if left(AllHW,1)="""" then
			Memory = mid(AllHW,2,instr(1,AllHW,""",",1)-2)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		else
			Memory = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox Memory
		'Get CPUName
		if left(AllHW,1)="""" then
			CPUName = mid(AllHW,2,instr(1,AllHW,""",",1)-2)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,""",",1)-1)
		else
			CPUName = mid(AllHW,1,instr(1,AllHW,",",1)-1)
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,",",1))
		end if
		'msgbox CPUName
		'Get LastScaned Date (not used)
		if left(AllHW,1)="""" then
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,vbCrlf,1)-1)
		else
			AllHW = right(AllHW,len(AllHW)-instr(1,AllHW,vbCrlf,1)-1)
		end if
		
		
		'msgbox CurrPC & vbCrlf & IPAddress & vbCrlf & OS & vbCrlf & OSVersion & vbCrlf & OSSP & vbCrlf & Manufacturer & vbCrlf & ModelName & vbCrlf & SerialNumber & vbCrlf & Memory & vbCrlf & CPUName
		
		str = "Select * from " & PSTbl & " where Name='" & CurrPC & "';"
		rs.Open str, adoconn, 3, 3 'OpenType, LockType
		if not rs.eof then
			rs.MoveFirst
			if len(rs("LastDiscovered") & "") = 0 then rs("LastDiscovered") = "2001-01-01" 'Fix DB issues
			if len(rs("FirstDiscovered") & "") = 0 then rs("FirstDiscovered") = format(date()-1, "YYYY-MM-DD") 'Fix DB issues
			if format(rs("LastDiscovered"), "YYYY-MM-DD") <> format(date(), "YYYY-MM-DD") then
				rs("LastDiscovered") = format(date(), "YYYY-MM-DD")
				'msgbox "date"
			end if
			
			' if not rs("Version") = CurrVer then
				' if instr(1,outputl,"<p><b>Software Added or Changed:</b></p>",1) = 0 then
					' 'Header Info
					' outputl = outputl & "<p><b>Software Added or Changed:</b></p>" & vbcrlf
					' outputl = outputl & "<table>" & vbcrlf
					' outputl = outputl & "<tr>" & vbcrlf
					' outputl = outputl & "  <th>Computer</th>" & vbcrlf
					' outputl = outputl & "  <th>Application</th>" & vbcrlf
					' outputl = outputl & "  <th>Publisher</th>" & vbcrlf
					' outputl = outputl & "  <th>Previous Version</th>" & vbcrlf
					' outputl = outputl & "  <th>New Version</th>" & vbcrlf
					' outputl = outputl & "</tr>" & vbcrlf
				' end if
				
				' outputl = outputl & "<tr>" & vbcrlf
				' outputl = outputl & "  <td>" & CurrPC & "</td>" & vbcrlf
				' outputl = outputl & "  <td>" & CurrApp & "</td>" & vbcrlf
				' outputl = outputl & "  <td>" & CurrPub & "</td>" & vbcrlf
				' outputl = outputl & "  <td>" & rs("Version") & "</td>" & vbcrlf
				' outputl = outputl & "  <td>" & CurrVer & "</td>" & vbcrlf
				' outputl = outputl & "</tr>" & vbcrlf
				
				' 'msgbox CurrApp & ": Updated on " & CurrPC & " from " & rs("Version") & " to " & CurrVer
				' rs("Version") = CurrVer
				' rs("Publisher") = CurrPub
			' end if
			
			'Update PC entry with the latest info
			rs("IPAddress") = IPAddress 
			rs("OS") = OS
			rs("OSVersion") = OSVersion
			rs("OSSP") = OSSP
			rs("Manufacturer") = Manufacturer
			rs("ModelName") = ModelName
			rs("SerialNumber") = SerialNumber
			rs("Memory") = Memory
			rs("CPUName") = CPUName
			
			'msgbox CurrPC & " - " & IPAddress & ": finished updating"
			
			rs.update
		else
			if instr(1,outputl,"<p><b>Hardware Added:</b></p>",1) = 0 then
				'Header Info
				outputl = outputl & "<p><b>Hardware Added:</b></p>" & vbcrlf
				outputl = outputl & "<table>" & vbcrlf
				outputl = outputl & "<tr>" & vbcrlf
				outputl = outputl & "  <th>Computer</th>" & vbcrlf
				outputl = outputl & "  <th>IPAddress</th>" & vbcrlf
				outputl = outputl & "  <th>OSVersion</th>" & vbcrlf
				outputl = outputl & "  <th>OSSP</th>" & vbcrlf
				outputl = outputl & "  <th>Manufacturer</th>" & vbcrlf
				outputl = outputl & "  <th>ModelName</th>" & vbcrlf				
				outputl = outputl & "  <th>SerialNumber</th>" & vbcrlf				
				outputl = outputl & "  <th>Memory</th>" & vbcrlf				
				outputl = outputl & "  <th>CPUName</th>" & vbcrlf
				outputl = outputl & "</tr>" & vbcrlf
			end if
			
			outputl = outputl & "<tr>" & vbcrlf
			outputl = outputl & "  <td>" & CurrPC & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & IPAddress & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & OSVersion & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & OSSP & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & Manufacturer & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & ModelName & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & SerialNumber & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & Memory & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & CPUName & "</td>" & vbcrlf
			outputl = outputl & "</tr>" & vbcrlf
			
			str = "INSERT INTO "  & PSSchema & "." & PSTbl & "(Name, IPAddress, OS, OSVersion, OSSP, Manufacturer, ModelName, SerialNumber, Memory, CPUName, FirstDiscovered, LastDiscovered) values('" & CurrPC & "','" & IPAddress & "','" & OS & "','" & OSVersion & "','" & OSSP & "','" & Manufacturer & "','" & ModelName & "','" & SerialNumber & "','" & Memory & "','" & CPUName & "','" & format(date(), "YYYY-MM-DD")  & "','" & format(date(), "YYYY-MM-DD") & "');"
			adoconn.Execute(str)
			
			'msgbox "Added: " & CurrPC & " - " & IPAddress
		end if
		rs.close
		
	loop

	if instr(1,outputl,"<p><b>Hardware Added:</b></p>",1) > 0 then outputl = outputl & "</table>" & vbcrlf
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
