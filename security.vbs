includeFile "c:\scripts\variables.config"
includeFile "c:\scripts\logla.vbs"
Set objConnection = CreateObject("ADODB.Connection")
Set objConnection2 = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")
Set objRecordset2 = CreateObject("ADODB.Recordset")
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjServer = ObjFSO.OpenTextFile("c:\scripts\ServerList.list", ForReading)
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
Set Locator = CreateObject("WbemScripting.SWbemLocator")
		Locator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy
Logla "security","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do
	strComputer=ObjServer.ReadLine
	Logla "security","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\CIMV2", strUser, strPass,"MS_409","NTLMDomain:" + strDomain)
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "security", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			Else
				Logla "security", "Connected"
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406
		On Error Goto 0

	If Not IsObject(objWMIService) Then
		Logla "security", "Connection Failed!!"
		Exit Do
	End If
'	On Error Goto 0
					  
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 

	objRecordset.Open "SELECT [Path] FROM [Win2016Mig].[dbo].[IISPaths] WHERE ""ServerDomain""='"&strDomain &"' AND ""ServerName""='"&strComputer&"'", objConnection, adOpenStatic, adLockOptimistic
	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
		Path=Trim(objRecordSet.Fields.Item("Path"))
		Wscript.Echo Replace(Path,"\","\\")
		Retry = 0
		Do
			On Error Resume Next
			Err.Clear

			Set objFile = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & _
Replace(Path,"\","\\")&"'" )
			Retry=Retry + 1
'			Wscript.Echo "ErrorNumber After:"&Err.Number
'			Wscript.Echo "ErrorNumber After:"&Err.Description
			If Err.Number <> 0 then
				Wscript.Echo "Unable to get Directory Security. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			End if	
		Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406
		On Error Goto 0

		If objFile.GetSecurityDescriptor(objSD) = 0 Then
			
			For Each objAce in objSD.DACL
	
				Wscript.Echo ""
				Wscript.Echo "Trustee: " & objAce.Trustee.Name
				Wscript.Echo "TrusteeDomain: " & objACE.Trustee.Domain
				Wscript.Echo "Ace Flags: " & objAce.AceFlags    
				Wscript.Echo "ACE Type: " & objACE.AceType
				Wscript.Echo "Access Mask: " & objAce.AccessMask
				Wscript.Echo "ControlFlags: " & objSD.ControlFlags
				UpdateDBSecurity strDomain, strComputer,Path,objAce.Trustee.Name,objAce.Trustee.Domain,objAce.AceFlags,objAce.AceType,objAce.AccessMask,objSD.ControlFlags
			Next
			
		End If
	objRecordSet.MoveNext
	Loop
Loop While False
Loop
  
Sub UpdateDBSecurity(ServerDomain, ServerName,Path,Trustee,TrusteeDomain,AceFlags,AceType,AccessMask,ControlFlags)

	objConnection2.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNMAE;Password=PASSWORD;" 


	objRecordset2.Open "SELECT * FROM Win2016Mig.dbo.Security WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Path""='"&Path &"' AND ""Trustee""='"&Trustee &"' AND ""TrusteeDomain""='"&TrusteeDomain & "'", objConnection2, adOpenStatic, adLockOptimistic
	If objRecordset2.EOF Then 
		ObjRecordset2.Close
		ObjRecordset2.Open "INSERT INTO Win2016Mig.dbo.Security (""ServerDomain"", ""ServerName"",""Path"",""Trustee"",""TrusteeDomain"",""AceFlags"",""AceType"",""AccessMask"",""ControlFlags"") VALUES ('"&ServerDomain&"','"&ServerName&"','"&Path&"','"&Trustee&"','"&TrusteeDomain&"','"&AceFlags&"','"&AceType&"','"&AccessMask&"','"&ControlFlags&"'); " , objConnection2, adOpenStatic, adLockOptimistic	
	End If
objConnection2.Close 
End Sub
objConnection.Close 

Sub includeFile(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

Function LZ(ByVal Number)
  If Number < 10 Then
    LZ = "0" & CStr(Number)
  Else
    LZ = CStr(Number)
  End If
End Function

Function TimeStamp
  Dim CurrTime
  CurrTime = Now()

  TimeStamp = CStr(Year(CurrTime)) & "-" _
    & LZ(Month(CurrTime)) & "-" _
    & LZ(Day(CurrTime)) & " " _
    & LZ(Hour(CurrTime)) & ":" _
    & LZ(Minute(CurrTime)) & ":" _
    & LZ(Second(CurrTime))
End Function

'https://technet.microsoft.com/en-us/library/2006.05.scriptingguy.aspx