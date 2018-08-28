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
Dim dbrecord()
		Locator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy
Logla "programs","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do
	strComputer=ObjServer.ReadLine
	Logla "programs","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\CIMV2", strUser, strPass,"MS_409","NTLMDomain:" + strDomain)
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "programs", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406

	On Error Goto 0

	If Not IsObject(objWMIService) Then
		Logla "programs", "Connection Failed!!"
		Exit Do
	Else
		Logla "programs", "Connected"
	End If
	On Error Resume Next
	Err.Clear

	Set colItems = objWMIService.ExecQuery("Select Name,Version,InstallLocation FROM " & _
					"Win32_Product")
	If colItems.Count = 0  Then
		Wscript.Echo "Win32_Environment Failed!!"
		Logla "programs", "Win32_Environment Failed!!"
		Wscript.Echo "If the server is 2003 - In the Windows Components Wizard, select Management and Monitoring Tools and then click Details, Management and Monitoring Tools dialog box, select WMI Windows Installer Provider!!"
		Logla "programs", "If the server is 2003 - In the Windows Components Wizard, select Management and Monitoring Tools and then click Details, Management and Monitoring Tools dialog box, select WMI Windows Installer Provider!!"
		Exit Do
	End If
	Logla "programs", "Win32_Environment Success!!"
	On Error Goto 0
	i=0
	ReDim dbrecord(colItems.count,2)
	For Each objSoftware in colItems
		Wscript.Echo "Name: " & objSoftware.Name
				dbrecord(i,0)=objSoftware.Name
		Wscript.Echo "Version: " & objSoftware.Version
				dbrecord(i,1)=objSoftware.Version
		Wscript.Echo "Location: " & objSoftware.InstallLocation
				dbrecord(i,2)=objSoftware.InstallLocation
		i=i+1
	Next
	UpdateDBProducts strComputer, dbrecord, colItems.count

	
Loop While False
Loop



Sub UpdateDBProducts(ServerName,dbrecord,count)
	Logla "programs", "Connecting to DB"
	
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 
				On Error Resume Next
				If Err.Number<>0 Then
					Logla "programs", "DB Connection Failed!!"
					Exit Sub
				End If
				Logla "programs", "Connected to DB Success!!"
				Err.Clear
				On Error Goto 0
	Logla "programs", "Connecting InstalledPrograms DB"
	i=0
	For i = 0 to count-1
	
		objRecordset.Open "SELECT * FROM Win2016Mig.dbo.InstalledPrograms WHERE ""ServerName""='"&ServerName &"' AND ""Program""='"& dbrecord(i,0) &"' AND ""Version""='"& dbrecord(i,1) &"' AND ""Path""='"&dbrecord(i,2) &"'", objConnection, adOpenStatic, adLockOptimistic
		If objRecordset.EOF Then 
			objRecordset.Close
			Logla "programs", "Inserting the Record - Cannot Find Record "&dbrecord(i,0)&" on DB"
			ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.InstalledPrograms (""ServerName"", ""Program"", ""Version"", ""Path"") VALUES ('"&ServerName&"','"&dbrecord(i,0) &"','"&dbrecord(i,1) &"','"&dbrecord(i,2) &"'); " , objConnection, adOpenStatic, adLockOptimistic
		Else
			Logla "programs", "No Update - The Record "&dbrecord(i,0)&" is already in the Table"
		End If
		If objRecordset.State= 1 then
			objRecordset.Close
		End If
	Next

objConnection.Close 
End Sub

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



