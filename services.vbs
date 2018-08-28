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
Logla "services","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do
	strComputer=ObjServer.ReadLine
	Logla "services","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\CIMV2", strUser, strPass,"MS_409","NTLMDomain:" + strDomain)
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "services", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406

	If Not IsObject(objWMIService) Then
		Logla "services", "Connection Failed!!"
		Exit Do
	Else
		Logla "services", "Connected"
	End If

	Err.Clear
   Set colProcess = objWMIService.ExecQuery ("SELECT * FROM " & _
                    "Win32_Service ")
	If Not IsObject(colProcess) Then
		Logla "services", "Win32_Service Failed!!"
		Exit Do
	End If
	Logla "services", "Win32_Service Success!!"
	On Error Goto 0
	i=0
	ReDim dbrecord(colProcess.count,24)
    For Each objItem In colProcess

		Wscript.Echo "AcceptPause: " & objItem.AcceptPause
		dbrecord(i,0)=objItem.AcceptPause
		Wscript.Echo "AcceptStop: " & objItem.AcceptStop
		dbrecord(i,1)=objItem.AcceptStop
		Wscript.Echo "Caption: " & objItem.Caption
		dbrecord(i,2)=objItem.Caption
		Wscript.Echo "CheckPoint: " & objItem.CheckPoint
		dbrecord(i,3)=objItem.CheckPoint
		Wscript.Echo "CreationClassName: " & objItem.CreationClassName
		dbrecord(i,4)=objItem.CreationClassName
		If Not IsNull(objItem.Description) Then
			Description=Replace(objItem.Description,"'","")
		End If
		dbrecord(i,5)=Description
		Wscript.Echo "DesktopInteract: " & objItem.DesktopInteract
		dbrecord(i,6)=objItem.DesktopInteract
		Wscript.Echo "DisplayName: " & objItem.DisplayName
		dbrecord(i,7)=objItem.DisplayName
		Wscript.Echo "ErrorControl: " & objItem.ErrorControl
		dbrecord(i,8)=objItem.ErrorControl
		Wscript.Echo "ExitCode: " & objItem.ExitCode
		dbrecord(i,9)=objItem.ExitCode
		Wscript.Echo "InstallDate: " & objItem.InstallDate
		dbrecord(i,10)=objItem.InstallDate
		Wscript.Echo "Name: " & objItem.Name
		dbrecord(i,11)=objItem.Name
		Wscript.Echo "PathName: " & objItem.PathName
		dbrecord(i,12)=objItem.PathName
		Wscript.Echo "ProcessId: " & objItem.ProcessId
		dbrecord(i,13)=objItem.ProcessId
		Wscript.Echo "ServiceSpecificExitCode: " & objItem.ServiceSpecificExitCode
		dbrecord(i,14)=objItem.ServiceSpecificExitCode
		Wscript.Echo "ServiceType: " & objItem.ServiceType
		dbrecord(i,15)=objItem.ServiceType
		Wscript.Echo "Started: " & objItem.Started
		dbrecord(i,16)=objItem.Started
		Wscript.Echo "StartMode: " & objItem.StartMode
		dbrecord(i,17)=objItem.StartMode
		Wscript.Echo "StartName: " & objItem.StartName
		dbrecord(i,18)=objItem.StartName
		Wscript.Echo "State: " & objItem.State
		dbrecord(i,19)=objItem.State
		Wscript.Echo "Status: " & objItem.Status
		dbrecord(i,20)=objItem.Status
		Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
		dbrecord(i,21)=objItem.SystemCreationClassName
		Wscript.Echo "SystemName: " & objItem.SystemName
		dbrecord(i,22)=objItem.SystemName
		Wscript.Echo "TagId: " & objItem.TagId
		dbrecord(i,23)=objItem.TagId
		Wscript.Echo "WaitHint: " & objItem.WaitHint	  
		dbrecord(i,24)=objItem.WaitHint
		i=i+1
	Next
	UpdateDBServices strDomain,strComputer, dbrecord, colProcess.count
Loop While False
Loop



Sub UpdateDBServices(ServerDomain, ServerName,dbrecord,count)
	Logla "services", "Connecting to DB"
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 
				On Error Resume Next
				If Err.Number<>0 Then
					Logla "services", "DB Connection Failed!!"
					Exit Sub
				End If
				Logla "services", "Connected to DB Success!!"
				Err.Clear
				On Error Goto 0
	Logla "services", "Connecting InstalledServices DB"
	i=0
	For i = 0 to count-1
		objRecordset.Open "SELECT * FROM Win2016Mig.dbo.InstalledServices WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Caption""='"&dbrecord(i,2) &"' AND ""Name""='"&dbrecord(i,11) &"' AND ""PathName""='"&dbrecord(i,12) &"'", objConnection, adOpenStatic, adLockOptimistic
		If objRecordset.EOF Then 
			objRecordset.Close
			Logla "services", "Inserting the Record - Cannot Find Record "&dbrecord(i,2)&" on DB"
'			Logla "services", "INSERT INTO Win2016Mig.dbo.InstalledServices (""ServerDomain"", ""ServerName"",""AcceptPause"",""AcceptStop"",""Caption"",""CheckPoint"",""CreationClassName"",""Description"",""DesktopInteract"",""DisplayName"",""ErrorControl"",""ExitCode"",""InstallDate"",""Name"",""PathName"",""ProcessId"",""ServiceSpecificExitCode"",""ServiceType"",""Started"",""StartMode"",""StartName"",""State"",""Status"",""SystemCreationClassName"",""SystemName"",""TagId"",""WaitHint"") VALUES ('"&ServerDomain&"','"&ServerName&"','"&dbrecord(i,0)&"','"&dbrecord(i,1)&"','"&dbrecord(i,2)&"','"&dbrecord(i,3)&"','"&dbrecord(i,4)&"','"&dbrecord(i,5)&"','"&dbrecord(i,6)&"','"&dbrecord(i,7)&"','"&dbrecord(i,8)&"','"&dbrecord(i,9)&"', '"&dbrecord(i,10)&"','"&dbrecord(i,11)&"','"&dbrecord(i,12)&"','"&dbrecord(i,13)&"','"&dbrecord(i,14)&"','"&dbrecord(i,15)&"','"&dbrecord(i,16)&"','"&dbrecord(i,17)&"','"&dbrecord(i,18)&"','"&dbrecord(i,19)&"','"&dbrecord(i,20)&"','"&dbrecord(i,21)&"', '"&dbrecord(i,22)&"','"&dbrecord(i,23)&"','"&dbrecord(i,24) &"'); "
			ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.InstalledServices (""ServerDomain"", ""ServerName"",""AcceptPause"",""AcceptStop"",""Caption"",""CheckPoint"",""CreationClassName"",""Description"",""DesktopInteract"",""DisplayName"",""ErrorControl"",""ExitCode"",""InstallDate"",""Name"",""PathName"",""ProcessId"",""ServiceSpecificExitCode"",""ServiceType"",""Started"",""StartMode"",""StartName"",""State"",""Status"",""SystemCreationClassName"",""SystemName"",""TagId"",""WaitHint"") VALUES ('"&ServerDomain&"','"&ServerName&"','"&dbrecord(i,0)&"','"&dbrecord(i,1)&"','"&dbrecord(i,2)&"','"&dbrecord(i,3)&"','"&dbrecord(i,4)&"','"&dbrecord(i,5)&"','"&dbrecord(i,6)&"','"&dbrecord(i,7)&"','"&dbrecord(i,8)&"','"&dbrecord(i,9)&"', '"&dbrecord(i,10)&"','"&dbrecord(i,11)&"','"&dbrecord(i,12)&"','"&dbrecord(i,13)&"','"&dbrecord(i,14)&"','"&dbrecord(i,15)&"','"&dbrecord(i,16)&"','"&dbrecord(i,17)&"','"&dbrecord(i,18)&"','"&dbrecord(i,19)&"','"&dbrecord(i,20)&"','"&dbrecord(i,21)&"', '"&dbrecord(i,22)&"','"&dbrecord(i,23)&"','"&dbrecord(i,24) &"'); " , objConnection, adOpenStatic, adLockOptimistic
		Else
			Logla "services", "No Update - The Record "&dbrecord(i,2)&" is already in the Table"
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

