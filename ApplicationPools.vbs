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
Logla "apppools","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do
	strComputer=ObjServer.ReadLine
	Logla "apppools","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\MicrosoftIISv2", strUser, strPass,"MS_409","NTLMDomain:" + strDomain)
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "apppools", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406

		If Not IsObject(objWMIService) Then
			Logla "apppools", "Connection Failed!!"
			Exit Do
		End If
	Err.Clear

	Set colItems = objWMIService.ExecQuery("SELECT * FROM IIsApplicationPoolSetting") 
		Logla "apppools", "Getting IIsApplicationPoolSetting"
		If Not IsObject(colItems) Then
			Logla "apppools", "IIsApplicationPoolSetting Failed!!"
			Exit Do
		End If
		Logla "apppools", "IIsApplicationPoolSetting Success!!"
	On Error Goto 0
	i=0
	ReDim dbrecord(colItems.count,6)
	For Each objItem in colItems 

		Wscript.Echo "ApplicationPool: " & objItem.Name
		dbrecord(i,0)=objItem.Name
		If objItem.AppPoolIdentityType = 0 then
			Identity="NT AUTHORITY\SYSTEM"
		ElseIf objItem.AppPoolIdentityType = 1 then
			Identity="NT AUTHORITY\LOCAL SERVICE"
		ElseIf objItem.AppPoolIdentityType = 2 then
			Identity="NT AUTHORITY\NETWORK SERVICE"
		ElseIf objItem.AppPoolIdentityType = 3 then
			Identity="WAMUser"
		End If 
		dbrecord(i,1)=objItem.AppPoolIdentityType		
		Wscript.Echo "Identity : " & Identity 
		Wscript.Echo "WAMUserName : " & objItem.WAMUserName 
		dbrecord(i,2)=objItem.WAMUserName		
		Wscript.Echo "WAMUserPass : " & objItem.WAMUserPass 
		dbrecord(i,3)=objItem.WAMUserPass	
		PeriodicRestartSchedule=""
		For Each objItem1 in objItem.PeriodicRestartSchedule  
			PeriodicRestartSchedule=PeriodicRestartSchedule & "," & objItem1
		Next
		PeriodicRestartSchedule=RIGHT(PeriodicRestartSchedule, LEN(PeriodicRestartSchedule)-1)
		Wscript.Echo "PeriodicRestartSchedule : " &PeriodicRestartSchedule
		dbrecord(i,4)=PeriodicRestartSchedule	
		On Error Resume Next
		Err.Clear
		Wscript.Echo "ManagedPipelineMode : "& objItem.ManagedPipelineMode
		If Err.Number <> 0 Then
			WScript.Echo "Error in Getting ManagedPipelineMode: " & Err.Description
			ManagedPipelineMode="1"
			Err.Clear
		Else
			ManagedPipelineMode=objItem.ManagedPipelineMode
		End If
		dbrecord(i,5)=ManagedPipelineMode	
		Wscript.Echo "ManagedRuntimeVersion : "& objItem.managedRuntimeVersion
		If Err.Number <> 0 Then
			WScript.Echo "Error in Getting ManagedRuntimeVersion: " & Err.Description
			ManagedRuntimeVersion="No Managed Code"
			Err.Clear
		Else
			If objItem.managedRuntimeVersion="" Then
				ManagedRuntimeVersion="v2.0"
			Else
				ManagedRuntimeVersion=objItem.managedRuntimeVersion
			End If
		End If
		dbrecord(i,6)=ManagedRuntimeVersion
		On Error Goto 0
		i=i+1
	Next
	UpdateDBAppPools strDomain,strComputer,dbrecord,colItems.count
Loop While False
Loop


Sub UpdateDBAppPools(ServerDomain,ServerName,dbrecord,count)
	Logla "apppools", "Connecting to DB"
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 
				On Error Resume Next
				If Err.Number<>0 Then
					Logla "apppools", "DB Connection Failed!!"
					Exit Sub
				End If
				Logla "apppools", "Connected to DB Success!!"
				Err.Clear
				On Error Goto 0
	Logla "apppools", "Connecting AppPools DB"
	i=0
	For i = 0 to count-1
	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.AppPools WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Name""='"&dbrecord(i,0) &"'", objConnection, adOpenStatic, adLockOptimistic
	If objRecordset.EOF Then 
		objRecordset.Close
		Logla "apppools", "Inserting the Record - Cannot Find Record "&dbrecord(i,0)&" on DB"
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.AppPools (""ServerDomain"",""ServerName"",""Name"", ""AppPoolIdentityType"", ""WAMUserName"", ""WAMUserPass"",""PeriodicRestartSchedule"",""ManagedPipelineMode"",""ManagedRuntimeVersion"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&dbrecord(i,0) &"' ,'"&dbrecord(i,1) &"' ,'"&dbrecord(i,2) &"' ,'"&dbrecord(i,3) &"','"& dbrecord(i,4) &"','"& dbrecord(i,5) &"','"& dbrecord(i,6) &"'); " , objConnection, adOpenStatic, adLockOptimistic
		Else
			Logla "apppools", "No Update - The Record "&dbrecord(i,0)&" is already in the Table"
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
