includeFile "c:\scripts\variables.config"
includeFile "c:\scripts\logla.vbs"
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjServer = ObjFSO.OpenTextFile("c:\scripts\ServerList.list", ForReading)
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
Set Locator = CreateObject("WbemScripting.SWbemLocator")
		Locator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy
Logla "websites","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do
	strComputer=ObjServer.ReadLine
	Logla "websites","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\microsoftiisv2", strUser, strPass,,"NTLMDomain:" + strDomain)
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "websites", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			Else
				Logla "websites", "Connected"
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406
		On Error Goto 0

	If Not IsObject(objWMIService) Then
		Logla "websites", "Connection Failed!!"
		Exit Do
	End If
'	On Error Goto 0

	Set colItems = objWMIService.ExecQuery( "SELECT * FROM IIsWebServerSetting",,48) 
	Logla "websites", "Getting IIsWebServerSetting"
	If Not IsObject(colItems) Then
		Logla "websites", "IIsWebServerSetting Failed!!"
		Exit Do
	End If
	Logla "websites", "IIsWebServerSetting Success!!"
	For Each objItem in colItems 
		Wscript.Echo "Name: " & objItem.Name
		SplitName=split(objItem.Name,"/")
		ShortName=SplitName(1)
		Wscript.Echo "ShortName:" & ShortName
		Wscript.Echo "ServerDescription: " & objItem.ServerComment 
		Wscript.Echo "AppPoolId : " & objItem.AppPoolId  
		Wscript.Echo "AuthAdvNotifyDisable : " & objItem.AuthAdvNotifyDisable 
		Wscript.Echo "AuthAnonymous  : " & objItem.AuthAnonymous  
		Wscript.Echo "AuthBasic  : " & objItem.AuthBasic  
		Wscript.Echo "AuthChangeDisable  : " & objItem.AuthChangeDisable  
		Wscript.Echo "AuthChangeUnsecure  : " & objItem.AuthChangeUnsecure  
		Wscript.Echo "AuthFlags   : " & objItem.AuthFlags
		If IsNumeric(Ubound(objItem.ServerBindings)) Then
			For i = 0 to Ubound(objItem.ServerBindings)
				Wscript.Echo "ServerBindings HostName: " & _
					objItem.ServerBindings(i).HostName
				Wscript.Echo "ServerBindings IP: " & _
					objItem.ServerBindings(i).IP
				Wscript.Echo "ServerBindings Port: " & _
					objItem.ServerBindings(i).Port
				UpdateIISServerBindings strDomain,strComputer,objItem.ServerComment,ShortName,objItem.ServerBindings(i).HostName,objItem.ServerBindings(i).IP,objItem.ServerBindings(i).Port
			Next
		End If
		If IsNumeric(Ubound(objItem.SecureBindings)) Then
			For i = 0 to Ubound(objItem.SecureBindings)
				Wscript.Echo "SecureBindings IP: " & _
					objItem.SecureBindings(i).IP
				Wscript.Echo "SecureBindings Port: " & _
					objItem.SecureBindings(i).Port
'		Logla "websites", "Updating IIS SecureBinding DB"
				UpdateIISSecureBindings strDomain,strComputer,objItem.ServerComment,ShortName,objItem.SecureBindings(i).IP,objItem.SecureBindings(i).Port
'		Logla "websites", "Updated IIS SecureBinding DB"
			Next
		End If		
		
		If IsNumeric(Ubound(objItem.MimeMap)) Then
			For i = 0 to Ubound(objItem.MimeMap)
				Wscript.Echo "MimeMap MimeType: " & _
					objItem.MimeMap(i).MimeType
				Wscript.Echo "MimeMap Extension: " & _
					objItem.MimeMap(i).Extension
'		Logla "websites", "Updating IIS MimeMap DB"
				UpdateIISMimeMap strDomain,strComputer,objItem.ServerComment,ShortName,objItem.MimeMap(i).MimeType,objItem.MimeMap(i).Extension
'		Logla "websites", "Updated IIS MimeMap DB"
			Next
		End If	

		If IsNumeric(Ubound(objItem.ScriptMaps)) Then		
			For i = 0 to Ubound(objItem.ScriptMaps)
				Wscript.Echo "Extensions: " & objItem.ScriptMaps(i).Extensions
				Wscript.Echo "ScriptProcessor: " & objItem.ScriptMaps(i).ScriptProcessor
				Wscript.Echo "Flags: " & objItem.ScriptMaps(i).Flags
				Wscript.Echo "IncludedVerbs: " & objItem.ScriptMaps(i).IncludedVerbs
'		Logla "websites", "Updating IIS ScriptMaps DB" 
				UpdateIISScriptMaps strDomain,strComputer,objItem.ServerComment,ShortName,objItem.ScriptMaps(i).Extensions,objItem.ScriptMaps(i).ScriptProcessor,objItem.ScriptMaps(i).Flags,objItem.ScriptMaps(i).IncludedVerbs
'		Logla "websites", "Updated IIS ScriptMaps DB"
			Next
		End If
		
		If IsNumeric(Ubound(objItem.HttpCustomHeaders)) Then		
			For i = 0 to Ubound(objItem.HttpCustomHeaders)
				Wscript.Echo "HttpCustomHeaders KeyName: " & _
					objItem.HttpCustomHeaders(i).KeyName 
				Wscript.Echo "HttpCustomHeaders Value: " & _
					objItem.HttpCustomHeaders(i).Value 
'		Logla "websites", "Updating HTTPCustomHeaders DB"
				UpdateHTTPCustomHeaders strDomain,strComputer,objItem.ServerComment,ShortName,objItem.HttpCustomHeaders(i).KeyName,objItem.HttpCustomHeaders(i).Value
'		Logla "websites", "Updated HTTPCustomHeaders DB"
			Next
		End If
    Next
	Logla "websites", "Getting IIsFilterSetting"

	Set colItems = objWMIService.ExecQuery _
	("Select * from IIsFilterSetting")
	If Not IsObject(colItems) Then
		Logla "websites", "IIsFilterSetting Failed!!"
		Exit Do
	End If
	Logla "websites", "IIsFilterSetting Success!!"
	For Each objItem in colItems
		Wscript.Echo "Filter Path: " & objItem.FilterPath
		Wscript.Echo "Filter State: " & objItem.FilterState
		Wscript.Echo "Name: " & objItem.Name
		SplitName=split(objItem.Name,"/")
		if SplitName(1) = "FILTERS" Then
			ShortName = "ROOT"
		else
			ShortName = SplitName(1)
		End if
		Wscript.Echo "ShortName:" & ShortName
'	Logla "websites", "Updating IISISAPIFilters DB"
		IISISAPIFilters strDomain,strComputer,objItem.Name,objItem.FilterPath,objItem.FilterState
'	Logla "websites", "Updated IISISAPIFilters DB"
	Next
	Logla "websites", "Getting IIsWebVirtualDirSetting"
	Set colItems = objWMIService.ExecQuery( "SELECT * FROM IIsWebVirtualDirSetting",,48) 
	If Not IsObject(colItems) Then
		Logla "websites", "IIsWebVirtualDirSetting Failed!!"
		Exit Do
	End If
	Logla "websites", "IIsWebVirtualDirSetting Success!!"
	For Each objItem in colItems 
		Wscript.Echo "Name: " & objItem.Name
		SplitName=split(objItem.Name,"/")
		Wscript.Echo "ShortName:" & SplitName(1)
		Wscript.Echo "Physical Path: " & objItem.Path 	
		Wscript.Echo "AppPoolId : " & objItem.AppPoolId 
'	Logla "websites", "Updating IISPaths DB"
		IISPaths strDomain,strComputer,SplitName(1),objItem.Name,objItem.Path,objItem.AppPoolId
'	Logla "websites", "Updated IISPaths DB"
	Next
Loop While False
Loop



Sub UpdateIISScriptMaps(ServerDomain,ServerName,ServerComment,ShortName,Extensions,ScriptProcessor,Flags,IncludedMaps)
	'On Error Resume Next
	On Error Goto 0
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 
	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.ScriptMaps WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""WebSiteId""='"&ShortName &"' AND ""Extensions""='"&Extensions &"'", objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Connect to DB."
			End if
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.ScriptMaps (""ServerDomain"",""ServerName"",""ServerComment"",""WebSiteId"", ""Extensions"", ""ScriptProcessor"", ""Flags"",""IncludedMaps"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&ServerComment &"' ,'"&ShortName &"' ,'"&Extensions &"' ,'"&ScriptProcessor &"' ,'"&Flags &"','"& IncludedMaps &"'); " , objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Insert into DB."
			End if
	End If
objConnection.Close	
On Error Goto 0

End Sub

Sub UpdateIISMimeMap(ServerDomain,ServerName,ServerComment,ShortName,MimeType,Extension)
	'On Error Resume Next
	On Error Goto 0
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 
	
	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.MimeTypes WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""WebSiteId""='"&ShortName &"' AND ""MimeType""='"&MimeType &"' AND ""Extension""='"&Extension &"'", objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Connect to DB."
			End if
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.MimeTypes (""ServerDomain"",""ServerName"",""ServerComment"",""WebSiteId"",""MimeType"" ,""Extension"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&ServerComment &"' ,'"&ShortName &"' ,'"&MimeType &"' ,'"&Extension &"'); " , objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Insert into DB."
			End if
	End If
objConnection.Close	
On Error Goto 0
End Sub

Sub UpdateIISSecureBindings(ServerDomain,ServerName,ServerComment,ShortName,IP,Port)
	'On Error Resume Next
	On Error Goto 0
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWROD;" 

	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.SecureBindings WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""WebSiteId""='"&ShortName &"' AND ""IP""='"&IP &"' AND ""Port""='"&Port &"'", objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Connect to DB."
			End if
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.SecureBindings (""ServerDomain"",""ServerName"",""ServerComment"",""WebSiteId"",""IP"" ,""Port"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&ServerComment &"' ,'"&ShortName &"' ,'"&IP &"' ,'"&Port &"'); " , objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Insert into DB."
			End if
	End If
objConnection.Close	
On Error Goto 0
End Sub

Sub UpdateIISServerBindings(ServerDomain,ServerName,ServerComment,ShortName,Hostname,IP,Port)
	'On Error Resume Next
	On Error Goto 0
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 

	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.ServerBindings WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""WebSiteId""='"&ShortName &"' AND ""IP""='"&IP &"' AND ""Port""='"&Port &"'", objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Connect to DB."
			End if
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.ServerBindings (""ServerDomain"",""ServerName"",""ServerComment"",""WebSiteId"",""Hostname"",""IP"" ,""Port"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&ServerComment &"' ,'"&ShortName &"' ,'"& Hostname &"' ,'"&IP &"' ,'"&Port &"'); " , objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Insert into DB."
			End if
	End If
objConnection.Close	
On Error Goto 0
End Sub


Sub UpdateHTTPCustomHeaders(ServerDomain,ServerName,ServerComment,ShortName,KeyName,Value)
	'On Error Resume Next
	On Error Goto 0
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 

	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.HTTPCustomHeaders WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""WebSiteId""='"&ShortName &"' AND ""KeyName""='"&KeyName &"' AND ""Value""='"&Value &"'", objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Connect to DB."
			End if
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.HTTPCustomHeaders (""ServerDomain"",""ServerName"",""ServerComment"",""WebSiteId"",""KeyName"" ,""Value"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&ServerComment &"' ,'"&ShortName &"' ,'"&KeyName &"' ,'"&Value &"'); " , objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Insert into DB."
			End if
	End If
objConnection.Close	
On Error Goto 0
End Sub

Sub IISISAPIFilters(ServerDomain,ServerName,ShortName,FilterPath,FilterState)
	'On Error Resume Next
	On Error Goto 0
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 

	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.ISAPIFilters WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""WebSiteId""='"&ShortName &"' AND ""FilterPath""='"&FilterPath &"'", objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Connect to DB."
			End if
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.ISAPIFilters (""ServerDomain"",""ServerName"",""WebSiteId"",""FilterPath"" ,""FilterState"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&ShortName &"' ,'"&FilterPath &"' ,'"&FilterState &"'); " , objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Insert into DB."
			End if
	End If
objConnection.Close	
On Error Goto 0
End Sub

Sub IISPaths(ServerDomain,ServerName,ShortName,FullName, Path,PoolId)
	'On Error Resume Next
	On Error Goto 0
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 

	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.IISPaths WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""WebSiteId""='"&ShortName &"' AND ""Path""='"&Path &"'", objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Connect to DB."
			End if
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.IISPaths (""ServerDomain"",""ServerName"",""WebSiteId"",""FullName"",""Path"" ,""PoolId"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&ShortName &"' ,'"&FullName &"' ,'"&Path &"' ,'"&PoolId &"'); " , objConnection, adOpenStatic, adLockOptimistic
			If Err.Number <> 0 then
				logla "websites","Unable to Insert into DB."
			End if
	End If
objConnection.Close	
On Error Goto 0
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


