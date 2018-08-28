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
Logla "environment","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do
	strComputer=ObjServer.ReadLine
	Logla "environment","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\CIMV2", strUser, strPass,"MS_409","NTLMDomain:" + strDomain)
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "environment", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			Else
				Logla "environment", "Connected"
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406
		On Error Goto 0

	If Not IsObject(objWMIService) Then
		Logla "environment", "Connection Failed!!"
		Exit Do
	End If
'	On Error Goto 0
					  
	On Error Resume Next
	Err.Clear
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment") 
	If Not IsObject(colItems) Then
		Logla "environment", "Win32_Environment Failed!!"
		Exit Do
	End If
	Logla "environment", "Win32_Environment Success!!"
 
  	For Each objItem in colItems 
		Wscript.Echo "Caption: " & objItem.Caption
		Wscript.Echo "Description: " & objItem.Description 
		Wscript.Echo "Name: " & objItem.Name 
		Wscript.Echo "System Variable: " & objItem.SystemVariable 
		Wscript.Echo "User Name: " & objItem.UserName 
		Wscript.Echo "Variable Value: " & objItem.VariableValue 
		If Not IsNull(objItem.Description) Then
			Description=Replace(objItem.Description,"'","")
		End If
		UpdateDBEnvironment strDomain,strComputer, objItem.Caption,Description,objItem.Name ,objItem.SystemVariable ,objItem.UserName,objItem.VariableValue 

		If trim(objItem.Name)="Path" Then
			PathArr=split(objItem.VariableValue,";")
			ReDim dbrecord(UBound(PathArr)-1,0)
			i=0
			For Each PathNew in PathArr 
				dbrecord(i,0)=PathNew
				i=i+1
			Next
			UpdateDBPath strDomain,strComputer,dbrecord,UBound(PathArr)-1
		End If
	Next 

Loop While False
Loop


Sub UpdateDBEnvironment(ServerDomain, ServerName,Caption,Description,Name,SystemVariable,UserName,VariableValue)
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=SERVERNAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=UID;Password=PWD;" 

	objRecordset.Open "SELECT * FROM DATABASE WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Caption""='"&Caption &"' AND ""Description""='"&Description &"' AND ""Name""='"&Name &"' AND ""UserName""='"&UserName &"' AND""VariableValue""='"&VariableValue &"'", objConnection, adOpenStatic, adLockOptimistic
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO DATABASE (""ServerDomain"", ""ServerName"",""Caption"",""Description"",""Name"",""SystemVariable"",""UserName"",""VariableValue"") VALUES ('"&ServerDomain&"','"&ServerName&"','"&Caption&"','"&Description&"','"&Name&"','"&SystemVariable&"','"&UserName &"','"&VariableValue &"'); " , objConnection, adOpenStatic, adLockOptimistic
	End If

objConnection.Close 
End Sub



Sub UpdateDBPath(ServerDomain, ServerName,dbrecord,count)
	Logla "environment", "Connecting to DB"
	
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=SERVERNAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=UID;Password=PWD;" 
				On Error Resume Next
				If Err.Number<>0 Then
					Logla "environment", "DB Connection Failed!!"
					Exit Sub
				End If
				Logla "environment", "Connected to DB Success!!"
				Err.Clear
				On Error Goto 0
	Logla "environment", "Connecting InstalledPrograms DB"
	i=0
	For i = 0 to count
	objRecordset.Open "SELECT * FROM DATABASE WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Path""='"&dbrecord(i,0) &"'", objConnection, adOpenStatic, adLockOptimistic	
		If objRecordset.EOF Then 
			objRecordset.Close
			Logla "environment", "Inserting the Record - Cannot Find Path "&dbrecord(i,0)&" on DB"
			ObjRecordset.Open "INSERT INTO DATABASE (""ServerDomain"", ""ServerName"",""Path"") VALUES ('"&ServerDomain&"','"&ServerName&"','"&dbrecord(i,0)&"'); " , objConnection, adOpenStatic, adLockOptimistic
		Else
			Logla "environment", "No Update - The Record "&dbrecord(i,0)&" is already in the Table"
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


