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
Logla "shares","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do
	strComputer=ObjServer.ReadLine
	Logla "shares","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\CIMV2", strUser, strPass,"MS_409","NTLMDomain:" + strDomain)
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "shares", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			Else
				Logla "shares", "Connected"
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406
		On Error Goto 0

	If Not IsObject(objWMIService) Then
		Logla "shares", "Connection Failed!!"
		Exit Do
	End If
'	On Error Goto 0
					  
	On Error Resume Next
	Err.Clear
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Share")
	If Not IsObject(colItems) Then
		Logla "shares", "Win32_Share Failed!!"
		Exit Do
	End If
	Logla "shares", "Win32_Share Success!!"
	For Each objItem in colItems 
		Wscript.Echo "Name: " & objItem.Name
		Wscript.Echo "Caption: " & objItem.Caption 
		Wscript.Echo "Path: " & objItem.Path
		Wscript.Echo "Type: " & objItem.Type
		Wscript.Echo "Description: " & objItem.Description
		If Not IsNull(objItem.Description) Then
			Description=Replace(objItem.Description,"'","")
		End If
		UpdateDBShares strDomain,strComputer, objItem.Name,objItem.Caption,objItem.Path,objItem.Type,Description
	Next
Loop While False
Loop



Sub UpdateDBShares(ServerDomain, ServerName,Name,Caption,Path,Typee,Description)
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 



	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.ServerShares WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Name""='"&Name &"' AND ""Path""='"&Path &"'", objConnection, adOpenStatic, adLockOptimistic
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.ServerShares (""ServerDomain"", ""ServerName"",""Name"",""Caption"",""Path"",""Type"",""Description"") VALUES ('"&ServerDomain&"','"&ServerName&"','"&Name&"','"&Caption&"','"&Path&"','"&Typee&"','"&Description &"'); " , objConnection, adOpenStatic, adLockOptimistic
	End If

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
