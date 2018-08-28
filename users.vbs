includeFile "c:\scripts\variables.config"
includeFile "c:\scripts\logla.vbs"
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjServer = ObjFSO.OpenTextFile("c:\scripts\ServerList.list", ForReading)
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
Set Locator = CreateObject("WbemScripting.SWbemLocator")
		Locator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy	 

     Wscript.Echo "---------------------------------------------------------------"
     Wscript.Echo "----------USERS------------------------------------------------"
     Wscript.Echo "---------------------------------------------------------------"	 
		
Logla "users","************STARTING***************"
Do Until ObjServer.AtEndOfStream
Do

	strComputer=ObjServer.ReadLine

	Logla "users","Connecting to " & strDomain&"-"&strComputer
	Retry=0
	Do
		On Error Resume Next
		Err.Clear
		Set objWMIService = Locator.ConnectServer _
			(strComputer, "root\CIMV2", strUser, strPass,"MS_409","NTLMDomain:" + strDomain)		
			Retry=Retry + 1
			If Err.Number <> 0 or Err.Number<>-2147217406 then
				Wscript.Echo "Unable to connect to "&strComputer&". Retrying connection in 3 seconds..."
				Logla "users", "Unable to connect to Remote Server. Retrying connection in 3 seconds..."
				Wscript.Sleep (3000)
			Else
				Logla "users", "Connected"
			End if
	Loop until Err.Number=0	 or Retry = RetryMax or Err.Number=-2147217406
		On Error Goto 0

	If Not IsObject(objWMIService) Then
		Logla "users", "Connection Failed!!"
		Exit Do
	End If
	On Error Goto 0

	Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount ") 
	Logla "users", "Getting Win32_UserAccount"
	If Not IsObject(colItems) Then
		Logla "users", "Win32_UserAccount Failed!!"
		Exit Do
	End If
	Logla "users", "Win32_UserAccount Success!!"

For Each objItem in colItems 
'    Wscript.Echo "Account Type: " & objItem.AccountType 
    Wscript.Echo "Caption: " & objItem.Caption 
    Wscript.Echo "Description: " & objItem.Description 
    Wscript.Echo "Disabled: " & objItem.Disabled 
    Wscript.Echo "Domain: " & objItem.Domain 
'    Wscript.Echo "Full Name: " & objItem.FullName 
    Wscript.Echo "Local Account: " & objItem.LocalAccount 
    Wscript.Echo "Lockout: " & objItem.Lockout 
    Wscript.Echo "Name: " & objItem.Name 
    Wscript.Echo "Password Changeable: " & objItem.PasswordChangeable 
    Wscript.Echo "Password Expires: " & objItem.PasswordExpires 
    Wscript.Echo "Password Required: " & objItem.PasswordRequired 
    Wscript.Echo "SID: " & objItem.SID 
    Wscript.Echo "SID Type: " & objItem.SIDType 
    Wscript.Echo "Status: " & objItem.Status 
    Wscript.Echo
	If Not IsNull(objItem.Description) Then
	
	Description=Replace(objItem.Description,"'","")
	End If
	
	UpdateDBUsers strDomain,strComputer,objItem.Caption,Description,objItem.Disabled,objItem.Domain,objItem.LocalAccount,objItem.Lockout,objItem.Name,objItem.PasswordChangeable,objItem.PasswordExpires,objItem.PasswordRequired,objItem.SID,objItem.SIDType,objItem.Status 
Next 

     Wscript.Echo "---------------------------------------------------------------"
     Wscript.Echo "----------GROUPS-----------------------------------------------"
     Wscript.Echo "---------------------------------------------------------------"	 
	 
	 
	 
Set colItems = objWMIService.ExecQuery ("Select * from Win32_Group") 
	Logla "users", "Getting Win32_Group"
	If Not IsObject(colItems) Then
		Logla "users", "Win32_Group Failed!!"
		Exit Do
	End If
	Logla "users", "Win32_Group Success!!"
 
For Each objItem in colItems 
    Wscript.Echo "Caption: " & objItem.Caption 
    Wscript.Echo "Description: " & objItem.Description 
    Wscript.Echo "Domain: " & objItem.Domain 
    Wscript.Echo "Local Account: " & objItem.LocalAccount 
    Wscript.Echo "Name: " & objItem.Name 
    Wscript.Echo "SID: " & objItem.SID 
    Wscript.Echo "SID Type: " & objItem.SIDType 
    Wscript.Echo "Status: " & objItem.Status 
    'Wscript.Echo strDomain,strComputer,objItem.Caption,objItem.Description,objItem.Domain,objItem.LocalAccount,objItem.Name,objItem.SID,objItem.SIDType,objItem.Status 
	If Not IsNull(objItem.Description) Then
	
	Description=Replace(objItem.Description,"'","")
	End If


	UpdateDBGroups strDomain,strComputer,objItem.Caption,Description,objItem.Domain,objItem.LocalAccount,objItem.Name,objItem.SID,objItem.SIDType,objItem.Status 
Next 

Loop While False
Loop

Sub UpdateDBUsers(ServerDomain,ServerName,Caption,Description,Disabled,Domain,LocalAccount,Lockout,Name,PasswordChangeable,PasswordExpires,PasswordRequired,SID,SIDType,Status)
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 
	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.ServerUsers WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Caption""='"&Caption &"' AND ""Description""='"&Description &"' AND ""Disabled""='"&Disabled &"' AND ""Domain""='"&Domain &"' AND ""LocalAccount""='"&LocalAccount &"' AND ""Lockout""='"&Lockout &"' AND ""Name""='"&Name &"' AND ""PasswordChangeable""='"&PasswordChangeable &"' AND ""PasswordExpires""='"&PasswordExpires &"' AND ""PasswordRequired""='"&PasswordRequired &"' AND ""SID""='"&SID &"' AND ""SIDType""='"&SIDType &"' AND ""Status""='"&Status &"'", objConnection, adOpenStatic, adLockOptimistic
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.ServerUsers (""ServerDomain"",""ServerName"",""Caption"", ""Description"", ""Disabled"",""Domain"", ""LocalAccount"",""Lockout"",""Name"",""PasswordChangeable"",""PasswordExpires"",""PasswordRequired"", ""SID"", ""SIDType"", ""Status"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&Caption &"' ,'"&Description &"' ,'"&Disabled &"','"&Domain &"' ,'"&LocalAccount &"','"&Lockout &"' ,'"&Name &"','"&PasswordChangeable &"','"&PasswordExpires &"','"&PasswordRequired &"','"&SID &"' ,'"&SIDType &"' ,'"&Status &"'); " , objConnection, adOpenStatic, adLockOptimistic
	End If
objConnection.Close	

End Sub






Sub UpdateDBGroups(ServerDomain,ServerName,Caption,Description,Domain,LocalAccount,Name,SID,SIDType,Status)
	objConnection.Open _ 
		"Provider=SQLOLEDB;Data Source=DATABASENAME;" & _ 
			"Initial Catalog=DATABASE;" & _ 
				"User ID=USERNAME;Password=PASSWORD;" 

	objRecordset.Open "SELECT * FROM Win2016Mig.dbo.ServerGroups WHERE ""ServerDomain""='"&ServerDomain &"' AND ""ServerName""='"&ServerName &"' AND ""Caption""='"&Caption &"' AND ""Description""='"&Description &"' AND ""Domain""='"&Domain &"' AND ""LocalAccount""='"&LocalAccount &"' AND ""Name""='"&Name &"' AND ""SID""='"&SID &"' AND ""SIDType""='"&SIDType &"' AND ""Status""='"&Status &"'", objConnection, adOpenStatic, adLockOptimistic
	If objRecordset.EOF Then 
		objRecordset.Close
		ObjRecordset.Open "INSERT INTO Win2016Mig.dbo.ServerGroups (""ServerDomain"",""ServerName"",""Caption"", ""Description"", ""Domain"", ""LocalAccount"",""Name"", ""SID"", ""SIDType"", ""Status"") VALUES ('"&ServerDomain &"' ,'"&ServerName &"' ,'"&Caption &"' ,'"&Description &"' ,'"&Domain &"' ,'"&LocalAccount &"','"&Name &"','"&SID &"' ,'"&SIDType &"' ,'"&Status &"'); " , objConnection, adOpenStatic, adLockOptimistic
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


 