
Public Function SSMS_Running_YN(process)

	Dim my_Obj
	Dim procs
	Set my_Obj = GetObject("winmgmts:")
	Set procs = my_Obj.ExecQuery("select * from win32_process where name='" & process & "'")

	If procs.Count > 0 Then
		SSMS_Running_YN = True
	Else
		SSMS_Running_YN = False
	End If

	Set my_Obj = Nothing
	Set procs = Nothing

End Function


Function Nz(inputVal, defaultVal)
  If inputVal Is Nothing Then
    Nz = defaultVal
  Else
    Nz = inputVal
  End If

End Function


Public Sub Launch_SSMS()
On Error Resume Next
''' Checks to see if SSMS running; If not, will launch and run query
	Dim launch_utility
	Dim my_server
	Dim qry_str
	Dim my_db
	Dim return_val
	Dim my_obj
  
  ''' Need to set server and database variables before running.
	my_server = ""
	my_db = ""
  
  ''' General name of SQL Server Management Studio winow
	ssms_name = "Microsoft SQL Server Management Studio"

  ''' Command line script to launch SSMS; Skips initial splash screen if server and
  ''' database are defined
  
  If Len(Nz(my_server, "")) > 0 and Len(my_db) > 0 Then
	  launch_utility = "Ssms -E -S " & my_server & " -d " & my_db & " -nosplash"
  End If


	qry_str = "SELECT TOP 1000 *" & "{ENTER}"
	qry_str = qry_str & "FROM [master].[INFORMATION_SCHEMA].[TABLES]"

	'return_val = Shell(launch_utility, 1)
	Set my_obj = WScript.CreateObject("WScript.Shell")
	my_obj.Run launch_utility
	WScript.Sleep 3000
	my_obj.AppActivate ssms_name
	WScript.Sleep 1000
	my_obj.SendKeys "+{ESC}"
	WScript.Sleep 1000
	my_obj.SendKeys qry_str
	WScript.Sleep 3000
	my_obj.SendKeys "{F5}"

	Set my_obj = Nothing

If Err.Number <> 0 Then
  WScript.Echo "Error in Launch_SSMS: " & Err.Description
  Err.Clear
End If
End Sub


Public Sub Run_Sample_Query()
On Error Resume Next
	Dim qry_str
	Dim ssms_name
	Dim sql_server_proc
	Dim my_obj

	ssms_name = "Microsoft SQL Server Management Studio"
	sql_server_proc = "Ssms.exe"


	qry_str = "SELECT TOP 1000 *" & "{ENTER}"
	qry_str = qry_str & "FROM [master].[INFORMATION_SCHEMA].[TABLES]"

	If SSMS_Running_YN(sql_server_proc) = 0 Then
		Call Launch_SSMS
	Else
		Set my_obj = WScript.CreateObject("WScript.Shell")
		my_obj.AppActivate ssms_name
		WScript.Sleep 1000
		my_obj.SendKeys "+{ESC}"
		WScript.Sleep 1000
		my_obj.SendKeys "^n"
		WScript.Sleep 2000
		my_obj.SendKeys qry_str
		WScript.Sleep 1000
		my_obj.SendKeys "{F5}"
	End If

	'my_obj.SendKeys "{NUMLOCK}"
If Err.Number <> 0 Then
  WScript.Echo "Error in Run_Sample_Query: " & Err.Description
  Err.Clear
End If

End Sub

Call Run_Sample_Query
