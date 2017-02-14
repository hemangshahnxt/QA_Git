'Window("Nextech Main Window").Window("Letter Writing").ActiveX("Merging Panel").Click 372,92
Dim TestDBName : TestDBName = Environment.Value("DBUnderTest")

Dim IsPresent
Dim StartUpMethod
StartUpMethod = Environment.Value("StartUpMethod")
'Log in to practice

RunAction "LogIn [GlobalPracticeActions]", oneIteration, TestDBName
If Dialog("regexpwndtitle:= Log In").WinButton("regexpwndtitle:=&Log In").Exist Then
	Dialog("regexpwndtitle:= Log In").WinButton("regexpwndtitle:=&Log In").Click
End If
wait 5
'Launching PRACTICE application

Window("regexpwndtitle:=Nextech.*").WinMenu("menuobjtype:=2").Select "Modules;Reports"
If Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Reports").Exist Then
	Reporter.ReportEvent micPass, "Reports Menu is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Reports Menu is not Present"
End If

Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Patients")
Call VerTabOpened ("Patients")
Call VerifyPatients()

'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Contacts")
'Call VerTabOpened ("Contacts")
'Call VerifyContacts()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Marketing")
'Call VerTabOpened ("Marketing")
'Call VerifyMarketing()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Inventory")
'Call VerTabOpened ("Inventory")
'Call VerifyInventory()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Scheduler")
'Call VerTabOpened ("Scheduler")
'Call VerifyScheduler()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Charges")
'Call VerTabOpened ("Charges")
'Call VerifyCharges()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Payments")
'Call VerTabOpened ("Payments")
'Call VerifyPayments()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Financial")
'Call VerTabOpened ("Financial")
'Call VerifyFinancial()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "ASC")
'Call VerTabOpened ("ASC")
'Call VerifyASC()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Administration")
'Call VerTabOpened ("Administration")
'Call VerifyAdministration()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Other")
'Call VerTabOpened ("Other")
'Call VerifyOther()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Prac. Analysis")
'Call VerTabOpened ("Prac. Analysis")
'Call VerifyAnalysis()
'
'Call ChangeTabs (Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").ActiveX("acx_name:=NxTab Control"), "Optical")
'Call VerTabOpened ("Optical")
'Call VerifyOptical()
'
Window("regexpwndtitle:=Nextech.*","index:=0").WinMenu("menuobjtype:=2").Select "Modules;Patients"
RunAction "LogOut [GlobalPracticeActions]", oneIteration

Sub VerifyPatients()
	

Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Patients", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet2" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Patients1", "Sheet2" 
'
RM = Datatable.GetSheet("Sheet2").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RM Step 1
	Datatable.GetSheet("Sheet2").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet2")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet2")
	Inde = 	Datatable.Value("Inde","Sheet2")
	
	noOfReports = Datatable.Value("noOfReports","Sheet2")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
'If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Message Box").Exist(2) Then
'		Reporter.ReportEvent micPass, "Check Point Present", "Verify Database Dialog is Present"
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Message Box").WinButton("regexpwndtitle:=OK").Click
'
'End if

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
	Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If


'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
'		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Message Box").Exist(1) Then
'			Reporter.ReportEvent micPass, "Check Point Present", "Verify Database Dialog is Present"
'			Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Message Box").WinButton("regexpwndtitle:=OK").Click
'		End If
	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
		Reporter.ReportEvent micPass, "Database Verified",""
	Else
		Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
	End If
	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
	While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
	Wend	
	End If
	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click


Next
End Sub

Sub VerifyContacts()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Contacts", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)
End Sub

Sub VerifyMarketing()
	
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Marketing1", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet2" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Marketing2", "Sheet2" 
'
RM = Datatable.GetSheet("Sheet2").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RM Step 1
	Datatable.GetSheet("Sheet2").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet2")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet2")
	Inde = 	Datatable.Value("Inde","Sheet2")
	
	noOfReports = Datatable.Value("noOfReports","Sheet2")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If

'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If
	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click


Next
End Sub

Sub VerifyInventory()
Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Inventory", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)
End Sub

Sub VerifyScheduler()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Scheduler1", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet2" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Scheduler2", "Sheet2" 
'
RM = Datatable.GetSheet("Sheet2").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RM Step 1
	Datatable.GetSheet("Sheet2").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet2")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet2")
	Inde = 	Datatable.Value("Inde","Sheet2")
	
	noOfReports = Datatable.Value("noOfReports","Sheet2")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If

'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If
	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click


Next
	
	
End Sub

Sub VerifyCharges()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Charges1", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet2" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Charges2", "Sheet2" 
'
RM = Datatable.GetSheet("Sheet2").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RM Step 1
	Datatable.GetSheet("Sheet2").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet2")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet2")
	Inde = 	Datatable.Value("Inde","Sheet2")
	
	noOfReports = Datatable.Value("noOfReports","Sheet2")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If

'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If
	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click


Next
End Sub

Sub VerifyPayments()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Payments1", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet2" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Payments2", "Sheet2" 
'
RM = Datatable.GetSheet("Sheet2").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RM Step 1
	Datatable.GetSheet("Sheet2").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet2")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet2")
	Inde = 	Datatable.Value("Inde","Sheet2")
	
	noOfReports = Datatable.Value("noOfReports","Sheet2")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If


'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If
	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click


Next
End Sub

Sub VerifyFinancial()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Financial1", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").Click 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet2" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Financial2", "Sheet2" 
'
RM = Datatable.GetSheet("Sheet2").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RM Step 1
	Datatable.GetSheet("Sheet2").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet2")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet2")
	Inde = 	Datatable.Value("Inde","Sheet2")
	
	noOfReports = Datatable.Value("noOfReports","Sheet2")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If

'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
	While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
	Wend	
End If
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click

Next

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet3" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Financial3", "Sheet3" 
'
RF = Datatable.GetSheet("Sheet3").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RF Step 1
	Datatable.GetSheet("Sheet3").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet3")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet3")
	Inde = 	Datatable.Value("Inde","Sheet3")
	
	noOfReports = Datatable.Value("noOfReports","Sheet3")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If

'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If
	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click

Next

End Sub

Sub VerifyASC()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","ASC", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)
End Sub

Sub VerifyAdministration()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Administration1", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick 325,500
''===================================================================================
'
Datatable.AddSheet "Sheet2" 'Adding a new sheet to the Datatable
'
''Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Administration2", "Sheet2" 
'
RM = Datatable.GetSheet("Sheet2").GetRowCount

'Call VerifyReports(RC)
For i = 1 To RM Step 1
	Datatable.GetSheet("Sheet2").SetCurrentRow(i)
	
	Reports = Datatable.Value("Reports","Sheet2")
	'SubReports = Datatable.Value("SubReports","Sheet2")
	Index = Datatable.Value("Index","Sheet2")
	Inde = 	Datatable.Value("Inde","Sheet2")
	
	noOfReports = Datatable.Value("noOfReports","Sheet2")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1

'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index

If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If

'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Message Box").Exist(2) Then
		Reporter.ReportEvent micPass, "Check Point Present", "Verify Database Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Message Box").WinButton("regexpwndtitle:=OK").Click

End if
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click

'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If
	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click


Next
End Sub

Sub VerifyOther()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Other", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)
End Sub

Sub VerifyAnalysis()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Analysis", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)
End Sub

Sub VerifyOptical()
	Datatable.AddSheet "Sheet1" 'Adding a new sheet to the Datatable

'Excel spreadsheet path to import content from excel
Datatable.ImportSheet "C:\Code\trunk\QA\UFT\Shared\Reports.xlsx","Optical", "Sheet1" 

RC = Datatable.GetSheet("Sheet1").GetRowCount

Call VerifyReports(RC)
End Sub

Function generateReports(noOfReports)
	
	'	yCo(yArrray)
	'	If noOfReports > 1 Then
	'		For i = 0 To yArray Step 1
	'			colName = "Index" & i
	'			yCo(i) = Datatable.Value(colName,"Sheet1")
	'		Next
	'	Else 
	'		yCo(yArray) = Datatable.Value("Index0","Sheet1")
	'	End if
	dim yArrray(13)	' = {"10,25,40,55,70}
	yArrray(0) =10
	yArrray(1) =25
	yArrray(2) =40
	yArrray(3) =55
	yArrray(4) =70
	yArrray(5) =85
	yArrray(6) =100
	yArrray(7) =115
	yArrray(8) =130
	yArrray(9) =145
	yArrray(10) =160
	yArrray(11) =175
	yArrray(12) =190
	yArrray(13) =205
	
'	yArrray = noOfReports - 1
End Function

Function GetColumnIndex(dl, name)
	dim c 
	c = dl.ColumnCount
	For i = 0 To c-1
		dim col
		set col = dl.GetColumn(i)
		If (col.ColumnTitle = name) Then
			GetColumnIndex = i
			exit function 
		End If
	Next
	GetColumnIndex = -1
	
End Function

Function CellValueToStr(Value)
    If IsNull(Value) Then
        CellValueToStr = ""
    Else
        CellValueToStr = CStr(Value)
    End If
End Function

Sub ChangeTabs(tabsTO, label)
	curTabIndex = tabsTO.Object.CurSel
	newTabIndex = GetTabIndex(tabsTO.Object, label)
	tabsTO.Object.CurSel  = newTabIndex
	Call tabsTO.FireEvent("SelectTab", newTabIndex, curTabIndex)
End Sub

Function GetPrefixIndex(t, preval)
	count = t.Size
	For i = 0 To count - 1 Step 1
		If Replace(t.Label(i), "&", "") = preval Then
			GetPrefixIndex = i

			Exit function
		End If
	Next
	GetPrefixIndex = -1
End Function

Sub ChangePrefix(prefsTO, preval)
	curPrefixIndex = prefsTO.Object.CurSel
'	newTabIndex = GetTabIndex(tabsTO.Object, "Notes")
	newPrefIndex = GetPrefixIndex(prefsTO.Object, preval)
	prefsTO.Object.CurSel  = newPrefixIndex
	Call prefsTO.FireEvent("SelectTab", newPrefixIndex, curPrefixIndex)
End Sub	

Sub GetExpectedValues()
		'Get the expected results for that row in the spreadsheet
	ExpRowBillIDStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(1).Value)
	ExpDateStr= CellValueToStr(xlSheetStr.Rows(RowInx).Columns(2).Value)
	ExpInputDateStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(3).Value)
	If ExpInputDateStr = "<TD>" Then
		ExpInputDateStr = TheCurrentDateStr
	End If
	
End Sub

Sub VerTabOpened (TabName)
	Select Case TabName
		Case "Patients"
		
		Case "Contacts"
		
		Case "Marketing"
		
		Case "Inventory"
		
		Case "Scheduler"
		
		Case "Charges"
		
		Case "Payments"
		
		Case "Financial"
		
		Case "ASC"
		
		Case "Administration"
		
		Case "Other"
		
		Case "Prac. Analysis"
		
		Case "Optical"
		
	End Select
End Sub

Function VerifyReports(RC)
	

For i = 1 To RC Step 1
	Datatable.GetSheet("Sheet1").SetCurrentRow(i)
'	Main = Datatable.Value("Window","Sheet1")
'	Module = Datatable.Value("Module","Sheet1")
'	Tab = Datatable.Value("Tab","Sheet1")
	Reports = Datatable.Value("Reports","Sheet1")
	'SubReports = Datatable.Value("SubReports","Sheet1")
	Index = Datatable.Value("Index","Sheet1")
	Inde = 	Datatable.Value("Inde","Sheet1")
	'Index1 = Datatable.Value("index1","Sheet1")
	'Index2 = Datatable.Value("index2","Sheet1")
	noOfReports = Datatable.Value("noOfReports","Sheet1")
	Call generateReports(noOfReports)
'	yArrray = noOfReports - 1
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing
'Window("regexpwndtitle:=Nextech \(UFT12200DB\)").Window("regexpwndtitle:=Reports").WinObject("attached text:=Available Reports","nativeclass:=Afx:42800000:8b").Set Reports


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick Inde, Index
wait (.5)


If Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report","enabled:=True").Exist Then
	Reporter.ReportEvent micPass, "Passed", "Report is available to be verified"
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click
Else
	
	Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinObject("nativeclass:=Afx:42800000:8b", "index:=1").DblClick inde,index
End If

Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndtitle:=Edit Report").Click

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Pick the Report to Edit is Present"
ELse
	Reporter.ReportEvent micFail, "Test Failed", "Pick the Report to Edit is not Present"
End If



'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick inde,noOfReports

'Clicking on the New Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Exist(5) Then
	Reporter.ReportEvent micPass, "Verified", "Verify Button is Present"
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If

'Closing the displayed report
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
'Set abc = CreateObject("Mercury.DeviceReplay")
'	abc.KeyDown 56
'	abc.KeyDown 62
'	wait 1
'	abc.Keyup 56
'	abc.Keyup 62
'Set abc = Nothing	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
	Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click
	
Else
'	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").Exist 
'	Reporter.ReportEvent micFail, "No more Sub Forms", "Nextech Practice Dialog is Present"
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist(1) Then
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
End If
	
For ii = 1 To noOfReports
		Datatable.GetSheet("Sheet1").SetCurrentRow(ii)
		Inde = 	Datatable.Value("Inde","Sheet1")
		Index = Datatable.Value("Index","Sheet1")
		noOfReports = Datatable.Value("noOfReports","Sheet1")
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinObject("nativeclass:=Afx:42800000:8b").DblClick Inde,Index	
		Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=New").Click
		
		If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Exist Then
			Reporter.ReportEvent micPass, "Check Point is Present", "Reports Dialog is Present"
		Else
			Reporter.ReportEvent micFail, "Check Point Failed", "Test Has FAILED!!!!"
		End If
		
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).WinButton("regexpwndtitle:=Verify").Click
	
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").Static("regexpwndtitle:=The database is up to date\.").Exist Then
	Reporter.ReportEvent micPass, "Database Verified",""
Else
	Reporter.ReportEvent micFail, "ERROR ERROR ERROR", "Needs Attention"
End If

If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1) Then
While Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Exist(1)
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Verify Database").WinButton("regexpwndtitle:=OK").Click	
Wend	
End If
wait(0.5)	
	'Closing the displayed report
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Close
wait(0.5)
'	Set abc = CreateObject("Mercury.DeviceReplay")
'		abc.KeyDown 56
'		abc.KeyDown 62
'		wait 1
'		abc.Keyup 56
'		abc.Keyup 62
'	Set abc = Nothing	
	If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").Exist Then
		Reporter.ReportEvent micPass, "Verified", "Save Custom Reports Dialog is Present"
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Save Custom Reports Dialog is not Present"
	End If
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").Dialog("regexpwndtitle:="&Reports).Dialog("regexpwndtitle:=Save Custom Report").WinButton("regexpwndtitle:=No").Click
	
Next
	Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Pick the Report to Edit").WinButton("regexpwndtitle:=Close").Click	
End If


Window("regexpwndtitle:=Nextech.*","index:=0").Window("regexpwndtitle:=Reports").WinButton("regexpwndclass:=Button","index:=2").Click


Next
End Function







