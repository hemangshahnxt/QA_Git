'On Error Resume Next
Dim ExpectedEqualsActual
Dim TestResult
Dim InName
Dim logmsg
Dim IsPresent
Dim TheCurrentDateStr
Dim NumPassedInt, NumFailedInt
Dim DBUnderTest
NumPassedInt = 0
NumFailedInt = 0
Dim MsgTextStr
Dim StartUpMethod

'On Error Resume Next
'Msgbox Environment.Value("DBUnderTest")

DBUnderTest = Environment.Value("DBName")
StartUpMethod = Environment.Value("StartUpMethod")

TheCurrentDateStr = CStr(date)

'Log in to practice


RunAction "LogIn [GlobalPracticeActions]", oneIteration, DBUnderTest


'Verify Practice opens and select the Patients Module
'Verify Practice opens and select the Patients Module

'Window("Nextech Main Window").Check CheckPoint("Nextech (UFT12100DB) Patients") @@ hightlight id_;_5052716_;_script infofile_;_ZIP::ssf3.xml_;_

If Window("Nextech Main Window").Exist (60) Then
	Call RecordResults(True, "Valid Nextech instance")
Else
	Call RecordResults(False, "Invalid Nextech instance")
End If

'Window("Nextech Main Window").WinToolbar("ModuleButtons").Press 1 @@ hightlight id_;_986666_;_script infofile_;_ZIP::ssf2.xml_;_

'***** Sandbox Play area for testing function and sub code.  Set a breakpoint on the first call to ChangeTabs, insert the statement(s) below and run from line to see if they work
'MsgBox (Dialog("text:=Nextech Practice").Static("text:=In the following window, please select your code file to import.").GetVisibleText())
'IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").Exist
'Window("Nextech Main Window").Dialog("Billing Information").Dialog("regexpwndtitle:=Additional Claim Dates").Highlight
'Window("Nextech Main Window").Dialog("Billing Information").WinButton("regexpwndtitle:=\.\.\.", "index:=0").Highlight
'(CheckPoint("Patient Summary for VAC32, Patient32"))


'******
'	IsPresent = Dialog("Nextech Practice", "This will delete all related item, tracking, order, return and surgery information, Are you sure\?")
'	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Click
'	Dialog("NexTech Practice").WinButton("No").Click
'
'done recording

'=======

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "General 1")
Call VerifyGen1()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "General 2")
Call VerTabOpened ("General 2")
Call VerifyGen2()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Custom")
Call VerTabOpened ("Custom")  		' Practice stopped working here ***
Call VerifyCustom()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Tracking")
Call VerTabOpened ("Tracking")
Call VerifyTracking()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Follow Up")
Call VerTabOpened ("Follow Up")
Call VerifyFollowUp()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Notes")
Call VerTabOpened ("Notes")
Call VerifyNotes()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Appts.")
Call VerTabOpened ("Appts.")
Call VerifyAppts()
'
Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Quotes")
Call VerTabOpened ("Quotes")
Call VerifyQuotes()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Insurance")
Call VerTabOpened ("Insurance")
Call VerifyInsurance()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Billing")
Call VerTabOpened ("Billing")
Call VerifyBilling()




'=======



'Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "History")
'Call VerTabOpened ("History")
'Call VerifyHistory()

'Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "NexPhoto")
'Call VerTabOpened ("NexPhoto")
'Call VerifyNexPhoto()
'
'
'' There is an issue here where the popup is happening in the function.  We either have to access the NexEMR tab via the keyboard, or
'' check for the popup in the change tab function (which doesn't sound like the best idea
'
''Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "NexEMR")
''had to add the following code because after a database refresh, the provider information is lost
'
'
Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl").Type micAltDwn + "e" + micAltUp @@ hightlight id_;_4852452_;_script infofile_;_ZIP::ssf7.xml_;_
''IsPresent = VerifyPopups("NexTech Practice", "You have 13 EMR Provider licenses, but have not assigned a provider to any of them\.  Would you like to setup your EMR Provider licenses now\?  You will not be able to lock any EMNs until you have assigned at least one provider to a license\.  You can conf.*")
''Call RecordResults(IsPresent, "EMR popup confirmed")
'If Window("Nextech Main Window").Dialog("NexTech Practice PUP1").Exist Then @@ hightlight id_;_1578542_;_script infofile_;_ZIP::ssf5.xml_;_
'Dialog("NexTech Practice").WinButton("No").Click 
'	
'End If
' @@ hightlight id_;_2298670_;_script infofile_;_ZIP::ssf6.xml_;_
'
' 'If Dialog("text:=" & titleText).Static("regexpwndtitle:=" & messageRegEx).Exist(5) Then
'wait (1)
Call VerTabOpened ("NexEMR")
Call VerifyNexEMR()
'
'
Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl").Type micAltDwn + "d" + micAltUp
If Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx License Configuration").Exist Then
	Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx License Configuration").WinButton("regexpwndtitle:=&No").Click
End If
'==============================================
Call VerTabOpened ("Medications")
Call VerifyMedications()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Labs")
Call VerTabOpened ("Labs")
Call VerifyLabs()

Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Dashboard")
Call VerTabOpened ("Dashboard")
Call VerifyDashboard ()
'*******
wait(1)



wait(1)


RunAction "LogOut [GlobalPracticeActions]", oneIteration

Sub VerifyDashboard()


End Sub

Sub VerifyNexEMR()

wait(1)

If Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").Exist(1) Then
    Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Click
	Call RecordResults(True, "Closed dictation license popup")
End If	
	
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Aged Unlocked EMNs").Exist (4) Then
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Aged Unlocked EMNs").WinButton("regexpwndtitle:=Ignore").Click
	Call RecordResults(True, "NexEMR: Aged Unlocked EMNs dialog opened.")	
End If
	
Window("Nextech Main Window").Window("Patients").WinButton("Configure Groups").Click

IsPresent = Window("Nextech Main Window").Dialog("Configure Groups").Check (CheckPoint("Configure Groups"))
	Call RecordResults(IsPresent, "NexEMR: Configure Groups")	
Window("Nextech Main Window").Dialog("Configure Groups").WinButton("Cancel").Click

Window("Nextech Main Window").Window("Patients").WinButton("EMR Lock Manager").Click

If Window("Nextech Main Window").Dialog("EMR Lock Manager").Exist (5) Then
	Window("Nextech Main Window").Dialog("EMR Lock Manager").WinButton("Close").Click
	Call RecordResults(True, "NexEMR: Lock Manager")	
End If

Window("Nextech Main Window").Window("Patients").WinButton("NexEMR: Edit EMN Templates").Click
WinMenu("ContextMenu").Select "Collection 1"

If Window("Nextech Main Window").Dialog("Select EMN Template").Exist (5) Then
	Call RecordResults(IsPresent, "NexEMR: EMN Templates")	
	Window("Nextech Main Window").Dialog("Select EMN Template").WinButton("Cancel").Click
End If

Window("Nextech Main Window").Window("Patients").WinButton("NexEMR: Edit EMN Templates").Click
WinMenu("ContextMenu").Select "New Collection..."

IsPresent = Window("Nextech Main Window").Dialog("text:=Input").Static("text:=Enter a name for the new.*").Check (CheckPoint("Enter a name for the new collection"))
	Call RecordResults(IsPresent, "NexEMR: Configure Groups")	
Window("Nextech Main Window").Dialog("Input").WinButton("Cancel").Click
Window("Nextech Main Window").Window("Patients").WinButton("NexEMR: Edit EMN Templates").Click
WinMenu("ContextMenu").Select "Manage Collections..."

IsPresent = Window("Nextech Main Window").Dialog("EMR Collections").Check (CheckPoint("EMR Collections"))
	Call RecordResults(IsPresent, "NexEMR: Collections")	
Window("Nextech Main Window").Dialog("EMR Collections").WinButton("Cancel").Click
Window("Nextech Main Window").Window("Patients").WinButton("NexEMR: Edit EMN Templates").Click
WinMenu("ContextMenu").Select "Manage Templates..."

IsPresent = Window("Nextech Main Window").Dialog("EMN Templates").Check (CheckPoint("EMN Templates"))
	Call RecordResults(IsPresent, "NexEMR: EMN Templates")	
Window("Nextech Main Window").Dialog("EMN Templates").WinButton("New").Click

IsPresent = Dialog("Select EMR Collection").Check (CheckPoint("Select EMR Collection"))
	Call RecordResults(IsPresent, "NexEMR: EMR Collections")	
Dialog("Select EMR Collection").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("EMN Templates").WinButton("Filter").Click

IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a list of required keywords_8"))
	Call RecordResults(IsPresent, "NexEMR: Configure Groups")	
Dialog("Input").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("EMN Templates").WinButton("Close").Click
Window("Nextech Main Window").Window("Patients").WinButton("Configure Columns").Click

IsPresent = Window("Nextech Main Window").Dialog("Configure Columns").Check (CheckPoint("Configure Columns"))
	Call RecordResults(IsPresent, "NexEMR: Configure Columns")	
Window("Nextech Main Window").Dialog("Configure Columns").WinButton("Close").Click
Window("Nextech Main Window").Window("Patients").WinButton("Pts. To Bill").Click

IsPresent = Window("Nextech Main Window").Dialog("EMNs To Be Billed").Check (CheckPoint("EMNs To Be Billed_2"))
	Call RecordResults(IsPresent, "NexEMR: EMNs To Be Billed")	
Window("Nextech Main Window").Dialog("EMNs To Be Billed").WinButton("Close").Click
wait(1)	
Window("Nextech Main Window").Window("Patients").WinButton("Problem List").Click

IsPresent = Window("Nextech Main Window").Dialog("Problem List For VAC32,").Check (CheckPoint("Problem List For VAC32, Patient32"))
	Call RecordResults(IsPresent, "NexEMR: Problem List")	
Window("Nextech Main Window").Dialog("Problem List For VAC32,").WinButton("New Problem").Click
IsPresent = Dialog("Problem Editor").Check (CheckPoint("Problem Editor"))
	Call RecordResults(IsPresent, "NexEMR: Problem Editor")	

' Chronicity ellipse

Dialog("Problem Editor").WinButton("WinButton").Click

IsPresent = Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box_7"))
	Call RecordResults(IsPresent, "NexEMR: Edit Combo Box")	
Dialog("Edit Combo Box").WinButton("Close").Click
' SNOMED Code ellipse

Dialog("Problem Editor").WinButton("WinButton_2").Click

IsPresent = Dialog("UMLS Setup").Check (CheckPoint("UMLS Setup"))
	Call RecordResults(IsPresent, "NexEMR: UMLS Setup")	


Dialog("UMLS Setup").WinButton("Cancel").Click

IsPresent = VerifyPopups("NexTech Practice", "The UMLS login information is not set\.  A UMLS account is required to search for codes\.\r\n\r\nYou may configure the UMLS login information by navigating to the Administrator module\.  Then, click the 'Tools' menu item and select 'UMLS Settings\.\.\.'")


'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("The UMLS login information is not set.  A UMLS account is required to search for codes.    You may configure the UMLS login information by navigating to the Administrator module.  Then, click the 'Tools' menu item and select 'UMLS Settings'"))
	Call RecordResults(IsPresent, "NexEMR: UMLS popup")	
Dialog("NexTech Practice").WinButton("OK").Click

IsPresent = Dialog("UTS Code Import").Check (CheckPoint("UTS Code Import"))
	Call RecordResults(IsPresent, "NexEMR: UTS Code Import")	
Dialog("UTS Code Import").WinButton("Cancel").Click
Dialog("Problem Editor").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("Problem List For VAC32,").WinButton("Close").Click
Window("Nextech Main Window").Window("Patients").WinButton("EMR Summary").Click

IsPresent = Window("Nextech Main Window").Dialog("EMR Summary").Check (CheckPoint("EMR Summary"))
	Call RecordResults(IsPresent, "NexEMR: EMR Summary")	
Window("Nextech Main Window").Dialog("EMR Summary").WinButton("Configure").Click

IsPresent = Dialog("EMR Summary Configuration").Check (CheckPoint("EMR Summary Configuration"))
	Call RecordResults(IsPresent, "NexEMR: EMR Summary Configuration")	
Dialog("EMR Summary Configuration").WinButton("Add Category").Click

IsPresent = Dialog("Single Selection").Static("Select a category:").Check (CheckPoint("Select a category:"))
	Call RecordResults(IsPresent, "NexEMR: Configure Groups")	
Dialog("Single Selection").WinButton("Cancel").Click
Dialog("EMR Summary Configuration").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("EMR Summary").WinButton("Close").Click


Window("Nextech Main Window").Window("Patients").WinButton("Pt. Summary").Click

If IsPresent = Dialog("Patient Summary").Exist Then
	Call RecordResults(IsPresent, "NexEMR: Patient Summary")	
Dialog("Patient Summary").WinButton("OK").Click 
	
End If

IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Patient Summary for VAC32, Patient32").Check (CheckPoint("Patient Summary for VAC32, Patient32"))
Call RecordResults(IsPresent, "NexEMR: Patient Summary")	

Window("Nextech Main Window").Dialog("Patient Summary for VAC32,").WinButton("Help").Click

'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("The Patient Summary bottom screen displays the following information for the current patient:   - All future Appointments, and the last three appointments that have occurred within the last three months.  - All unpaid Bills, and the last three bills that"))

IsPresent = VerifyPopups("NexTech Practice", "All unpaid Bills, and the last three bills that")

Call RecordResults(IsPresent, "NexEMR: Configure Groups")	
Dialog("NexTech Practice").WinButton("OK").Click
Window("Nextech Main Window").Dialog("Patient Summary for VAC32,").WinButton("Configure").Click

IsPresent = Dialog("Configure Patient Summary").Check (CheckPoint("Configure Patient Summary"))
	Call RecordResults(IsPresent, "NexEMR: Configure Patient Summary")	
Dialog("Configure Patient Summary").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("Patient Summary for VAC32,").WinButton("Close").Click
Window("Nextech Main Window").Window("Patients").WinButton("Pts. Seen").Click

If Window("Nextech Main Window").Dialog("text:=EMR Search and Review -.* EMNs").Exist (5) Then
	Call RecordResults(IsPresent, "NexEMR: Search and Review Opened correctly")	
	'Window("Nextech Main Window").Dialog("text:=EMR Search and Review -.* EMNs").WinButton("regexpwndtitle:=Close").Click
End If

Window("Nextech Main Window").Dialog("EMR Search and Review").WinButton("Sign Selected EMNs").Click
IsPresent = Window("Nextech Main Window").Dialog("EMR Search and Review").Dialog("NexTech Practice").Static("You must select at least").Check (CheckPoint("You must select at least one EMN."))
	Call RecordResults(IsPresent, "NexEMR: Popup")	
Window("Nextech Main Window").Dialog("EMR Search and Review").Dialog("NexTech Practice").WinButton("OK").Click
Window("Nextech Main Window").Dialog("EMR Search and Review").WinButton("Close").Click
Window("Nextech Main Window").Window("Patients").WinButton("EMR Analysis").Click

IsPresent = Window("Nextech Main Window").Dialog("EMR Analysis").Check (CheckPoint("EMR Analysis"))
	Call RecordResults(IsPresent, "NexEMR: EMR Analysis")	
Window("Nextech Main Window").Dialog("EMR Analysis").WinButton("Add").Click

IsPresent = Dialog("EMR Analysis Configuration").Check (CheckPoint("EMR Analysis Configuration"))
	Call RecordResults(IsPresent, "NexEMR: EMR Analysis Configuration")	
Dialog("EMR Analysis Configuration").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("EMR Analysis").WinButton("Edit").Click

'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("You must first select a configuration from the list before editing."))
IsPresent = VerifyPopups("NexTech Practice", "You must first select a configuration from the list before editing\.")
Call RecordResults(IsPresent, "NexEMR: Popup")	
Dialog("NexTech Practice").WinButton("OK").Click
Window("Nextech Main Window").Dialog("EMR Analysis").WinButton("Export Results").Click

'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("There are no results in the list to export."))
IsPresent = VerifyPopups("NexTech Practice", "There are no results in the list to export\.")
Call RecordResults(IsPresent, "NexEMR: Popup")	
Dialog("NexTech Practice").WinButton("OK").Click
Window("Nextech Main Window").Dialog("EMR Analysis").WinButton("Load Results").Click

'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("You must first select a configuration from the list before displaying results\."))
IsPresent = VerifyPopups("NexTech Practice", "You must first select a configuration from the list before displaying results\.")

Call RecordResults(IsPresent, "NexEMR: Select a Service Popup")	
Dialog("NexTech Practice").WinButton("OK").Click
Window("Nextech Main Window").Dialog("EMR Analysis").WinButton("Close").Click
Window("Nextech Main Window").Window("Patients").WinButton("Wellness").Click

IsPresent = Window("Nextech Main Window").Dialog("Wellness Alerts for VAC32,").Check (CheckPoint("Wellness Alerts for VAC32, Patient32"))
	Call RecordResults(IsPresent, "NexEMR: Wellness Alerts")	
Window("Nextech Main Window").Dialog("Wellness Alerts for VAC32,").WinButton("Close").Click
End Sub

Sub VerifyNexPhoto()

IsPresent = Window("Nextech Main Window").Window("Patients").WinObject("regexpwndtitle:=Global Search").Exist
Call RecordResults(IsPresent, "NexPhoto Tab is Open")

'Clicking on Global Search Button
Window("Nextech Main Window").Window("Patients").WinObject("regexpwndtitle:=Global Search").Click
IsPresent = Window("Nextech Main Window").Window("regexpwndtitle:=NexPhoto Global Search").WinObject("regexpwndtitle:=Search").Exist
Call RecordResults(IsPresent, "NexPhoto Global Search Window is Present")
Window("Nextech Main Window").Window("regexpwndtitle:=NexPhoto Global Search").Close

'Click on Import Button
Window("Nextech Main Window").Window("Patients").WinObject("regexpwndtitle:=Import").Click
IsPresent = Window("regexpwndtitle:=Import From Computer").WinObject("regexpwndtitle:=Apply").Exist
Call RecordResults(IsPresent, "Import From Computer Window is Present")
If Window("regexpwndtitle:=Import From Computer").Window("regexpwndtitle:=Import From Computer: Setup").Exist Then
	Window("regexpwndtitle:=Import From Computer").Window("regexpwndtitle:=Import From Computer: Setup").WinObject("regexpwndtitle:=Cancel").Click
End If
If Window("regexpwndtitle:=Import From Computer").Window("regexpwndtitle:=Import From Computer: Setup").Dialog("regexpwndtitle:=Photo Import").Exist Then
	Window("regexpwndtitle:=Import From Computer").Window("regexpwndtitle:=Import From Computer: Setup").Dialog("regexpwndtitle:=Photo Import").WinButton("regexpwndtitle:=&No").Click
End If

Window("regexpwndtitle:=Import From Computer").Close
End Sub

Sub VerifyHistory()

	'Clicking on New Button and Attach Existing Folder in WinMenu
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Attach Existing Folder"
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Browse for Folder").Exist
Call RecordResults(IsPresent, "Browse for Folder Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Browse for Folder").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on WinMenu Import and Attach Existing File
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import and Attach Existing File"
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Open Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Import From Scanner/Camera---Scan as PDF...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import from Scanner/Camera;Scan as PDF..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import From Scanner/Camera---Scan as multi-page PDF...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import from Scanner/Camera;Scan as Multi-Page PDF..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import From Scanner/Camera---Scan as Image...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import from Scanner/Camera;Scan as Image..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import From Scanner/Camera---Scan Multiple Documents...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import from Scanner/Camera;Scan Multiple Documents..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Associate With Patient").Exist
Call RecordResults(IsPresent, "Multiple Documents Scan Dialog is Present")

'Clicking on Associate with Patients Button in Multiplan Scan Documents
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Associate With Patient").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Group as Single Document in Multiple Scan Documents
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Group as Single Document").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Begin Scan... button in Multiple Document Scan 
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Begin Scan\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Show Image Preview Button in Multiple Document Scan 
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Show Image Preview").Click
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Scanned Photo Viewer").WinButton("regexpwndtitle:=C&lose").Exist
Call RecordResults(IsPresent, "Scanned Photo Viewer Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Scanned Photo Viewer").WinButton("regexpwndtitle:=C&lose").Click

Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Import from Scanner/Camera---Select TWAIN Input Source...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import from Scanner/Camera;Select TWAIN Input Source..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Source").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Select Source Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Source").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Import from PDA
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import from PDA"
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import from Device
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Import from Device"
IsPresent = Dialog("regexpwndtitle:=Device Import").WinButton("regexpwndtitle:=&Close").Exist
Call RecordResults(IsPresent, "Device Import Dialog is Present")
Dialog("regexpwndtitle:=Device Import").WinButton("regexpwndtitle:=&Close").Click

'Clicking on Create CCD Summary...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Create CCD Summary..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Create New Document
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Create New Document"
wait 3
IsPresent = Window("regexpwndtitle:=Word").Exist
Call RecordResults(IsPresent, "Word Document is Present")
Window("regexpwndtitle:=Word").Close

'Clicking on Merge New Document
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Merge New Document"
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Browse for file dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Merge New Packet
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Merge New Packet"
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Packet").WinButton("regexpwndtitle:=\.\.\.").Exist
Call RecordResults(IsPresent, "Merge packet Dialog is Present")

'Clicking on the ellipsis button in Merge Packet Dialog
Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Packet").WinButton("regexpwndtitle:=\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter Writing Packets").Exist
Call RecordResults(IsPresent, "Configure Letter Writing Packets is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter Writing Packets").WinButton("regexpwndtitle:=Add").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter.*").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Add copy
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter Writing Packets").WinButton("regexpwndtitle:=Add Copy").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter.*").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter.*").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=Configure Letter.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on Merge
Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Packet").WinButton("regexpwndtitle:=Merge").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Select Packet").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Packet").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on New Case History
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "New Case History"
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Select a Preference Card").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Select a Preference Card Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Select a Preference Card").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Parse Transcription File...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=New").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Parse Transcription File..."
IsPresent = Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").WinButton("regexpwndtitle:=&Add File\(s\)\.\.\.").Exist
Call RecordResults(IsPresent, "Transcription Parsing Dialog is Present")
Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").WinButton("regexpwndtitle:=&Add File\(s\)\.\.\.").Click
IsPresent = Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Browse for Files Dialog is Present")

Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Configure Button in Transaction parsing Dialog
Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").WinButton("regexpwndtitle:=Confi&gure\.\.\.").Click
IsPresent = Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").Dialog("regexpwndtitle:=Configure Transcription Parsing").Exist
Call RecordResults(IsPresent, "Configure Transcription Parsing Dialog is Present")
Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").Dialog("regexpwndtitle:=Configure Transcription Parsing").WinButton("regexpwndtitle:=Cancel").Click

'Closing the Transcription Parsing Dialog
Dialog("regexpwndtitle:=Transcription Parsing for Multiple Patients").WinButton("regexpwndtitle:=&Close").Click

'Clicking on Import from Scanner/Camera
'Clicking on Import From Scanner/Camera---Scan as PDF...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Scan as PDF..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import From Scanner/Camera---Scan as multi-page PDF...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Scan as Multi-Page PDF..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import From Scanner/Camera---Scan as Image...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Scan as Image..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import From Scanner/Camera---Scan Multiple Documents...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Scan Multiple Documents..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Associate With Patient").Exist
Call RecordResults(IsPresent, "Multiple Documents Scan Dialog is Present")

'Clicking on Associate with Patients Button in Multiplan Scan Documents
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Associate With Patient").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Group as Single Document in Multiple Scan Documents
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Group as Single Document").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Begin Scan... button in Multiple Document Scan 
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Begin Scan\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Multiple Document Scan").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Show Image Preview Button in Multiple Document Scan 
Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Show Image Preview").Click
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Scanned Photo Viewer").WinButton("regexpwndtitle:=C&lose").Exist
Call RecordResults(IsPresent, "Scanned Photo Viewer Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Scanned Photo Viewer").WinButton("regexpwndtitle:=C&lose").Click

Window("Nextech Main Window").Dialog("regexpwndtitle:=Multiple Document Scan").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Import from Scanner/Camera---Select TWAIN Input Source...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Select TWAIN Input Source..."
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Source").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Select Source Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=Select Source").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Record Audio Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Record Audio").Click
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Open Default Folder
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Open Default Folder").Click
wait(2)
IsPresent = Window("regexpwndclass:=CabinetWClass","regexpwndtitle:=VAC32-Patient32-43").Exist
Call RecordResults(IsPresent, "Documents window is Open")
Window("regexpwndclass:=CabinetWClass","regexpwndtitle:=VAC32-Patient32-43").Close

'Clicking on Edit Categories
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Edit Categories").Click
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&New").Exist
Call RecordResults(IsPresent, "Note / Follow-Up Categories Dialog is Present")

'Clicking on New Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&New").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").Exist
Call RecordResults(IsPresent, "Input Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Adv. EMR Merge...
Window("Nextech Main Window").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&Adv\. EMR Merge\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:= Default Categories").Exist
Call RecordResults(IsPresent, "Default Categories Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:= Default Categories").WinButton("regexpwndtitle:=Cancel","index:=1").Click

'Clicking on Combine Categories 
Window("Nextech Main Window").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&Combine Categories\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Combine Note Categories").Exist
Call RecordResults(IsPresent, "Combine Note Categories Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Combine Note Categories").WinButton("regexpwndtitle:=Close").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=Clos&e").Click

'Clicking on the Blank Pane 
Window("Nextech Main Window").Window("Patients").WinObject("object class:=Afx:42800000:8b").Click 300,23

'Clicking on Detach File Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Detach File\(s\)").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Detach And Delete File"
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on the Blank Pane 
Window("Nextech Main Window").Window("Patients").WinObject("object class:=Afx:42800000:8b").Click 300,23

'Clicking on Detach File Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Detach File\(s\)").Click
Window("Nextech Main Window").WinMenu("menuobjtype:=3").Select "Detach And Delete File"
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Import from Device
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Import from Device").Click
IsPresent = Dialog("regexpwndtitle:=Device Import").WinButton("regexpwndtitle:=&Close").Exist
Call RecordResults(IsPresent, "Device Import Dialog is Present")
Dialog("regexpwndtitle:=Device Import").WinButton("regexpwndtitle:=&Close").Click
End Sub

Sub SelectNewPayRefAdj()
	
	Window("Nextech Main Window").Window("Patients").WinCheckBox("Remember Column Widths").Set "OFF"
	Window("Nextech Main Window").Window("Patients").WinCheckBox("Remember Column Widths").Set "ON"
	Window("Nextech Main Window").Window("Patients").WinCheckBox("Remember Column Widths").Type  micTab
	Window("Nextech Main Window").Window("Patients").WinButton("Billing: New Bill button").Type  micTab 
'	Window("Window").Window("Window").WinButton("New Bill").Type  micTab
	Window("Nextech Main Window").Window("Patients").WinButton("New Pay/ Ref/ Adj").Type  micReturn
	'WinMenu("ContextMenu").Select "Create a New &Payment"

	
End Sub

Sub VerifyBilling()

	' Click Claim History, if first time, select don't show me again
	Window("Nextech Main Window").Window("Patients").WinButton("Claim History").Click
	If Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndtitle:=Practice").WinCheckBox("regexpwndtitle:=Don't Show Me Again").Exist (1) Then
		Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndtitle:=Practice").WinCheckBox("regexpwndtitle:=Don't Show Me Again").Set "ON"
		Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click
		Call RecordResults(IsPresent, "Don't show again popup opened and clicked")	
	Else	
		IsPresent = Window("Nextech Main Window").Window("Claim History").Static("0 of 0").Check (CheckPoint("0 of 0_2"))
		Call RecordResults(IsPresent, "Billing: Claim History button click dialog opened")	
		Window("Nextech Main Window").Window("Claim History").Close
	End If	
	
	' Click Responsible Parties
	Window("Nextech Main Window").Window("Patients").WinButton("Responsible Parties").Click
	IsPresent = Window("Nextech Main Window").Dialog("Responsible Party").Check (CheckPoint("Responsible Party"))
	Call RecordResults(IsPresent, "Billing: Responsible Parties button click dialog opened")	
	Window("Nextech Main Window").Dialog("Responsible Party").WinButton("Cancel").Click

	' Click Show Packages

	Window("Nextech Main Window").Window("Patients").WinButton("Show Packages").Click
	IsPresent = Window("Nextech Main Window").Dialog("Packages").Check (CheckPoint("Packages"))
	Call RecordResults(IsPresent, "Billing: Show Packages button click dialog opened")	
	Window("Nextech Main Window").Dialog("Packages").Close
	
	' Click Show Quotes

	Window("Nextech Main Window").Window("Patients").WinButton("Show Quotes").Click
	IsPresent = Window("Nextech Main Window").Dialog("Quote").Check (CheckPoint("Quote"))
	Call RecordResults(IsPresent, "Billing: Show Quotes button click dialog opened")	
	Window("Nextech Main Window").Dialog("Quote").Close


	' Click EMNs To Be Billed
	
	Window("Nextech Main Window").Window("Patients").WinButton("EMNs To Be Billed").Click
	IsPresent = Window("Nextech Main Window").Dialog("EMNs To Be Billed").Check (CheckPoint("EMNs To Be Billed"))
	Call RecordResults(IsPresent, "Billing: EMNs To Be Billed button click dialog opened")	
	Window("Nextech Main Window").Dialog("EMNs To Be Billed").WinButton("Close").Click
'=======	
	' Click Optical Order History
	
	Window("Nextech Main Window").Window("Patients").WinButton("Optical Order History").Click
	IsPresent = Window("Nextech Main Window").Dialog("Optical Order History").Check (CheckPoint("Optical Order History for VAC32, Patient32"))
	Call RecordResults(IsPresent, "Billing: Optical Order History button click dialog opened")	
	Window("Nextech Main Window").Dialog("Optical Order History").WinButton("Close").Click
	
	' Click New Bill  (a ton ton here!)
	Window("Nextech Main Window").Window("Patients").WinButton("Billing: New Bill button").Click
	IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Check (CheckPoint("Billing Information"))
	Call RecordResults(IsPresent, "Billing: New Bill button click dialog opened")	
'=======
		' Click Merge To Word
		Window("Nextech Main Window").Dialog("Billing Information").WinButton("Merge To Word").Click
		IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("There are no charges on this bill. The merge will not be generated."))
		Call RecordResults(IsPresent, "Billing: Merge to word button click dialog opened")	
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click
	
		' Click Print Preview
		Window("Nextech Main Window").Dialog("Billing Information").WinButton("Print  Preview").Click
		IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("There are no charges on this bill. The preview will not be generated."))
		Call RecordResults(IsPresent, "Billing: Print Preview button click dialog opened")	
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click
			
		' Click the Toggle filter button
		Window("Nextech Main Window").Dialog("Billing Information").WinButton("Button").Click
		IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a list of required keywords"))
		Call RecordResults(IsPresent, "Billing: Filter button click dialog opened")	
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("Input").WinButton("Cancel").Click
		
		' Click Diagnosis QuickList
		
		Window("Nextech Main Window").Dialog("Billing Information").WinButton("Diagnosis QuickList").Click
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("My QuickList").Check CheckPoint("My QuickList")
		Call RecordResults(IsPresent, "Billing: Diagnosis QuickList button click dialog opened")	
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("My QuickList").WinButton("Close").Click
	
		' Click Apply Discounts to All Charges
		
		Window("Nextech Main Window").Dialog("Billing Information").WinButton("Apply Discount to All").Click
		IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("All charges on this bill are original or void charges, and cannot be edited."))
		Call RecordResults(IsPresent, "Billing: Apply Discount Button to All button click dialog opened")	
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click
	
		' Click Modify Discounts
		Window("Nextech Main Window").Dialog("Billing Information").WinButton("Modify Discounts").Click
		IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Modify Discounts").Check (CheckPoint("Modify Discounts_2"))
		Call RecordResults(IsPresent, "Billing: Modify Discounts Button to All button click dialog opened")	
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("Modify Discounts").WinButton("Cancel").Click
	
		' Click Show Suggestions
		
		Window("Nextech Main Window").Dialog("Billing Information").WinButton("Show Suggestions").Click
		IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Suggested Sales").Check (CheckPoint("Suggested Sales_2"))
		Call RecordResults(IsPresent, "Billing: Show Suggestions button click dialog opened")	
		Window("Nextech Main Window").Dialog("Billing Information").Dialog("Suggested Sales").Close
'=======
' Billing tab
	

		' Click the Add/Edit Services button and select each menu item (a ton x ton x ton)
		
		' Add one time check here for CPT copyright 


' Billing Codes
		
			Window("Nextech Main Window").Dialog("Billing Information").WinButton("Add/Edit Services").Click
			WinMenu("ContextMenu").Select "Edit Service/Diag Codes"

'			IsPresent = OptionalStep.Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("Current Procedural Terminology (CPT) is copyright 2015 American Medical Association. All Rights Reserved.   No fee schedules, basic units, relative values, or related listings are included in CPT. The AMA assumes no liability for the data contained herei"))
			'IsPresent = OptionalStep.Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Exists(2)
			IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Exist(2)
			
			If IsPresent Then
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click
				Call RecordResults(IsPresent, "Billing: First-Time CPT copyright popup closed.")
			Else
				Call RecordResults(True, "Billing: First-Time CPT copyright popup was not present")
			End If
			
			IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").Check (CheckPoint("Billing Codes"))
		
			Call RecordResults(IsPresent, "Billing: Add/Edit Services button click dialog opened")	
	
			' A: Edit Service/Diag Codes
				' Click the service codes Add button
				
'=======
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("AddSvcCode").Click
				IsPresent = Dialog("New Service Code").Check (CheckPoint("New Service Code"))
				Dialog("New Service Code").WinButton("Cancel").Click
				Call RecordResults(IsPresent, "Billing: Add button click dialog opened")	
		
				' Click the service codes Mark Inactive button

				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Mark Inactive").Click
				IsPresent = Dialog("Deactivate Service Code").Check (CheckPoint("Deactivate Service Code"))
				Call RecordResults(IsPresent, "Billing: Mark Inactive button click dialog opened")	
				Dialog("Deactivate Service Code").WinButton("No").Click
				
				' Click the service codes Delete button
				
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Delete").Click
				If Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").Dialog("regexpwndtitle:=Practice").Static("regexpwndtitle:=This Service Code exists.*").Exist (5) Then
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click
					Call RecordResults(True, "Expected Service Code exists popup present")	
				Else
					Call RecordResults(False, "Expected Service Code exists popup did NOT display")					
				End If
				
							
				' Click the service codes Filter button
				
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Filter").Click
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a list of required keywords_2"))
				Call RecordResults(IsPresent, "Billing: Filter button click dialog opened")		
				Dialog("Input").WinButton("Cancel").Click
		
	'			 Click the Description ellipse button and select Edit Prescription Details for Quote
				
				
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("WinButton").Click
				WinMenu("ContextMenu").Select "Edit Procedure Description for Quotes..."
				IsPresent = Dialog("Procedure Description").Check (CheckPoint("Procedure Description"))
				Call RecordResults(IsPresent, "Billing: Description ellipse button click dialog opened")	
				Dialog("Procedure Description").WinButton("Cancel").Click
				
				' Click the Pay Group ellipse button		
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("WinButton_2").Click
				IsPresent = Dialog("Configure Pay Groups").Check (CheckPoint("Configure Pay Groups"))
				Call RecordResults(IsPresent, "Billing: Pay Group button click dialog opened")	

					' Click Add
					Dialog("Configure Pay Groups").WinButton("Add New Pay Group").Click
					IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a new Pay Group name:"))
					Call RecordResults(IsPresent, "Billing: Add New Pay Group button click dialog opened")	
					Dialog("Input").WinButton("Cancel").Click
					Dialog("Configure Pay Groups").WinButton("Cancel").Click

				' Click Advanced Setup... button
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Advanced Setup").Click
					IsPresent = Dialog("Advanced Pay Group Configurati").Check (CheckPoint("Advanced Pay Group Configuration"))
					Call RecordResults(IsPresent, "Billing: Advanced Setup button click dialog opened")	
	
						' Click the Pay Group To Apply ellipse button
						Dialog("Advanced Pay Group Configurati").WinButton("WinButton").Click
						IsPresent = Dialog("Configure Pay Groups").Check (CheckPoint("Configure Pay Groups_2"))
						Call RecordResults(IsPresent, "Billing: Pay Group To Apply ellipse button click dialog opened")	
				
							' Click the Add New button
							Dialog("Configure Pay Groups").WinButton("Add New Pay Group").Click
							IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a new Pay Group name:_2"))
							Call RecordResults(IsPresent, "Billing: Add New Pay Group button click dialog opened")
							
							Dialog("Input").WinButton("Cancel").Click
						Dialog("Configure Pay Groups").WinButton("Cancel").Click
					Dialog("Advanced Pay Group Configurati").WinButton("Close").Click

				' Click the Category Select button
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Select").Click
					IsPresent = Dialog("Service / Inventory Categories").Check (CheckPoint("Service / Inventory Categories"))
					Call RecordResults(IsPresent, "Billing: Category Select button click dialog opened")

						' Click the Add New button
						Dialog("Service / Inventory Categories").WinButton("Add New").Click
						IsPresent = Dialog("Input").Check (CheckPoint("Input_6"))
						Call RecordResults(IsPresent, "Billing: Add New button click dialog opened")
						
						Dialog("Input").WinButton("Cancel").Click
					Dialog("Service / Inventory Categories").WinButton("Cancel").Click
				
				' Click the Remove button
				
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Remove").Click
					IsPresent = Dialog("Practice").Static("This Service Code exists").Check (CheckPoint("Are you sure you wish to remove this code's categories?"))
					Call RecordResults(IsPresent, "Billing: Remove button click dialog opened")	
					Dialog("Practice").WinButton("No").Click
					
				' Click the Categorize Multiple button
				
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Categorize Multiple").Click
					IsPresent = Dialog("Update Categories").Check (CheckPoint("Update Categories"))
					Call RecordResults(IsPresent, "Billing: Categorize Multiple button click dialog opened")	
				
					' Click the Select button
					
						Dialog("Update Categories").WinButton("Select").Click
						IsPresent = Dialog("Service / Inventory Categories").Check (CheckPoint("Service / Inventory Categories_2"))
						Call RecordResults(IsPresent, "Billing: Services / Inventory button click dialog opened")	

						' Click the Add New button
						
							Dialog("Service / Inventory Categories").WinButton("Add New").Click
							IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter the new category name"))
							Call RecordResults(IsPresent, "Billing: Advanced Setup button click dialog opened")	
							Dialog("Input").WinButton("Cancel").Click
							
						Dialog("Service / Inventory Categories").WinButton("Cancel").Click
				
						' Click the Update button
						
						
						Dialog("Update Categories").WinButton("Update").Click
						IsPresent = Dialog("Practice").Static("This Service Code exists").Check (CheckPoint("Are you sure you want to configure all the selected services to have no category?"))
						Call RecordResults(IsPresent, "Billing: Update button click dialog opened")	
						Dialog("Practice").WinButton("No").Click

					Dialog("Update Categories").WinButton("Cancel").Click

				' Click the Shop Fee ellipse button
				
	
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("ShopFeeEllipsebutton").Click
				IsPresent = Dialog("Edit Shop Fees").Check (CheckPoint("Edit Shop Fees"))
				Call RecordResults(IsPresent, "Billing: Shop Fee ellipse button click dialog opened")	
				Dialog("Edit Shop Fees").WinButton("Close").Click
				' Click the Non-Billable Codes button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Non-Billable Codes").Click
				IsPresent = Dialog("Configure Non-Billable").Check (CheckPoint("Configure Non-Billable Service Codes"))
				Call RecordResults(IsPresent, "Billing: Non-Billable Codes button click dialog opened")	
				Dialog("Configure Non-Billable").WinButton("Cancel").Click
				' Click the Update Standard Fees button and select menu item:
					' Update by Percentage
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Update Standard Fees").Click
				WinMenu("ContextMenu").Select "Update By Percentage..."
				IsPresent = Dialog("Update Prices By Percentage").Check (CheckPoint("Update Prices By Percentage"))
				Call RecordResults(IsPresent, "Billing: Update Prices By Percentage button click dialog opened")	
				Dialog("Update Prices By Percentage").WinButton("Cancel").Click

					' Update by RVU			
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Update Standard Fees").Click
				WinMenu("ContextMenu").Select "Update By RVU..."
				IsPresent = Dialog("Update Prices By RVU").Check (CheckPoint("Update Prices By RVU"))
				Call RecordResults(IsPresent, "Billing: Shop Fee ellipse button click dialog opened")	
				Dialog("Update Prices By RVU").WinButton("Close").Click

					' Update from file
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Update Standard Fees").Click
				WinMenu("ContextMenu").Select "Update From File..."
				IsPresent= Dialog("Standard Fee Import and").Check (CheckPoint("Standard Fee Import and Update"))
				Call RecordResults(IsPresent, "Billing: Shop Fee ellipse button click dialog opened")	

						' Click the Browse button
					Dialog("Standard Fee Import and").WinButton("Browse").Click
					IsPresent = Dialog("Open").Check (CheckPoint("Open"))
					Call RecordResults(IsPresent, "Billing: Browse button click dialog opened")	
					Dialog("Open").WinButton("Cancel").Click

						' Click the Update button
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Update Standard Fees").Click
					WinMenu("ContextMenu").Select "Update From File..."
					' Click the Help button
					Dialog("Standard Fee Import and").WinButton("Help").Click
					IsPresent = Dialog("Practice").Check (CheckPoint("Practice_4"))
					Call RecordResults(IsPresent, "Billing: Help button click dialog opened")
					Dialog("Practice").WinButton("OK").Click
					Dialog("Standard Fee Import and").WinButton("Cancel").Click

				' Click Additional Service code Setup button and select menu item:
					' Claim
					
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Additional Service Code").Click
					WinMenu("ContextMenu").Select "Claim Setup..."
					IsPresent = Dialog("Claim Setup").Check (CheckPoint("Claim Setup"))
					Call RecordResults(IsPresent, "Billing: Claim Setup dialog opened")	
						
						' NDC Defaults
						Dialog("Claim Setup").WinButton("NDC Defaults").Click
						IsPresent = Dialog("NDC Defaults").Check (CheckPoint("NDC Defaults"))
						Call RecordResults(IsPresent, "Billing: Claim Setup dialog opened")	
						Dialog("NDC Defaults").WinButton("Cancel").Click

						' Click Configure Charge Level Providers
						Dialog("Claim Setup").WinButton("Configure Charge Level").Click
						IsPresent = Dialog("Configure Charge Level").Check (CheckPoint("Configure Charge Level Providers"))
						Call RecordResults(IsPresent, "Billing: Claim Setup dialog opened")	
						Dialog("Configure Charge Level").WinButton("Close").Click
						
					Dialog("Claim Setup").WinButton("Close").Click
				
'					 UB Setup
					
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Additional Service Code").Click
					WinMenu("ContextMenu").Select "UB Setup..."
					IsPresent = Dialog("UB Setup").Check (CheckPoint("UB Setup"))
					Call RecordResults(IsPresent, "Billing: UB Setup dialog opened")
					
						' Click the Advanced Revenue Code button
						Dialog("UB Setup").WinButton("Advanced Revenue Code").Click
						IsPresent = Dialog("Advanced Revenue Code").Check (CheckPoint("Advanced Revenue Code Setup"))
						Call RecordResults(IsPresent, "Billing: UB: Advanced Revenue Code Setup dialog opened")	
						Dialog("Advanced Revenue Code").WinButton("Close").Click
						
						' UB: ICD Setup Ellipse click	
						Dialog("UB Setup").WinButton("ICDSetupEllipse").Click
						IsPresent = Dialog("ICD Procedure Code Setup").Check (CheckPoint("ICD Procedure Code Setup"))
						Call RecordResults(IsPresent, "Billing: UB: ICD Procedure Code dialog opened")

							' Click UB: ICD 10 Add button
							Dialog("ICD Procedure Code Setup").WinButton("Add").Click
							IsPresent = Dialog("Blank Code").Static("You cannot enter a blank").Check (CheckPoint("You cannot enter a blank code."))
							Call RecordResults(IsPresent, "Billing: Add dialog opened")	
							Dialog("Blank Code").WinButton("OK").Click
							
						Dialog("ICD Procedure Code Setup").WinButton("Close").Click
					Dialog("UB Setup").WinButton("Close").Click
					
					' CCDA Setup
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Additional Service Code").Click
					WinMenu("ContextMenu").Select "CCDA Setup..."
					IsPresent = Dialog("CCDA Setup").Check (CheckPoint("CCDA Setup"))
					Call RecordResults(IsPresent, "Billing: CCDA Setup dialog opened")	
					Dialog("CCDA Setup").WinButton("Close").Click
										
					' Anesthesia/Facility Setup
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Additional Service Code").Click
					WinMenu("ContextMenu").Select "Anesthesia/Facility Setup..."
					IsPresent = Dialog("Anesthesia/Facility Setup").Check (CheckPoint("Anesthesia/Facility Setup"))
					Call RecordResults(IsPresent, "Billing: CCDA Setup dialog opened")	
					Dialog("Anesthesia/Facility Setup").WinButton("Close").Click
					
				' Click the Modifiers Add button
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Add Modifiers").Click
					IsPresent = Dialog("Modifier").Check (CheckPoint("Modifier"))
					Call RecordResults(IsPresent, "Billing: CCDA Setup dialog opened")	
					Dialog("Modifier").WinButton("Cancel").Click
				
				' Click the Multiple Service Code / Modifier Linking button
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Multiple Service Code").Click
					IsPresent = Dialog("Multiple Service Code").Check (CheckPoint("Multiple Service Code - Modifier Setup"))
					Call RecordResults(IsPresent, "Billing: Multiple Service Code - Modifier Setup dialog opened")	
					Dialog("Multiple Service Code").WinButton("Close").Click
				
				' Click the Diagnosis Codes Add button (ICD-9)
				
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinRadioButton("ICD-9").Set
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Add").Click
					IsPresent = Dialog("Diagnosis Code").Check (CheckPoint("Diagnosis Code"))
					Call RecordResults(IsPresent, "Billing: ICD9 Diagnosis Code Setup dialog opened")	
					Dialog("Diagnosis Code").WinButton("Cancel").Click
				
				' Click the Diagnosis Codes Add button (ICD-10)
				
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinRadioButton("ICD-10").Set
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Add").Click
					IsPresent = Dialog("Import ICD-10 Codes").Check (CheckPoint("Import ICD-10 Codes"))
					Call RecordResults(IsPresent, "Billing: ICD10 Diagnosis Code Setup dialog opened")	
					Dialog("Import ICD-10 Codes").WinButton("Cancel").Click

									
				' Click Payment Categories button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Payment Categories").Click
				IsPresent = Dialog("Payment Categories").Check (CheckPoint("Payment Categories"))
				Call RecordResults(IsPresent, "Billing: Payment Categories dialog opened")
				
					' Click Add
					Dialog("Payment Categories").WinButton("Add").Click
					IsPresent = Dialog("Practice").Static("Enter a Name for the New").Check (CheckPoint("Enter a Name for the New Item:"))
					Call RecordResults(IsPresent, "Billing: Add button click dialog opened")
					Dialog("Practice").WinButton("Cancel").Click
					
				Dialog("Payment Categories").WinButton("Close").Click

				' Click Config. Finance Charge Settin button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Config. Finance Charge").Click
				IsPresent = Dialog("Configure Charge Interest").Check (CheckPoint("Configure Charge Interest Settings"))
				Call RecordResults(IsPresent, "Billing: Finance Charge Settings button click dialog opened")	
				Dialog("Configure Charge Interest").WinButton("Cancel").Click
		
				' Click Configure billing Columns button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Configure Billing Columns").Click
				IsPresent = Dialog("Configure Billing Columns").Check (CheckPoint("Configure Billing Columns"))
				Call RecordResults(IsPresent, "Billing: Configure Billing Columns button click dialog opened")	
				Dialog("Configure Billing Columns").WinButton("Cancel").Click
		
				' Click Service Code/Diagnosis Code button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Service Code/Diagnosis").Click
				IsPresent = Dialog("Service Code/Diagnosis").Check (CheckPoint("Service Code/Diagnosis Code Linking"))
				Call RecordResults(IsPresent, "Billing: Service Code/Diagnosis Code button click dialog opened")	
		
				' Click Add/Edit Links button
				Dialog("Service Code/Diagnosis").WinButton("Add/Edit Links").Click
				Dialog("NexTech Practice_2").Static("Please select a service code").Check CheckPoint("PleaseSelectSvcCode")
				Dialog("NexTech Practice_2").WinButton("OK").Click
				Dialog("Service Code/Diagnosis").WinButton("Close").Click
		
				' Click Inactive codes, select Inactive Service Codes
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Inactive Codes").Click
				WinMenu("ContextMenu").Select "Inactive Service Codes..."
				IsPresent = Dialog("Inactive Service Codes").Check (CheckPoint("Inactive Service Codes"))
				Call RecordResults(IsPresent, "Billing: Inactive Codes > Service Codes button click dialog opened")	
				Dialog("Inactive Service Codes").WinButton("Close").Click
		
				' Click Inactive codes, select Inactive Diagnosis Codes
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Inactive Codes").Click
				WinMenu("ContextMenu").Select "Inactive Diagnosis Codes..."
				IsPresent = Dialog("Inactive Diagnosis Codes").Check (CheckPoint("Inactive Diagnosis Codes"))
				Call RecordResults(IsPresent, "Billing: Inactive Codes > Inactive Diagnosis Codes button click dialog opened")	
				Dialog("Inactive Diagnosis Codes").WinButton("Close").Click
				
				
				 'Click Inactive codes, select Inactive Modifiers
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Inactive Codes").Click
				WinMenu("ContextMenu").Select "Inactive Modifiers..."
				IsPresent = Dialog("Inactive Modifiers").Check (CheckPoint("Inactive Modifiers"))
				Call RecordResults(IsPresent, "Billing: Inactive Codes > Inactive Modifiers button click dialog opened")	
				Dialog("Inactive Modifiers").WinButton("Close").Click
'=======		
				' Click Import Code File button and select menu item Import AMA Codes
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Import Code File").Click
				wait(.5)
				WinMenu("ContextMenu").Select "Import AMA Code File"

				'fix this!!!
				'	BalNetPaymentsStr = Window("Nextech Main Window").Window("Patients").Static("NetPaymentsLabel").GetROProperty("text")
				'Dialog("NexTech Practice_2").Static("Please select a service code").Check CheckPoint("In the following window, please select your code file to import._2")
				
				IsPresent = VerifyPopups("NexTech Practice", "In the following window, please select your code file to import.")
				
'				MsgBox ("IsPresent: " & IsPresent)

'				IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("In the following window, please select your code file to import."))
				'Dialog("NexTech Practice").Static("Please select a service").Check CheckPoint("In the following window, please select your code file to import.")
				
				'Dialog("text:="&NxPracticecaption).Static("nativeclass:=Static", "text:="&ExpectedMessage).Exists(0)



				Call RecordResults(IsPresent, "Billing: Responsible Parties button click dialog opened")	
				Dialog("NexTech Practice").WinButton("OK").Click
				IsPresent = Dialog("Open").Check (CheckPoint("Open_2"))
				Call RecordResults(IsPresent, "Billing: Responsible Parties button click dialog opened")	
				Dialog("Open").WinButton("Cancel").Click
				
				' Click Import Code File button and select menu item Import OHIP Codes
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Import Code File").Click
				WinMenu("ContextMenu").Select "Import OHIP Code File"
				IsPresent = Dialog("Import OHIP Codes").Check (CheckPoint("Import OHIP Codes"))
				Call RecordResults(IsPresent, "Billing: Responsible Parties button click dialog opened")	
				Dialog("Import OHIP Codes").WinButton("Cancel").Click

				' Click Import AMA Codes button and select menu item:
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Import AMA Codes").Click
				WinMenu("ContextMenu").Select "Import AMA Service Codes..."
				
				
'				IsPresent = Dialog("Download?").Static("Your AMA file versions").Check (CheckPoint("Your AMA file versions are mismatched.  Would you like to check for the latest files now?"))
'				Call RecordResults(IsPresent, "Billing: AMA versions mismatched from Import AMA Service Codes")	
'				Dialog("Download?").WinButton("No").Click
'				
				IsPresent = Dialog("Import AMA Codes").Check (CheckPoint("Import AMA Codes"))
				Call RecordResults(IsPresent, "Billing: Import AMA Codes dialog opened")
				
				' Click the Check for Updates button
				Dialog("Import AMA Codes").WinButton("Check for Updates").Click
				IsPresent = Dialog("NexTech Practice").Check (CheckPoint("NexTech Practice_5"))
				Call RecordResults(IsPresent, "Billing: Check for Updates popup response opened")
				
'				 Click the Filter button
				Dialog("NexTech Practice").WinButton("OK").Click
				
				Dialog("Import AMA Codes").WinButton("Filter").Click
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a list of required keywords_3"))
				Call RecordResults(IsPresent, "Billing: Popup to enter keywords")
				
				Dialog("Input").WinButton("Cancel").Click
				' Import AMA Service Codes
				Dialog("Import AMA Codes").WinButton("Import").Click
				
				IsPresent = VerifyPopups("NexTech Practice", "You must select at least one code before importing.")

				'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("You must select at least one code before importing."))
				Call RecordResults(IsPresent, "Billing: Popup to select a code")
				
				Dialog("NexTech Practice").WinButton("OK").Click
				Dialog("Import AMA Codes").WinButton("Cancel").Click

				' Import AMA Diagnosis Codes
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Import AMA Codes").Click
				WinMenu("ContextMenu").Select "Import AMA Diagnosis Codes..."
				
'				IsPresent = Dialog("Download?").Static("Your AMA file versions").Check (CheckPoint("Your AMA file versions are mismatched.  Would you like to check for the latest files now?"))
'				Call RecordResults(IsPresent, "Billing: AMA versions mismatched from Import AMA Dagnosis Codes")	
'				Dialog("Download?").WinButton("No").Click
'				
				IsPresent = Dialog("Import AMA Codes").Check (CheckPoint("Import AMA Codes"))
				Call RecordResults(IsPresent, "Billing: Import AMA Codes dialog open.")
				
				' Click the Check for Updates button
				Dialog("Import AMA Codes").WinButton("Check for Updates").Click
'				IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("No newer AMA code data sets were found.  You are up to date with the latest codes."))
				Dialog("NexTech Practice").WinButton("OK").Click
'				
				' Click the Filter button
				
				Dialog("Import AMA Codes").WinButton("Filter").Click
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a list of required keywords_4"))
				Call RecordResults(IsPresent, "Billing: Keyword popup")
				Dialog("Input").WinButton("Cancel").Click
				
				' Click the Import button
				Dialog("Import AMA Codes").WinButton("Import").Click
'				IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("You must select at least one code before importing._2"))
				IsPresent = VerifyPopups("NexTech Practice", "You must select at least one code before importing\.")
	'Msgbox ("IsPresent : " & IsPresent)
				Call RecordResults(IsPresent, "Billing: Please select a service code: Line#:1356")
				Dialog("NexTech Practice").WinButton("OK").Click
				Dialog("Import AMA Codes").WinButton("Cancel").Click

				' Click Import AMA Codes, select Import AMA Codes
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Import AMA Codes").Click
				WinMenu("ContextMenu").Select "Import AMA Service Modifiers..."

'				IsPresent = Dialog("Download?").Static("Your AMA file versions").Check (CheckPoint("Your AMA file versions are mismatched.  Would you like to check for the latest files now?"))
'				Call RecordResults(IsPresent, "Billing: AMA versions mismatched")	
'				Dialog("Download?").WinButton("No").Click
'				
				IsPresent = Dialog("Import AMA Codes").Check (CheckPoint("Import AMA Codes_3"))
				Call RecordResults(IsPresent, "Billing: Select Import AMA Services Modifiers...")

				Dialog("Import AMA Codes").WinButton("Filter").Click
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a list of required keywords_5"))
				Call RecordResults(IsPresent, "Billing: Filter button")	
				Dialog("Input").WinButton("Cancel").Click
				
				' Import AMA Service Modifiers
				Dialog("Import AMA Codes").WinButton("Import").Click
'				IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("You must select at least one code before importing._3"))
				IsPresent = VerifyPopups("NexTech Practice", "You must select at least one code before importing\.")
				Call RecordResults(IsPresent, "Billing: Import Click")	
				Dialog("NexTech Practice").WinButton("OK").Click
				Dialog("Import AMA Codes").WinButton("Cancel").Click

				' Click the Discount Categories button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Discount Categories").Click
				IsPresent = Dialog("Discount Categories").Check (CheckPoint("Discount Categories"))
				Call RecordResults(IsPresent, "Billing: Discount Categories")	
				
				' Click the Add button
				Dialog("Discount Categories").WinButton("Add").Click
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a new discount category:"))
				Call RecordResults(IsPresent, "Billing: Popup verified")	
				Dialog("Input").WinButton("Cancel").Click
				Dialog("Discount Categories").WinButton("Close").Click
				
				' Click the Receipt Config button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Receipt Config").Click
				IsPresent = Dialog("Configure Receipt").Check (CheckPoint("Configure Receipt"))
				Call RecordResults(IsPresent, "Billing: Click Receipt Config")	
				Dialog("Configure Receipt").WinButton("OK").Click
				
				' Click the Update TOS button.
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").WinButton("Update TOS").Click
				IsPresent = Dialog("Update TOS").Check (CheckPoint("Update TOS"))
				Call RecordResults(IsPresent, "Billing: Click Update TOS")	
				
				' Click the Update button.
				Dialog("Update TOS").WinButton("Update").Click
'				IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("You must select at least 1 code."))
				IsPresent = VerifyPopups("NexTech Practice", "You must select at least 1 code\.")
				Call RecordResults(IsPresent, "Billing: Popup opened")	
				Dialog("NexTech Practice").WinButton("OK").Click
				Dialog("Update TOS").WinButton("Close").Click

				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Billing Codes").Close
				
				'B: Edit Inventory Items
				Window("Nextech Main Window").Dialog("Billing Information").WinButton("Add/Edit Services").Click
				wait(1)
				WinMenu("ContextMenu").Select "Edit Inventory Items"
				wait(1)
				IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").Check (CheckPoint("Inventory Items"))
				Call RecordResults(IsPresent, "Billing: Add/Edit Services")	
				
				' Click the New Item button and select menu item:
'				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("New Item").Click
'				WinMenu("ContextMenu").Select ""
'----------------------------------------
				
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("New Item").Click
				WinMenu("ContextMenu").Select "New Product"
				IsPresent = Dialog("New Item").Check (CheckPoint("New Item_2"))
				Call RecordResults(IsPresent, "Billing: Add New Inventory Item")	
				
				Dialog("New Item").WinButton("Select").Click
				IsPresent = Dialog("Service / Inventory Categories").Check (CheckPoint("Service / Inventory Categories_5"))
				Call RecordResults(IsPresent, "Billing: Clicked select")	
				
				Dialog("Service / Inventory Categories").WinButton("Add New").Click
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter the new category name_4"))
				Call RecordResults(IsPresent, "Billing: Enter category name")	
				
				Dialog("Input").WinButton("Cancel").Click
				Dialog("Service / Inventory Categories").WinButton("Cancel").Click
				Dialog("New Item").WinButton("Cancel").Click
'				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("New Item").Click
'				wait (.1)
'				WinMenu("ContextMenu").Select "New Frame"
'				IsPresent = Dialog("Frames Data").Check (CheckPoint("Frames Data"))
'				Call RecordResults(IsPresent, "Billing: Frames data")	
'				
'				Dialog("Frames Data").WinButton("Select").Click
'				IsPresent = Dialog("Service / Inventory Categories").Check (CheckPoint("Service / Inventory Categories_6"))
'				Call RecordResults(IsPresent, "Billing: Inventory categories")	
'				
'				Dialog("Service / Inventory Categories").WinButton("Cancel").Click
'				Dialog("Frames Data").WinButton("Cancel").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("New Item").Click
				WinMenu("ContextMenu").Select "New Contact Lens"
				IsPresent = Dialog("Contact Lens").Check (CheckPoint("Contact Lens"))
				Call RecordResults(IsPresent, "Billing: Contact Lens")	
				
				Dialog("Contact Lens").WinButton("WinButton").Click
				IsPresent = Dialog("Edit Contact Lens Manufacturer").Check (CheckPoint("Edit Contact Lens Manufacturer"))
				Call RecordResults(IsPresent, "Billing: Lens Manufacturer")	
'				
'				Dialog("Edit Contact Lens Manufacturer").WinButton("Add").Click
'				Dialog("Edit Contact Lens Manufacturer").WinButton("Delete").Click
'				Dialog("Edit Contact Lens Manufacturer").Dialog("Delete?").WinButton("Yes").Click
				Dialog("Edit Contact Lens Manufacturer").Close
				Dialog("Contact Lens").WinButton("Select").Click
				IsPresent = Dialog("Service / Inventory Categories").Check (CheckPoint("Service / Inventory Categories_7"))
				Call RecordResults(IsPresent, "Billing: Svs/Inv Cats")	
				
				Dialog("Service / Inventory Categories").WinButton("Cancel").Click
				Dialog("Contact Lens").WinButton("WinButton_2").Click
				IsPresent = Dialog("Edit Contact Lens Type").Check (CheckPoint("Edit Contact Lens Type"))
				Call RecordResults(IsPresent, "Billing: Add/Edit Services")	
				
'				Dialog("Edit Contact Lens Type").WinButton("Add").Click
'				Dialog("Edit Contact Lens Type").WinButton("Delete").Click
'				Dialog("Edit Contact Lens Type").Dialog("Delete?").WinButton("Yes").Click
				Dialog("Edit Contact Lens Type").WinButton("Close").Click
				Dialog("Contact Lens").WinButton("Cancel").Click
				


'----------------------------------------


				
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Delete Item").Click
'				IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("This will delete all related item, tracking, order, return and surgery information, Are you sure?"))
				
				'Modified 20160223, 10:52
''				If Dialog("regexpwndtitle:=Practice").Exist Then
''					Call RecordResults(IsPresent, "Billing: Delete popup")	
''				End If			

				IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").Dialog("NexTech Practice").Exist
				If IsPresent Then
					Call RecordResults(True, "Valid Nextech instance")
					Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").Dialog("NexTech Practice").WinButton("No").Click @@ hightlight id_;_1117706_;_script infofile_;_ZIP::ssf1.xml_;_
				End If

				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("WinButton").Click
				IsPresent = Dialog("Edit Shop Fees").Check (CheckPoint("Edit Shop Fees_2"))
				Call RecordResults(IsPresent, "Billing: Shop Fees")	
				
				Dialog("Edit Shop Fees").WinButton("Close").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Adv. Revenue Code Setup").Click
				IsPresent = Dialog("Advanced Revenue Code").Check (CheckPoint("Advanced Revenue Code Setup_2"))
				Call RecordResults(IsPresent, "Billing: Adv Rev Code Setup")	
				
				Dialog("Advanced Revenue Code").WinButton("Close").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("NDC Info").Click
				WinMenu("ContextMenu").Select "Edit Claim Note..."
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a default billing note to add to charges with 'Claim' checked:"))
				Call RecordResults(IsPresent, "Billing: Edit Claim Note")	
				
				Dialog("Input").WinButton("Cancel").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("NDC Info").Click
				WinMenu("ContextMenu").Select "NDC Defaults..."
				IsPresent = Dialog("NDC Defaults").Check (CheckPoint("NDC Defaults_2"))
				Call RecordResults(IsPresent, "Billing: NDC Info")	
				
				Dialog("NDC Defaults").WinButton("Cancel").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Pending Cases").Click
				IsPresent = Dialog("Pending Case Histories").Check (CheckPoint("Pending Case Histories"))
				Call RecordResults(IsPresent, "Billing: Pending Cases")	
				
				Dialog("Pending Case Histories").WinButton("Close").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Adjust").Click
				IsPresent = Dialog("Item Adjustment").Check (CheckPoint("Item Adjustment"))
				Call RecordResults(IsPresent, "Billing: Item Adjustment")	
				
				Dialog("Item Adjustment").WinButton("Cancel").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Transfer").Click
'				IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("You cannot use this feature without two managed locations that track inventory for this product."))
				If Dialog("regexpwndtitle:=Transfer Inventory Quantities").Exist Then
					Dialog("regexpwndtitle:=Transfer Inventory Quantities").WinButton("regexpwndtitle:=Cancel").Click
				Else
					Dialog("NexTech Practice").WinButton("OK").Click
				End If	
								
				
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Select").Click
				IsPresent = Dialog("Service / Inventory Categories").Check (CheckPoint("Service / Inventory Categories_4"))
				Call RecordResults(IsPresent, "Billing: Service / Inv Cats")	
				
				Dialog("Service / Inventory Categories").WinButton("Add New").Click
				IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter the new category name_3"))
				Call RecordResults(IsPresent, "Billing: Add New")	
				
				Dialog("Input").WinButton("Cancel").Click
				Dialog("Service / Inventory Categories").WinButton("Cancel").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Remove").Click
				IsPresent = Dialog("Practice").Static("This Service Code exists").Check (CheckPoint("Are you sure you wish to remove this item's categories?"))
				Call RecordResults(IsPresent, "Billing: Remove popup")	
				
				Dialog("Practice").WinButton("No").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Add Supplier").Click
				IsPresent = Dialog("Choose A Supplier").Check (CheckPoint("Choose A Supplier"))
				Call RecordResults(IsPresent, "Billing: Choose Supplier")	
				
				Dialog("Choose A Supplier").WinButton("Add New").Click
				IsPresent = Window("Nextech Main Window").Dialog("Create New Contact").Check (CheckPoint("Create New Contact_2"))
				Call RecordResults(IsPresent, "Billing: Add New")	
				
				Window("Nextech Main Window").Dialog("Create New Contact").WinButton("Cancel").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Inactive Inventory Items").Click
				IsPresent = Dialog("Inactive Inventory Items").Check (CheckPoint("Inactive Inventory Items"))
				Call RecordResults(IsPresent, "Billing: Inactive Inventory")	
				
				Dialog("Inactive Inventory Items").WinButton("Close").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").WinButton("Link Products To Services").Click
				IsPresent = Dialog("Link Products To Services").Check (CheckPoint("Link Products To Services"))
				Call RecordResults(IsPresent, "Billing: Link Prod to Svc")	
				
				Dialog("Link Products To Services").WinButton("Close").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Inventory Items").Close
				
				
				' C: Edit Service Code Insurance Notes
				Window("Nextech Main Window").Dialog("Billing Information").WinButton("Add/Edit Services").Click
				WinMenu("ContextMenu").Select "Edit Service Code Insurance Notes"
				IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Insurance Code Notes").Check (CheckPoint("Insurance Code Notes_3"))
				Call RecordResults(IsPresent, "Billing: Insurance Code Notes")	
				
				' Click the Advanced button
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Insurance Code Notes").WinButton("Advanced").Click
				IsPresent = Dialog("Insurance Company Code").Check (CheckPoint("Insurance Company Code Notes_3"))
				Call RecordResults(IsPresent, "Billing: Insurance Code Notes Advanced")	
				
				IsPresent = Dialog("Insurance Company Code").WinButton("Close").Click
				Window("Nextech Main Window").Dialog("Billing Information").Dialog("Insurance Code Notes").WinButton("Close").Click

' Insurance tab

Window("Nextech Main Window").Dialog("Billing Information").ActiveX("NexTech DataList Control").Type micAltDwn + "i" + micAltUp
Window("Nextech Main Window").Dialog("Billing Information").WinButton("Select Ref Phy").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Practice").Static("There is no Referring").Check (CheckPoint("There is no Referring Physician selected on the General 2 tab."))
Call RecordResults(IsPresent, "Billing: Select Ref Phys button clicked")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("Practice").WinButton("OK").Click

' Click Select PCP button
Window("Nextech Main Window").Dialog("Billing Information").WinButton("Select PCP").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Practice").Static("There is no Referring").Check (CheckPoint("There is no Primary Care Physician selected on the General 2 tab."))
Call RecordResults(IsPresent, "Billing: Select PCP button click")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("Practice").WinButton("OK").Click

' Click the insurance referral ellipse button
Window("Nextech Main Window").Dialog("Billing Information").WinButton("regexpwndtitle:=\.\.\.", "index:=1").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("There are no active insurance companies for this patient which are marked to use insurance referrals."))
Call RecordResults(IsPresent, "Billing: Select Insurance Referral Ellipse button click")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click

' Click the Open Form button
Window("Nextech Main Window").Dialog("Billing Information").WinButton("Open Form").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("The bill must have at least one charge before a claim form can be made"))
Call RecordResults(IsPresent, "Billing: Select Open Form button click")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click

' Click the Print Form button
Window("Nextech Main Window").Dialog("Billing Information").WinButton("Print Form").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("The bill must have at least one charge before a claim form can be made_2"))
Call RecordResults(IsPresent, "Billing: Select Print Form button click")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click

' Click the Additional Claim Fields button
Window("Nextech Main Window").Dialog("Billing Information").WinButton("Additional Claim Fields").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Additional HCFA Claim").Check (CheckPoint("Additional HCFA Claim Fields"))
Call RecordResults(IsPresent, "Billing: Select Additional Claim Fields button click")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("Additional HCFA Claim").WinButton("Cancel").Click

' Click the Additional Claims Dates ellipse
Window("Nextech Main Window").Dialog("Billing Information").WinButton("regexpwndtitle:=\.\.\.", "index:=0").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("regexpwndtitle:=Additional Claim Dates").Exist
Call RecordResults(IsPresent, "Additional Claim Dates dialog opens successfully")
Window("Nextech Main Window").Dialog("Billing Information").Dialog("regexpwndtitle:=Additional Claim Dates").WinButton("regexpwndtitle:=OK").Click

' Additional Info. tab
Window("Nextech Main Window").Dialog("Billing Information").ActiveX("NexTech DataList Control").Type micAltDwn + "a" + micAltUp


' Click the Discharge Status ellipse button
Window("Nextech Main Window").Dialog("Billing Information").WinButton("DischargeListEllipseButton").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("Edit Discharge Status").Check (CheckPoint("Edit Discharge Status List"))
Call RecordResults(IsPresent, "Billing: Select Discharge Status Ellipse button click")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("Edit Discharge Status").WinButton("Close").Click
' Click Edit Additional Charge Information button
Window("Nextech Main Window").Dialog("Billing Information").WinButton("Edit Additional Charge").Click
IsPresent = Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").Static("There are no charges on").Check (CheckPoint("There are no charges currently on the bill. You must add charges before you can edit their data."))
Call RecordResults(IsPresent, "Billing: Select Edit Additional Charge Info button click")	
Window("Nextech Main Window").Dialog("Billing Information").Dialog("NexTech Practice").WinButton("OK").Click


	
Call SelectNewPayRefAdj()
WinMenu("ContextMenu").Select "Create a New &Payment"
wait(.5)


'==============================

IsPresent = Window("Nextech Main Window").Dialog("New Payment dialog").Check (CheckPoint("New Payment for dialog"))
Call RecordResults(IsPresent, "Billing: New Pay/Ref/Adj button clicked")	

' Click the Payment Category ellipse button

Window("Nextech Main Window").Dialog("New Payment dialog").WinButton("Payment Cat Ellipse").Click
IsPresent = Dialog("Payment Categories").Check (CheckPoint("Payment Categories_2"))
Call RecordResults(IsPresent, "Billing: Payment Category ellipse button clicked")	
Dialog("Payment Categories").WinButton("Close").Click

' Click the Description ellipse button

Window("Nextech Main Window").Dialog("New Payment dialog").WinButton("Description Ellipse").Click
IsPresent = Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box_2"))
Call RecordResults(IsPresent, "Billing: New Pay/Ref/Adj button clicked")	
Dialog("Edit Combo Box").WinButton("Close").Click

' Click the Cash Drawer ellipse button

Window("Nextech Main Window").Dialog("New Payment dialog").WinButton("Cash Drawer Ellipse").Click
IsPresent = Dialog("Edit Cash Drawer Sessions").Check (CheckPoint("Edit Cash Drawer Sessions"))
Call RecordResults(IsPresent, "Billing: New Pay/Ref/Adj button clicked")	
Dialog("Edit Cash Drawer Sessions").WinButton("Close").Click

' Click the Card Type ellipse button (no longer present with ICCP)

'Window("Nextech Main Window").Dialog("New Payment dialog").WinButton("Card Type Ellipse").Click
'IsPresent = Dialog("Edit Charge Cards").Check (CheckPoint("Edit Charge Cards"))
'Call RecordResults(IsPresent, "Billing: New Pay/Ref/Adj button clicked")	
'Dialog("Edit Charge Cards").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("New Payment dialog").WinButton("Cancel").Click


' Create a New Pre-Payment

Call SelectNewPayRefAdj()
WinMenu("ContextMenu").Select "Create a New Pr&e-Payment"
wait(.5)
IsPresent = Window("Nextech Main Window").Dialog("New Pre-Payment dialog").Check (CheckPoint("New Pre-Payment dialog"))
Call RecordResults(IsPresent, "Billing: New Pre-Payment dialog opened")

		' Click the Payment Category ellipse button
		
Window("Nextech Main Window").Dialog("New Pre-Payment dialog").WinButton("Payment Category ellipse").Click
IsPresent = Dialog("Payment Categories").Check (CheckPoint("Payment Categories_3"))
Call RecordResults(IsPresent, "Billing: Payment Category ellipse clicked")	
Dialog("Payment Categories").WinButton("Close").Click
		
		' Click the Description ellipse button
		
Window("Nextech Main Window").Dialog("New Pre-Payment dialog").WinButton("Description ellipse").Click
IsPresent = Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box_3"))
Call RecordResults(IsPresent, "Billing: Description ellipse clicked")
Dialog("Edit Combo Box").WinButton("Close").Click
		
		' Click the Cash Drawer ellipse button
		
Window("Nextech Main Window").Dialog("New Pre-Payment dialog").WinButton("Cash Drawer ellipse").Click
IsPresent = Dialog("Edit Cash Drawer Sessions").Check (CheckPoint("Edit Cash Drawer Sessions_2"))
Call RecordResults(IsPresent, "Billing: Cash Drawer ellipse clicked")
Dialog("Edit Cash Drawer Sessions").WinButton("Close").Click

' Click the Card Type ellipse button (removed, no longer present when ICCP enabled)
	
'Window("Nextech Main Window").Dialog("New Pre-Payment dialog").WinButton("Card Type ellipse").Click
'IsPresent = Dialog("Edit Charge Cards").Check (CheckPoint("Edit Charge Cards_2"))
'Call RecordResults(IsPresent, "Billing: Card Type ellipse clicked")
'Dialog("Edit Charge Cards").WinButton("Cancel").Click
'
Window("Nextech Main Window").Dialog("New Pre-Payment dialog").WinButton("Cancel").Click
	
	' Create a New Adjustment
Call SelectNewPayRefAdj()
WinMenu("ContextMenu").Select "Create a New &Adjustment"
wait(.5)
IsPresent = Window("Nextech Main Window").Dialog("New Adjustment for dialog").Check (CheckPoint("New Adjustment for dialog checkpoint"))
Call RecordResults(IsPresent, "Billing: New Adjustment dialog opened")
	
		' Click the Adjustment Category ellipse button
		
Window("Nextech Main Window").Dialog("New Adjustment for dialog").WinButton("Adjustment Category ellipse").Click
IsPresent = Dialog("Payment Categories").Check (CheckPoint("Payment Categories_4"))
Call RecordResults(IsPresent, "Billing: Adjustment Category ellipse")
Dialog("Payment Categories").WinButton("Close").Click

		' Click the Description ellipse button
		
		
Window("Nextech Main Window").Dialog("New Adjustment for dialog").WinButton("Description ellipse").Click
IsPresent = Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box_4"))
Call RecordResults(IsPresent, "Billing: Description ellipse")
Dialog("Edit Combo Box").WinButton("Close").Click

		' Click the Filter button
		
Window("Nextech Main Window").Dialog("New Adjustment for dialog").WinButton("Filter").Click
IsPresent = Dialog("Input").Static("Enter a list of required").Check (CheckPoint("Enter a list of required keywords_7"))
Call RecordResults(IsPresent, "Billing: Filter opened")
Dialog("Input").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("New Adjustment for dialog").WinButton("Cancel").Click

	' Create a New Refund	
	
Call SelectNewPayRefAdj()
WinMenu("ContextMenu").Select "Create a New &Refund"	
	
wait(.5)
IsPresent = Window("Nextech Main Window").Dialog("Available Payments").Check (CheckPoint("Available Payments"))
Call RecordResults(IsPresent, "Billing: Available Payments dialog opened")

Window("Nextech Main Window").Dialog("Available Payments").WinButton("OK").Click

'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("In order to automatically process this refund via credit card processing, you must choose the payment you are refunding. Are you sure you want to continue without choosing a payment to refund?"))
'IsPresent = VerifyPopups("NexTech Practice", "In order to automatically process this refund via credit card processing, you must choose the payment you are refunding\.\nAre you sure you want to continue without choosing a payment to refund\?")
'Call RecordResults(IsPresent, "Billing: Available Payments popup message")

If Dialog("NexTech Practice").WinButton("Yes").Exist Then
	Dialog("NexTech Practice").WinButton("Yes").Click
End If


If Window("Nextech Main Window").Dialog("New Refund dialog").Exist (2) Then
	Call RecordResults(IsPresent, "Billing: New Refund dialog opened")
Else
	Call RecordResults(IsPresent, "Billing: New Refund dialog failed to opened")
End If

		' Click the Refund Category ellipse button
		
Window("Nextech Main Window").Dialog("New Refund dialog").WinButton("Refund Category ellipse").Click
IsPresent = Dialog("Payment Categories").Check (CheckPoint("Payment Categories_5"))
Call RecordResults(IsPresent, "Billing: New Refund dialog opened")
Dialog("Payment Categories").WinButton("Close").Click
		
		' Click the Description ellipse button

Window("Nextech Main Window").Dialog("New Refund dialog").WinButton("Description ellipse").Click
IsPresent = Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box_5"))
Call RecordResults(IsPresent, "Billing: Description ellipse")
Dialog("Edit Combo Box").WinButton("Close").Click
		
		' Click the Cash Drawer ellipse button
		
Window("Nextech Main Window").Dialog("New Refund dialog").WinButton("Cash Drawer ellipse").Click
IsPresent = Dialog("Edit Cash Drawer Sessions").Check (CheckPoint("Edit Cash Drawer Sessions_3"))
Call RecordResults(IsPresent, "Billing: Cash Drawer ellipse")
Dialog("Edit Cash Drawer Sessions").WinButton("Close").Click
Window("Nextech Main Window").Dialog("New Refund dialog").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("Billing Information").WinButton("Cancel").Click

' Clck Preview History

Window("Nextech Main Window").Window("Patients").WinButton("Preview History").Click
IsPresent = Window("Nextech Main Window").Dialog("Print History Filter").Check (CheckPoint("Print History Filter"))
Call RecordResults(IsPresent, "Billing: Print History filter")
Window("Nextech Main Window").Dialog("Print History Filter").WinButton("OK").Click
wait(2)


'***** Figure this out from the Preview History button click


'IsPresent = Window("Nextech Main Window").Window("Patient Fin History (PP)").Static("Total:10").Check (CheckPoint("Total:10"))
Call RecordResults(IsPresent, "Billing: Patient Financial History report preview")
Window("Nextech Main Window").Window("Patient Fin History (PP)").Close




'***********************************************************

' Click Preview Statement

Window("Nextech Main Window").Window("Patients").WinButton("Preview Statement").Click
wait(2)
IsPresent = Window("Nextech Main Window").Dialog("Statements").Check (CheckPoint("Statements"))
Call RecordResults(IsPresent, "Billing: Statement")
	
	' Click the Configuration button
	
Window("Nextech Main Window").Dialog("Statements").WinButton("Configuration").Click
IsPresent = Dialog("Configure Statement").Check (CheckPoint("Configure Statement"))
Call RecordResults(IsPresent, "Billing: Configure Statement")

		' Click the Configure AR Notes button
		
Dialog("Configure Statement").WinButton("Configure AR Notes").Click
IsPresent = Dialog("Configure Statement AR").Check (CheckPoint("Configure Statement AR Notes"))
Call RecordResults(IsPresent, "Billing: Configure Statement AR Notes")
Dialog("Configure Statement AR").WinButton("Cancel").Click
Dialog("Configure Statement").WinButton("Cancel").Click

		' Click the Preview button		
		
Window("Nextech Main Window").Dialog("Statements").WinButton("Preview").Click
IsPresent = Window("Nextech Main Window").Window("Patient Statement Detailed").Static("9 of 9").Check (CheckPoint("9 of 9"))
Call RecordResults(IsPresent, "Billing: Display statement")
Window("Nextech Main Window").Window("Patient Statement Detailed").Close		
		
' Click Apply Manager (be sure to select the Applies to Bills tab)

Window("Nextech Main Window").Window("Patients").WinButton("Apply Manager").Click
IsPresent = Window("Nextech Main Window").Dialog("Apply Manager").Check (CheckPoint("Apply Manager"))
Call RecordResults(IsPresent, "Billing: Apply Manager")

' Click Unapply item
Window("Nextech Main Window").Dialog("Apply Manager").WinButton("Unapply Item").Click
'IsPresent = Dialog("NexTech Practice").Static("Please select a service").Check (CheckPoint("Please make a selection before unapplying."))
IsPresent = VerifyPopups("NexTech Practice", "Please make a selection before unapplying.")

Call RecordResults(IsPresent, "Billing: Popup to make selection")
Dialog("NexTech Practice").WinButton("OK").Click

Window("Nextech Main Window").Dialog("Apply Manager").WinButton("Apply New").Click
wait(.5)
IsPresent = Dialog("New Payment for VAC32,").Check (CheckPoint("New Payment for VAC32, Patient32"))
Call RecordResults(IsPresent, "Billing: New Payment dialog opens")
Dialog("New Payment for VAC32,").WinButton("Cancel").Click

Window("Nextech Main Window").Dialog("Apply Manager").WinButton("Close").Click
	
' Click the Search Button

Window("Nextech Main Window").Window("Patients").WinButton("Button").Click
IsPresent = Window("Nextech Main Window").Dialog("Search Billing Tab").Check (CheckPoint("Search Billing Tab"))
Call RecordResults(IsPresent, "Billing: Search Billing Tab dialog opens")
Window("Nextech Main Window").Dialog("Search Billing Tab").Close

' Click the Statement Notes ellipse button
Window("Nextech Main Window").Window("Patients").WinButton("General 1: Ellipsis WinButton").Click
IsPresent = Window("Nextech Main Window").Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box_6"))
Call RecordResults(IsPresent, "Billing: Edit Combo Box dialog opens")
Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click


' Close the Billing Info dialog

'Window("Nextech Main Window").Dialog("Billing Information").WinButton("Cancel").Click

End Sub

Sub VerifyInsurance()
	
	' Click Edit Insurance List
	
	Window("Nextech Main Window").Window("Patients").WinButton("Edit Insurance List").Click
	If Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Insurance Company Information").Exist (5) Then
		Call RecordResults(True, "Insurance: Edit Insurance Company Information opened.")
	Else
		Call RecordResults(False, "Insurance: Edit Insurance Company Information DID NOT open.")
	End If
	
		' Click Add Company
		' Practice stopped working here ***
		
	Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Add Company").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Input").Static("Enter a name for this").Check (CheckPoint("Enter a name for this Insurance Company:"))
	Call RecordResults(IsPresent, "Insurance: Add Company button click dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Input").WinButton("Cancel").Click
		
		' Click Add Contact
		
	Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Add Contact").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Input").Static("Enter a name for this").Check (CheckPoint("Enter a last name for this contact:"))
	Call RecordResults(IsPresent, "Insurance: Add Contact button click dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Input").WinButton("Cancel").Click


		' Click Manage Contacts
			
	Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Manage Contacts").Click
'	IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("NexTech Practice").Static("Be careful when using").Check (CheckPoint("Be careful when using this utility. You can change large amounts of data at once, and your changes cannot be undone"))
'	Call RecordResults(IsPresent, "Insurance: Be careful popup")	
'	Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("NexTech Practice").WinButton("OK").Click
'	
	IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Manage Insurance Contacts").Check (CheckPoint("Manage Insurance Contacts"))
	Call RecordResults(IsPresent, "Insurance: Manage Contacts click dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Manage Insurance Contacts").WinButton("Close").Click
	
	
		' Click Advanced Pay./Adj. Setup
		
	Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Advanced Pay. / Adj. Setup").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Advanced Insurance Payment").Check (CheckPoint("Advanced Insurance Payment / Adjustment Description Setup"))
	Call RecordResults(IsPresent, "Insurance: Manage Contacts click dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Advanced Insurance Payment").WinButton("Close").Click
	
'=========================
		
		' Click Claim Provider Setup (In Use)
		
		
	Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Claim Provider Setup (In").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Claim Providers Setup").Check (CheckPoint("Claim Providers Setup"))
	Call RecordResults(IsPresent, "Insurance: Claim Providers Setup click dialog opened")	

			'Click Advanced Setup
			'Nextech Stopped Working Here ***
			
	Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Claim Providers Setup").WinButton("Advanced Setup").Click
	IsPresent = Dialog("Advanced Claim Provider").Check (CheckPoint("Advanced Claim Provider Setup"))
	Call RecordResults(IsPresent, "Insurance: Advanced Setup click dialog opened")	
	Dialog("Advanced Claim Provider").WinButton("Close").Click
	Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Claim Providers Setup").WinButton("Cancel").Click
			
		' Click Edit Notes codes
		
	Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Edit Code Notes").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Insurance Code Notes").Check (CheckPoint("Insurance Code Notes"))
	Call RecordResults(IsPresent, "Insurance: Edit Notes Codes click dialog opened")	


			' Click Advanced
			
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Insurance Code Notes").WinButton("Advanced").Click
		IsPresent = Dialog("Insurance Company Code").Check (CheckPoint("Insurance Company Code Notes"))
		Call RecordResults(IsPresent, "Insurance: Advanced click dialog opened")	
		Dialog("Insurance Company Code").WinButton("Close").Click
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Insurance Code Notes").WinButton("Close").Click
			
		' Click Configure Payer IDs Per Location
		
		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Configure Payer IDs Per").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Configure Payer IDs Per").Check (CheckPoint("Configure Payer IDs Per Location"))
		Call RecordResults(IsPresent, "Insurance: Confugure Payer IDs Per Location click dialog opened")	
		
			' Click Payer ID ellipse
			Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Configure Payer IDs Per").WinButton("EligiblePayerIDEllipse").Click
			IsPresent = Dialog("Edit Payer List").Check (CheckPoint("Edit Payer List"))
			Call RecordResults(IsPresent, "Insurance: Eligible Payer ID ellipse button click dialog opened")	
			Dialog("Edit Payer List").WinButton("Close").Click
			
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Configure Payer IDs Per").WinButton("Close").Click
		wait (.5)
		
		' Click Edit Default Pay Group Information
		' Practice stopped working here ***		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Edit Default Pay Group").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Default Pay Group").Check (CheckPoint("Edit Default Pay Group Information"))
		Call RecordResults(IsPresent, "Insurance: Edit Default Pay Group Info button click dialog opened")	
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Default Pay Group").WinButton("Cancel").Click
		
		
		' Click CLIA Setup
		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("CLIA Setup").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("CLIA Number Setup").Check (CheckPoint("CLIA Number Setup"))
		Call RecordResults(IsPresent, "Insurance: CLIA Setup button click dialog opened")	
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("CLIA Number Setup").WinButton("Configure Service Codes").Click
		
			' Click Configure Service Codes
			
			IsPresent = Dialog("CLIA Service Code Setup").Check (CheckPoint("CLIA Service Code Setup"))
			Call RecordResults(IsPresent, "Insurance: CLIA Setup button click dialog opened")	
			Dialog("CLIA Service Code Setup").WinButton("Cancel").Click
			
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("CLIA Number Setup").WinButton("Cancel").Click
			
		' Click HCFA Box 24J
		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("HCFA Box 24J").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Box 24I / Box 24J").Check (CheckPoint("Edit Box 24I / Box 24J"))
		Call RecordResults(IsPresent, "Insurance: HCFA Box 24J button click dialog opened")	

			' Click Advanced Features
			Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Box 24I / Box 24J").WinButton("Advanced Features").Click
			IsPresent = Dialog("Advanced Insurance ID").Check (CheckPoint("Advanced Insurance ID Editing_ From Edit Box 24I_24J dialog"))
			Call RecordResults(IsPresent, "Insurance: Advanced Insurance ID Editing button click dialog opened")	
			Dialog("Advanced Insurance ID").WinButton("Close").Click
			
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Box 24I / Box 24J").WinButton("Close").Click

		' Click HCFA Group #

		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("HCFA Group #").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Insurance Groups").Check (CheckPoint("Insurance Groups"))
		Call RecordResults(IsPresent, "Insurance: HCFA Group # button click dialog opened")	
			
			' Click Advanced Features
			
			Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Insurance Groups").WinButton("Advanced Features").Click
			IsPresent = Dialog("Advanced Insurance ID").Check (CheckPoint("Advanced Insurance ID Editing_From Ins Groups dialog"))
			Call RecordResults(IsPresent, "Insurance: Advanced Insurance ID Editing button click dialog opened")	
			Dialog("Advanced Insurance ID").WinButton("Close").Click
		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Insurance Groups").WinButton("Close").Click
		
		' Click Network ID
		
		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Network ID").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Network ID").Check (CheckPoint("Edit Network ID"))
		Call RecordResults(IsPresent, "Insurance: Network ID button click dialog opened")	
		
			' Click Advanced Features
			
			Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Network ID").WinButton("Advanced Features").Click
			IsPresent = Dialog("Advanced Insurance ID").Check (CheckPoint("Advanced Insurance ID Editing_From Edit Network ID dialog"))
			Call RecordResults(IsPresent, "Insurance: Advanced Insurance ID Editing button click dialog opened")	
			Dialog("Advanced Insurance ID").WinButton("Close").Click
			
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Network ID").WinButton("Close").Click
		
		
		' Click HCFA Facility ID (32b)
		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("HCFA Facility ID (32b)").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Facility ID").Check (CheckPoint("Facility ID"))
		Call RecordResults(IsPresent, "Insurance: HCFA Facility ID (32b) button click dialog opened")	
		
			' Click Advanced Features
			Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Facility ID").WinButton("Advanced Features").Click
			IsPresent = Dialog("Advanced Facility ID Editing").Check (CheckPoint("Advanced Facility ID Editing"))
			Call RecordResults(IsPresent, "Insurance: Advanced Features button click dialog opened")	
			
		Dialog("Advanced Facility ID Editing").WinButton("Close").Click
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Facility ID").WinButton("Close").Click
		
		' Click HCFA Box 31
		
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("HCFA Box 31").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Box 31").Check (CheckPoint("Edit Box 31"))
		Call RecordResults(IsPresent, "Insurance: HCFA Box 31 button click dialog opened")	
	
			' Click Advanced Features
			Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Box 31").WinButton("Advanced Features").Click
			IsPresent = Dialog("Advanced Insurance ID").Check (CheckPoint("Advanced Insurance ID Editing"))
			Call RecordResults(IsPresent, "Insurance: Advanced Features button click dialog opened")	
			Dialog("Advanced Insurance ID").WinButton("Close").Click
			
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit Box 31").WinButton("Close").Click
		
		' Click UB92 Box51
		
				
		Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("UB92 Box 51").Click
		IsPresent = Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit UB92 Box 51").Check (CheckPoint("Edit UB92 Box 51"))
			Call RecordResults(IsPresent, "Insurance: HCFA Facility ID (32b) button click dialog opened")	
		
			' Click Advanced Features
			Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit UB92 Box 51").WinButton("Advanced Features").Click
			IsPresent = Dialog("Advanced Insurance ID").Check (CheckPoint("Advanced Insurance ID Editing_2"))
			Call RecordResults(IsPresent, "Insurance: Advanced Features button click dialog opened")	
			Dialog("Advanced Insurance ID").WinButton("Close").Click
			
		Window("Nextech Main Window").Dialog("Edit Insurance Company").Dialog("Edit UB92 Box 51").WinButton("Close").Click
		
	' Close Edit Insurance Company Information dialog
	Window("Nextech Main Window").Dialog("Edit Insurance Company").WinButton("Close").Click
	
	' Click Edit Service Notes
	Window("Nextech Main Window").Window("Patients").WinButton("Edit Service Notes").Click
	IsPresent = Window("Nextech Main Window").Dialog("Insurance Code Notes").Check (CheckPoint("Insurance Code Notes_2"))
	Call RecordResults(IsPresent, "Insurance: Edit Service Notes button click dialog opened")	
	
		' Click Advanced
		
		Window("Nextech Main Window").Dialog("Insurance Code Notes").WinButton("Advanced").Click
		IsPresent = Dialog("Insurance Company Code").Check (CheckPoint("Insurance Company Code Notes_2"))
		Call RecordResults(IsPresent, "Insurance: Advanced button click dialog opened")	
		Dialog("Insurance Company Code").WinButton("Close").Click
		
	Window("Nextech Main Window").Dialog("Insurance Code Notes").WinButton("Close").Click
	
	' Click Edit Referrals
	
	Window("Nextech Main Window").Window("Patients").WinButton("Edit Referrals").Click
	IsPresent = Window("Nextech Main Window").Dialog("Insurance Referrals").Check (CheckPoint("Insurance Referrals"))
	Call RecordResults(IsPresent, "Insurance: Edit Referrals button click dialog opened")	
		
		' Click Add
		Window("Nextech Main Window").Dialog("Insurance Referrals").WinButton("Add").Click
		IsPresent = Dialog("Enter New Insurance Referral").Check (CheckPoint("Enter New Insurance Referral"))
		Call RecordResults(IsPresent, "Insurance: Edit Referrals Add button click dialog opened")	
		Dialog("Enter New Insurance Referral").WinButton("Cancel").Click
		
	Window("Nextech Main Window").Dialog("Insurance Referrals").WinButton("Close").Click
	

	' Click Deductible/OOP	
	Window("Nextech Main Window").Window("Patients").WinButton("Deductible/OOP").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Deductible/Out Of").Check (CheckPoint("Edit Deductible/Out Of Pocket Amounts"))
	Call RecordResults(IsPresent, "Insurance: Deductible/OOP button click dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Deductible/Out Of").WinButton("Cancel").Click


	

End Sub

Sub VerifyQuotes()
	
' Click the New Quote button
	
	Window("Nextech Main Window").Window("Patients").WinButton("Quotes: New Quote button").Click
	wait(.1)
	If Window("Nextech Main Window").Dialog("Quote Information").Exist (3) Then
		Call RecordResults(True, "Quote Information dialog opened")
	Else
		Call RecordResults(False, "**** Quote Information dialog failed to open")
	End If
	
	' Click the Edit Text button
	
	Window("Nextech Main Window").Dialog("Quote Information").WinButton("Edit  Text").Click
	IsPresent = Window("Nextech Main Window").Dialog("Quote Information").Dialog("Quote Notes").Check (CheckPoint("Quote Notes"))
	Call RecordResults(IsPresent, "Quotes: Edit Text button click dialog opened")	
	
	' Click the Adminstrator button
	
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("Quote Notes").WinButton("Administrator").Click
	IsPresent = Dialog("Quote Administrator").Check (CheckPoint("Quote Administrator"))
	Call RecordResults(IsPresent, "Quotes: Administrator button click dialog opened")	
	Dialog("Quote Administrator").WinButton("Cancel").Click
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("Quote Notes").WinButton("Cancel").Click
	
	' Click the Make Default button
	
		Window("Nextech Main Window").Dialog("Quote Information").WinButton("Make Default").Click
	IsPresent = Window("Nextech Main Window").Dialog("Quote Information").Dialog("NexTech Practice").Check (CheckPoint("NexTech Practice_3"))
	Call RecordResults(IsPresent, "Quotes: Make Default button click dialog opened")	
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("NexTech Practice").WinButton("No").Click
	
	' Click the Quote info button
	
	Window("Nextech Main Window").Dialog("Quote Information").WinButton("Apply Discount to All").Click
	IsPresent = Window("Nextech Main Window").Dialog("Quote Information").Dialog("NexTech Practice").Check (CheckPoint("NexTech Practice_4"))
	Call RecordResults(IsPresent, "Quotes: Quote Info button click dialog opened")	
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("NexTech Practice").WinButton("OK").Click
	
	' Click the Modify Discounts button
	
	Window("Nextech Main Window").Dialog("Quote Information").WinButton("Modify Discounts").Click
	IsPresent = Window("Nextech Main Window").Dialog("Quote Information").Dialog("Modify Discounts").Check (CheckPoint("Modify Discounts"))
	Call RecordResults(IsPresent, "Quotes: Modify discounts button click dialog opened")	
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("Modify Discounts").WinButton("Cancel").Click
	
	' Click the Show Suggestions button
	Window("Nextech Main Window").Dialog("Quote Information").WinButton("Show Suggestions").Click
	IsPresent = Window("Nextech Main Window").Dialog("Quote Information").Dialog("Suggested Sales").Check (CheckPoint("Suggested Sales"))
	Call RecordResults(IsPresent, "Quotes: Suggested sales button click dialog opened")	
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("Suggested Sales").Close
	
	'Click the Place of service ellipse button
	
	Window("Nextech Main Window").Dialog("Quote Information").WinButton("WinButton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Quote Information").Dialog("Place of Service Designations").Check (CheckPoint("Place of Service Designations"))
	Call RecordResults(IsPresent, "Quotes: Place of service ellipse button click dialog opened")	
	
	'Click the Add button
	
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("Place of Service Designations").WinButton("Add").Click
	IsPresent = Dialog("Input").Check (CheckPoint("Input_5"))
	Call RecordResults(IsPresent, "Quotes: Add button click dialog opened")	
	Dialog("Input").WinButton("Cancel").Click
	Window("Nextech Main Window").Dialog("Quote Information").Dialog("Place of Service Designations").WinButton("Close").Click
	Window("Nextech Main Window").Dialog("Quote Information").WinButton("Cancel").Click
	
	'-----------------------------------
	
	
	
End Sub

Sub VerifyAppts()
	'------------------
	' Click the Recalls button

	Window("Nextech Main Window").Window("Patients").WinButton("Appts.: Recalls Button").Click
	IsPresent = Window("Nextech Main Window").Dialog("Recalls Needing Attention").Check (CheckPoint("Recalls Needing Attention - VAC32, Patient32"))
	Call RecordResults(IsPresent, "Appts: Recalls button click dialog opened")	
	
	' Click the Create Merge Group button
	
	Window("Nextech Main Window").Dialog("Recalls Needing Attention").WinButton("Create Merge Group").Click
	IsPresent = Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice").Check (CheckPoint("Practice_2"))
	Call RecordResults(IsPresent, "Appts: Create Merge Group button click dialog opened")	
	Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice").WinButton("OK").Click

	'Verify Already Exists popup is present and then close it
	If Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("NexTech Practice").Exist (1) Then
		Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("NexTech Practice").WinButton("OK").Click
	End If
	
	
	' click the merge to word button
	Window("Nextech Main Window").Dialog("Recalls Needing Attention").WinButton("Merge To Word").Click
	wait(2)	
	'Handle Practice new Word document popup 
	IsPresent = Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice").Check (CheckPoint("Practice_3"))
	Call RecordResults(IsPresent, "Appts: merge to word click dialog opened")	

	'Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice").WinButton("OK").Click
	'Handle the Are you sure popup
	'Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice").Static("A new Word document will").Check CheckPoint("NewWordDocAreYouSurePopup")
	'Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice").WinButton("OK").Click
'
	If Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("regexpwndtitle:=Practice").Static("regexpwndtitle:=Please adjust your filters so that there are recalls in the list\.").Exist (1) Then
		Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click
		Call RecordResults(True, "Adjust filters popup closed")
	End If
	wait (.1)

'	Call RecordResults(IsPresent, "Recalls: Please adjust filters popup")
'	
	' click the Create New Recall button	
	Window("Nextech Main Window").Dialog("Recalls Needing Attention").WinButton("Create New Recall").Click
	IsPresent = Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice - Create Patient").Check (CheckPoint("Practice - Create Patient Recall for VAC32, Patient32 (43)"))
	Call RecordResults(IsPresent, "Appts: Create New Recall button click dialog opened")	
	Window("Nextech Main Window").Dialog("Recalls Needing Attention").Dialog("Practice - Create Patient").WinButton("Cancel").Click
	Window("Nextech Main Window").Dialog("Recalls Needing Attention").WinButton("Close").Click
	
'------------------
	
End Sub

Sub VerifyNotes()


' Click Add Blank Note button (TBD)

' Click Add Macro button
	
	Window("Nextech Main Window").Window("Patients").WinButton("text:=A&dd Macro").Click
	IsPresent = Window("Nextech Main Window").Dialog("Add Macro").Check (CheckPoint("Add Macro"))
	Call RecordResults(IsPresent, "Notes: Add Macro button click dialog opened")	
	Window("Nextech Main Window").Dialog("Add Macro").WinButton("text:=Cancel").Click
	' Practice stopped working here ***
	' Click Edit Macros button
	
	Window("Nextech Main Window").Window("Patients").WinButton("Edit Macros").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Macros").Check (CheckPoint("Edit Macros"))
	Call RecordResults(IsPresent, "Notes: Edit Macro button click dialog opened")	

	Window("Nextech Main Window").Dialog("Edit Macros").WinButton("text:=Add Macro").Click
	IsPresent = Dialog("Input").Check (CheckPoint("Input_4"))
	Call RecordResults(IsPresent, "Notes: Add macro popup button click dialog opened")	
	
	Dialog("Input").WinButton("text:=&Cancel").Click
	'Window("Nextech Main Window").Dialog("Edit Macros").WinButton("Cancel").Click
	' comment out the above line once the descriptive object is declared for the cancel button.  Should this be local or global?
	Window("Nextech Main Window").Dialog("Edit Macros").WinButton("text:=Cancel").Click

	
	Window("Nextech Main Window").Window("Patients").WinButton("Edit Categories_2").Click
	IsPresent=Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Check (CheckPoint("Note / Follow-Up Categories_4"))
	Call RecordResults(IsPresent, "Notes: Note follow up button click dialog opened")	
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("text:=Clos&e").Click
	
	' Click Search Notes button
	
	Window("Nextech Main Window").Window("Patients").WinButton("Search Notes").Click
	IsPresent = Window("Nextech Main Window").Dialog("Search Notes For: VAC32,").Check (CheckPoint("Search Notes For: VAC32, Patient32"))
	Call RecordResults(IsPresent, "Notes: search notes button click dialog opened")	
	Window("Nextech Main Window").Dialog("Search Notes For: VAC32,").WinButton("text:=Close").Click


'========================	
	
End Sub

Sub VerifyCustom()

	' Click skin type ellipse button
	Window("Nextech Main Window").Window("Patients").WinButton("CList1ellipsebutton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box"))
	Call RecordResults(IsPresent, "Custom: Custom 1 ellipse dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click
	
	' Click Custom List 2 ellipse button
	
	Window("Nextech Main Window").Window("Patients").WinButton("CList2ellipsebutton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box"))
	Call RecordResults(IsPresent, "Custom: Custom 2 ellipse dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click

	' Click Custom List 3 ellipse button
	Window("Nextech Main Window").Window("Patients").WinButton("CList3ellipsebutton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box"))
	Call RecordResults(IsPresent, "Custom: Custom 3 ellipse dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click
	
	' Click Custom List 4 ellipse button
	Window("Nextech Main Window").Window("Patients").WinButton("CList4ellipsebutton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box"))
	Call RecordResults(IsPresent, "Custom: Custom 4 ellipse dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click
	
	' Click Custom List 5 ellipse button
	Window("Nextech Main Window").Window("Patients").WinButton("CList5ellipsebutton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box"))
	Call RecordResults(IsPresent, "Custom: Custom 5 ellipse dialog opened")		
	Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click
	
	' Click Custom List 6 ellipse button
	Window("Nextech Main Window").Window("Patients").WinButton("CList6ellipsebutton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Combo Box").Check (CheckPoint("Edit Combo Box"))
	Call RecordResults(IsPresent, "Custom: Custom 6 ellipse dialog opened")	
	Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click
	
		
End Sub

Sub VerifyMedications()
	
'Clicking on the first ellipsis Button in the Medication tab
	
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=\.\.\.", "index:=0").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").Exist Then
	Reporter.ReportEvent micPass, "Edit Medication List is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Medication List is not displayed"
End If

Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinMenu("menuobjtype:=3").Select "Import Medication"

'Verifying the Meidcation Import Dialog
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Type "alv"

'does dialog open

If Dialog("regexpwndtitle:=Medication Import").Exist Then
	Reporter.ReportEvent micPass, "Medical Import dialog is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "<<< Medical Import dialog List was not displayed!>>>"
End If



Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Search").Click
wait 2
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Help Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=&Help").Click

If Dialog("regexpwndtitle:=NexTech Practice User Manual").Exist Then
	Reporter.ReportEvent micPass, "NexTech Practice User Manual dialog is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "<<< NexTech Practice User Manual dialog List was not displayed!>>>"
End If

Window("NexTech Practice User").Close

'Clicking on Inactive Medications button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Inactive Medications").Click

If Dialog("regexpwndtitle:=Inactive Medications").Exist Then
	Reporter.ReportEvent micPass, "Inactive Medications dialog is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "<<< Inactive Medications dialog List was not displayed!>>>"
End If

Dialog("regexpwndtitle:=Inactive Medications","nativeclass:=#32770").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Latin Notation button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Edit Latin Notation").Click

If Dialog("regexpwndtitle:=Latin Prescription Settings").Exist Then
	Reporter.ReportEvent micPass, "Latin Prescription Settings dialog is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "<<< Latin Prescription Settings dialog was not displayed!>>>"
End If


Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Delete").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Close").Click

''Clicking on the Interactions Button
'Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Interactions").Click
'If Window("Nextech Main Window").Dialog("regexpwndtitle:=Drug Interactions.*").Exist Then
'	Reporter.ReportEvent micPass, "Interactions Dialog is open",""
'Else
'	Reporter.ReportEvent micFail, "Test Failed", "Interaction Dialog is Missing"	
'End If
'
''Clicking on Configure Severity Filters
'Window("Nextech Main Window").Dialog("regexpwndtitle:=Drug Interactions.*").WinButton("regexpwndtitle:=Configure Severity Filters").Click
'IsPresent = Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Drug Interactions.*").Dialog("regexpwndtitle:=Configure Drug.*").Exist
'Call RecordResults(IsPresent, "Configure Drug Interaction Severity Filters Dialog is Present")
'Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Drug Interactions.*").Dialog("regexpwndtitle:=Configure Drug.*").WinButton("regexpwndtitle:=Cancel").Click
'Window("Nextech Main Window").Dialog("regexpwndtitle:=Drug Interactions.*").WinButton("regexpwndtitle:=Close").Click
'
''Clicking on Med History Button
'Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Med History").Click
'If Window("Nextech Main Window").Dialog("text:=Medication History.*").Exist Then
'	Reporter.ReportEvent micPass, "Medication History Dialog appears",""
'Else
'	Reporter.ReportEvent micFail, "Test Failed", "Medication History Dialog missing"	
'End If
'Window("Nextech Main Window").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=Request Med History").Click
'IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
'Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
'Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
'
''Clicking on View Eligibility Details
'Window("Nextech Main Window").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=View Eligibility Details").Click
'If Dialog("regexpwndtitle:=Eligibility Details").Exist Then
'	Reporter.ReportEvent micPass, "Eligibility Details Dialog displays",""
'Else
'	Reporter.ReportEvent micFail, "Test Failed", "Eligibility Details Dialog not displayed"	
'End If
'Dialog("regexpwndtitle:=Eligibility Details").WinButton("regexpwndtitle:=Close").Click
'
''Clicking on Import into Current Meds
'Window("Nextech Main Window").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=Import into Current Meds").Click
'IsPresent = Dialog("regexpwndtitle:=No Historical Medications Checked").Exist
'Call RecordResults(IsPresent, "No Historical Medications Checked Dialog is Present")
'Dialog("regexpwndtitle:=No Historical Medications Checked").WinButton("regexpwndtitle:=OK").Click
'Window("Nextech Main Window").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=Close").Click
'
'Clicking on the ellipsis button on the bottom
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=\.\.\.", "index:=1").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").Exist Then
	Reporter.ReportEvent micPass, "Edit Allergy List Dialog displays",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Allergy List Dialog is not displayed"	
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinMenu("menuobjtype:=3").Select "Import Allergy"
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Edit Allergy List").Dialog("regexpwndtitle:=Allergy Import").Exist
Call RecordResults(IsPresent, "Allergy Import Dialog is Present")

'Try to search the searchbar with just 2 letters
Dialog("regexpwndtitle:=Allergy Import").WinEdit("regexpwndclass:=Edit").Type "Po"
Dialog("regexpwndtitle:=Allergy Import").WinButton("regexpwndtitle:=Search").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "This Test passes",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Needs Attention!!!!"
End If
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=Allergy Import").WinEdit("regexpwndclass:=Edit").Set ""

Dialog("regexpwndtitle:=Allergy Import").WinEdit("regexpwndclass:=Edit").Type "Poll"
Dialog("regexpwndtitle:=Allergy Import").WinButton("regexpwndtitle:=Search").Click
Dialog("regexpwndtitle:=Allergy Import").WinButton("regexpwndtitle:=Cancel").Click

Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinMenu("menuobjtype:=3").Select "Add Free Text Allergy"
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")

Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Edit").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Delete").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Update All Allergies").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Close").Click

'Clicking on ellipsis Button in the middle of the screen

Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=\.\.\.", "index:=2").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").Exist Then
	Reporter.ReportEvent micPass, "Edit Medication List is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Medication List is not displayed"
End If

Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinMenu("menuobjtype:=3").Select "Import Medication"
IsPresent = Dialog("regexpwndtitle:=Medication Import").Exist
Call RecordResults(IsPresent, "Medication Import Dialog is Present")

'Try to search the searchbar with just 2 letters
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Type "Al"
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Search").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "This Test is passes",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Needs Attention!!!!"
End If
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Set ""

'Verifying the Meidcation Import Dialog
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Type "alv"
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Search").Click
wait 2
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Help Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("NexTech Practice User").Exist
Call RecordResults(IsPresent, "Nextech practice User Manual is Present")
Window("NexTech Practice User").Close

'Clicking on Inactive Medications button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Inactive Medications").Click
IsPresent = Dialog("regexpwndtitle:=Inactive Medications").Exist
Call RecordResults(IsPresent, "Inactive Medications dialog is Present")
Dialog("regexpwndtitle:=Inactive Medications","nativeclass:=#32770").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Latin Notation button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Edit Latin Notation").Click
IsPresent = Dialog("regexpwndtitle:=Latin Prescription Settings").Exist
Call RecordResults(IsPresent, "Latin Prescription Setting Dialog is Present")

Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Delete").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Close").Click

'Clicking on Write From Quick List

Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Write From Quick List").Click
Window("Nextech Main Window").Window("Patients").ActiveX("NexTech DataList Control_3").WinObject("object class:=Afx:.*:.*").Click

IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").Exist
Call RecordResults(IsPresent, "NexERx User List Config Dialog is Present")

'Cliking on the Add Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Add").Click
Set ab = CreateObject("Mercury.DeviceReplay")

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=0").WinEdit("regexpwndtitle:=Medication Search...").Type "Cipro"
wait 2
ab.PressKey 28
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=1").Click
ab.PressKey 208
ab.PressKey 208
ab.PressKey 28

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").WinEdit("regexpwndclass:=Edit","index:=2").Type "30"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=2").Click
ab.PressKey 81
ab.PressKey 28

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=3").Click
ab.PressKey 50
ab.PressKey 23
ab.PressKey 28

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=4").Click
ab.SendString "oral"
ab.PressKey 28

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=5").Click
ab.SendString "BID"
ab.PressKey 28

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").WinEdit("regexpwndclass:=Edit","index:=4").Type "Test"

Set ab = Nothing

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User.*").Dialog("regexpwndtitle:=Add Quick List Item").WinButton("regexpwndtitle:=OK").Click

'Clicking on Edit Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick.*").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=1").WinObject("object class:=Afx:.*:.*").Click 126,26
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Edit").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Else
	Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexERx User Quick.*").Dialog("regexpwndtitle:=Edit Quick List").WinButton("regexpwndtitle:=Cancel").Click	
End If

'Clicking on Delete Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick.*").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=1").WinObject("object class:=Afx:.*:.*").Click 126,26
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Else
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click
End If


'Clicking on Import Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Import\.\.\.").Click
If Dialog("regexpwndtitle:=Multiple Selection").Exist Then
	Reporter.ReportEvent micPass, "Multiple Selection Dialog displayed",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Multiple Selection Dialog not displayed"
End If
Dialog("regexpwndtitle:=Multiple Selection").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Close").Click

'Clicking on Rx Print Setup
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Rx Print\nSetup").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescription Template Setup").Exist Then
	Reporter.ReportEvent micPass, "Prescription Template Dialog Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Prescription Template Dialog Not Present"
End If

Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Add New Default Template").Click
IsPresent = Dialog("regexpwndtitle:=Select a New Prescription Template").WinButton("regexpwndtitle:=Cancel").Exist
Call RecordResults(IsPresent, "Select a New Prescription Template Dialog is Present")
Dialog("regexpwndtitle:=Select a New Prescription Template").WinButton("regexpwndtitle:=Cancel").Click

Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Remove Selected Default Template").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Favorite Pharmacies
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Edit Favorite\nPharmacies").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Favorite Pharmacies.*").Exist Then
	Reporter.ReportEvent micPass, "Edit Favorite Pharmacies Dialog Box is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Favorite Pharmacies Dialog Box Not Found"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Favorite Pharmacies.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Word Template
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Edit Word\nTemplate").Click
'If word is installed do this
IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Select a template to edit").Exist
If IsPresent Then
	Call RecordResults(IsPresent, "Select a template to Edit Dialog is Present")
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Select a template to edit").WinButton("regexpwndtitle:=Cancel").Click
ElseIf Window("regexpwndtitle:=Nextech \(UFT.....DB\.*\)").Dialog("regexpwndtitle:=NexTech Practice").Exist (2) Then
	Call RecordResults(True, "Looks like word is not installed on this machine")
	Window("regexpwndtitle:=Nextech \(UFT.....DB\.*\)").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
End If


'Window("Nextech Main Window").Dialog("regexpwndtitle:=Select a template to edit").WinButton("regexpwndtitle:=Cancel").Click

''Clicking on Rx Needing Attention Button
'Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Rx Needing\nAttention").Click
'If Dialog("regexpwndtitle:=NexERx License Configuration").Exist Then
'	Dialog("regexpwndtitle:=NexERx License Configuration").WinButton("regexpwndtitle:=&No").Click
'End If
'
'If Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Exist Then
'	Reporter.ReportEvent micPass, "Prescriptions Needing Attention Dialog is Open",""
'Else
'	Reporter.ReportEvent micFail, "Test Failed", "Prescriptions Needing Attention Dialog is Not Open"
'End If
'
''Clicking on Rx Print Setup in Prescription Needing Attention Dialog
'Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescriptions Needing Attention").WinButton("regexpwndtitle:=Rx Print\nSetup").Click
'
'If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").Exist Then
'	Reporter.ReportEvent micPass, "Prescription Template Dialog Present",""
'Else
'	Reporter.ReportEvent micFail, "Test Failed", "Prescription Template Dialog Not Present"
'End If
'
'Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Add New Default Template").Click
'IsPresent = Dialog("regexpwndtitle:=Select a New Prescription Template").WinButton("regexpwndtitle:=Cancel").Exist
'Call RecordResults(IsPresent, "Prescription Template Setup Dialog is Present")
'Dialog("regexpwndtitle:=Select a New Prescription Template").WinButton("regexpwndtitle:=Cancel").Click
'Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Remove Selected Default Template").Click
'IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
'Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
'Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
'Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Cancel").Click
'
''Clicking on Edit Word Template in the Prescription Needing Attention Dialog
'Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescriptions Needing Attention").WinButton("regexpwndtitle:=Edit Word\nTemplate").Click
'IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Select a template to edit").Exist
'Call RecordResults(IsPresent, "Select a template to edit Dialog is Present")
'Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Select a template to edit").WinButton("regexpwndtitle:=Cancel").Click
'Window("Nextech Main Window").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Close
'
''Clicking on View Formulary Information
'Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=View Formulary Information").Click
'IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").Exist
'Call RecordResults(IsPresent, "NexFormulary Information Dialog is Present")
'Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=&Yes").Click
'If Window("Nextech Main Window").Dialog("regexpwndtitle:=NexFormulary Information for.*").Exist Then
'	Reporter.ReportEvent micPass, "NexFormulatory Information Dialog is Present",""
'Else
'	Reporter.ReportEvent micFail, "Test Failed", "NexFormulatory Information Dialog is not Present"
'End If
'
''Clicking on View Eligibility Details
'Window("Nextech Main Window").Dialog("regexpwndtitle:=NexFormulary Information for.*").WinButton("regexpwndtitle:=View Eligibility Details").Click
'IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexFormulary Information for.*").Dialog("regexpwndtitle:=Eligibility Details").Exist
'Call RecordResults(IsPresent, "Eligibility Details Dialog box is Present")
'Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=NexFormulary Information for.*").Dialog("regexpwndtitle:=Eligibility Details").WinButton("regexpwndtitle:=Close").Click
'Window("Nextech Main Window").Dialog("regexpwndtitle:=NexFormulary Information for.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on Add Button in Medication Schedule
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Add").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").Exist Then
	Reporter.ReportEvent micPass, "Medication Schedule Setup Dialog Opens",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Medication Schedule Setup Dialog did not Open"
End If

'Clicking on the Add Details Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Add Detail").Click
IsPresent = Dialog("regexpwndtitle:=Medication Schedule Detail").Exist
Call RecordResults(IsPresent, "Medication Schedule Detail Dialog is Present")

Dialog("regexpwndtitle:=Medication Schedule Detail").WinButton("regexpwndtitle:=OK").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Dialog("regexpwndtitle:=Medication Schedule Detail").WinButton("regexpwndtitle:=Write Prescription").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on the ellipsis Button on the Medication Schedule Detail Dialog
Dialog("regexpwndtitle:=Medication Schedule Detail").WinButton("regexpwndtitle:=\.\.\.").Click
If Dialog("regexpwndtitle:=Edit Medication List").Exist Then
	Reporter.ReportEvent micPass, "Edit Medication List Dialog is open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Medication List Dialog did not Open"
End If

Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Edit Medication List").WinMenu("menuobjtype:=3").Select "Import Medication"
IsPresent = Dialog("regexpwndtitle:=Medication Import").Exist
Call RecordResults(IsPresent, "Medication Import Dialog is Present")

'Try to search the searchbar with just 2 letters
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Type "Al"
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Search").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "This Test is passes",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Needs Attention!!!!"
End If
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Set ""

'Verifying the Meidcation Import Dialog
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Type "alv"
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Search").Click
wait 2
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Help Button
Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=&Help").Click
IsPresent = Window("NexTech Practice User").Exist
Call RecordResults(IsPresent, "NexTech Practice User Manual dialog is Present")
Window("NexTech Practice User").Close

'Clicking on Inactive Medications button
Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Inactive Medications").Click
IsPresent = Dialog("regexpwndtitle:=Inactive Medications").Exist
Call RecordResults(IsPresent, "Inactive Medications Dialog is Present")
Dialog("regexpwndtitle:=Inactive Medications").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Latin Notation button
Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Edit Latin Notation").Click
IsPresent = Dialog("regexpwndtitle:=Latin Prescription Settings").Exist
Call RecordResults(IsPresent, "Latin Prescription Settings Dialog is Present")
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Delete").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Cancel").Click
Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Close").Click
Dialog("regexpwndtitle:=Medication Schedule Detail").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Button in Medication Schedule Setup
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Edit Detail").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Remove Detail Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Remove Detail").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Load Form Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Load From").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Save To Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Save To").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Save New Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Save New").Click
If Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is Not Open"
End If

Dialog("regexpwndtitle:=Practice").WinEdit("regexpwndclass:=Edit").Type "Test"
Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Delete Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Delete","index:=0").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on the Preview Button
Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Preview").Click
IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Window("Nextech Main Window").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Cancel").Click

If Dialog("regexpwndtitle:=Practice").Exist Then
	Dialog("regexpwndtitle:=Practice").Winbutton("regexpwndtitle:=&Yes").Click
End If

End Sub
 
Sub VerifyGen1()

If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Nextech Practice").Winbutton("regexpwndtitle:=OK").Exist (1) Then
	Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Nextech Practice").Winbutton("regexpwndtitle:=OK").Click
End If

' Click the ellipsis button

	Window("Nextech Main Window").Window("Patients").WinButton("General 1: Ellipsis WinButton").Click
	IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Edit Prefixes").Check (CheckPoint("Edit Prefixes"))
	Call RecordResults(IsPresent, "Gen 1: Edit Prefixes dialog")	
	
'Clicking on Add and Remove Button
	Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Edit Prefixes").WinButton("regexpwndtitle:=Add").Click
	Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Edit Prefixes").WinButton("regexpwndtitle:=Remove").Click
	IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Edit Prefixes").Dialog("regexpwndtitle:=NexTech Practice").Exist
	Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
	Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Edit Prefixes").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

	Window("Nextech Main Window").Dialog("Edit Prefixes").WinButton("Cancel").Click
	
' Click the Send To button

	Window("Nextech Main Window").Window("Patients").WinButton("General 1: Send To button").Click
	wait(.1)
	WinMenu("ContextMenu").Select "Open Quickbooks Link..."
	wait(.1)
	IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Link To QuickBooks").Check (CheckPoint("Link To QuickBooks"))
	Call RecordResults(IsPresent, "Gen 1: Link To QuickBooks dialog opened")	

	Window("Nextech Main Window").Dialog("regexpwndtitle:=Link To QuickBooks").WinButton("regexpwndtitle:=Setup Patient Name / Acct\. Number Export").Click
	IsPresent = Dialog("regexpwndtitle:=Setup Patient Name / Account Number Export").Check (CheckPoint("Setup Patient Name / Account Number Export"))
	Call RecordResults(IsPresent, "Gen 1: Setup Patient Name / Account dialog opened")	

	Dialog("regexpwndtitle:=Setup Patient Name / Account Number Export").WinButton("regexpwndtitle:=Cancel").Click
	
'Clicking on Default Source /Deposit Account
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Link To QuickBooks").WinButton("regexpwndtitle:=Default Source / Deposit Accounts").Click
	If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Link To QuickBooks").Dialog("regexpwndtitle:=Practice").Exist (1) Then
		Reporter.ReportEvent micPass, "Practice Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
	End If

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Link To QuickBooks").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Link To QuickBooks").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Open")

Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Link To QuickBooks").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Link To QuickBooks").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Open")

Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Link To QuickBooks").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Window("Nextech Main Window").Dialog("Link To QuickBooks").WinButton("Close").Click

'Clicking on Send to and then WinMenu CareCredit
Window("Nextech Main Window").Window("Patients").WinButton("General 1: Send To button").Click
WinMenu("ContextMenu").Select "CareCredit..."
IsPresent = Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndclass:=#32770","index:=1").WinButton("regexpwndtitle:=Return to NexTech").Exist
Call RecordResults(IsPresent, "CareCredit redirect Dialog is Open")
Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndclass:=#32770","index:=1").WinButton("regexpwndtitle:=Return to NexTech").Click

'Clicking on Send to and then Send to QuickBooks
Window("Nextech Main Window").Window("Patients").WinButton("General 1: Send To button").Click
WinMenu("ContextMenu").Select "Send To Quickbooks"

IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Practice Dialog is Open")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click

IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Send to and then Open Mirror Link...
Window("Nextech Main Window").Window("Patients").WinButton("General 1: Send To button").Click
WinMenu("ContextMenu").Select "Open Mirror Link..."

IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Mirror data could not be opened").Exist
Call RecordResults(IsPresent, "Mirror data could not be opened Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Mirror data could not be opened").WinButton("regexpwndtitle:=&No").Click

'Clicking on Send to and then Open Inform Link...
Window("Nextech Main Window").Window("Patients").WinButton("General 1: Send To button").Click
WinMenu("ContextMenu").Select "Open Inform Link..."

IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
wait(1)
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

' <<<< Practice has stopped working here >>>>
wait(1)
IsPresent = Dialog("regexpwndtitle:=Open").Exist
Call RecordResults(IsPresent, "Open Dialog is Present")
Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Send to and then Open United Link...
Window("Nextech Main Window").Window("Patients").WinButton("General 1: Send To button").Click
WinMenu("ContextMenu").Select "Open United Link..."

IsPresent = Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Click

'Clicking on Groups Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Groups").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").WinButton("regexpwndtitle:=Add/Edit Groups").Exist
Call RecordResults(IsPresent, "Groups Dialog box is Present")

'Clicking on Add/Edit Groups
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").WinButton("regexpwndtitle:=Add/Edit Groups").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").Dialog("regexpwndtitle:=Groups","index:=2").Exist
Call RecordResults(IsPresent, "Groups Dialog is box is open the second time")

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").Dialog("regexpwndtitle:=Groups","index:=2").WinButton("regexpwndtitle:=New Group").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").Dialog("regexpwndtitle:=Groups","index:=2").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").Dialog("regexpwndtitle:=Groups","index:=2").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

'Closing Both the Groups Dialog box
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").Dialog("regexpwndtitle:=Groups","index:=2").WinButton("regexpwndtitle:=Cancel").Click

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Groups","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Reminder Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Reminders").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Reminders Sent for.*").Exist
Call RecordResults(IsPresent, "Reminder sent for Dialog box is Open")

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Reminders Sent for.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on Vision Prescriptions...
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Vision Prescriptions…").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:= Vision Prescriptions").Exist
Call RecordResults(IsPresent, "Vision Prescription Dialog is Open")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:= Vision Prescriptions").WinButton("regexpwndtitle:=Close").Click

'Clicking on Recalls... Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Recalls\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Exist
Call RecordResults(IsPresent, "Recall Needing Attention Dialog is Present")

'Clicking on Create Merge Group inside the Recalls Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").WinButton("regexpwndtitle:=Create Merge Group").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Practice Enter Group Name popup is present")

Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click

If Window("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=NexTech Practice").Static("regexpwndtitle:=That group name already exists, please choose.*").Exist (1) Then
	Window("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
	Call RecordResults(True, "Expected 'Group Already Exists' message appears ")
Else
	Call RecordResults(True, "No Already Exists popup")
End If

'Clicking on Merge to word
Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").WinButton("regexpwndtitle:=Merge To Word").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Practice Dialog is Present")
'We need to allow for the condition that the recalls needing attention list is empty, so...
If Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=Practice").Exist (1) Then
	Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click
Else
	Window("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click
End If

'Clicking on Create New Recall
Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").WinButton("regexpwndtitle:=Create New Recall").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:= Create Patient Recall for.*").Exist (1)
Call RecordResults(IsPresent, "Create Patient Recall Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").Dialog("regexpwndtitle:= Create Patient Recall for.*").WinButton("regexpwndtitle:=Cancel").Click

'Closing the Recall Needing Attention Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("text:=Recalls Needing Attention.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on Security... Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Security\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").Exist
Call RecordResults(IsPresent, "Security Groups Dialog is Present")

'Clicking on Add/Edit Groups in Security Group Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").WinButton("regexpwndtitle:=Add/Edit Groups\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").Dialog("regexpwndtitle:=Configure Security Groups").Exist
Call RecordResults(IsPresent, "Configure Security Groups is Present")

'Clicking on Add Group Button
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").Dialog("regexpwndtitle:=Configure Security Groups").WinButton("regexpwndtitle:=Add Group").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").Dialog("regexpwndtitle:=Configure Security Groups").Dialog("regexpwndtitle:=Practice").Exist
Call RecordResults(IsPresent, "Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").Dialog("regexpwndtitle:=Configure Security Groups").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

'Closing the dialog boxes
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").Dialog("regexpwndtitle:=Configure Security Groups").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Security Groups for.*").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Patient Summary... Button
Window("Nextech Main Window").Window("Patients").WinButton("regexpwndtitle:=Patient Summary\.\.\.").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Patient Summary for.*").Exist
Call RecordResults(IsPresent, "Patient Summary Dialog is Present")

'Close the Patient Summary description popup if present
If Window("regexpwndtitle:=Patient Summary for.*").Dialog("regexpwndtitle:=Patient Summary").WinCheckBox("regexpwndtitle:=Don't Show Me Again").Exist (1) Then
		Reporter.ReportEvent micPass, "Test Passed", "Patient Summary bottom screen description popup opened"
		Window("regexpwndtitle:=Patient Summary for.*").Dialog("regexpwndtitle:=Patient Summary").WinButton("regexpwndtitle:=OK").Click
End If

'Clicking on Configure Button inside Patient Summary Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Patient Summary for.*").WinButton("regexpwndtitle:=Configure").Click
IsPresent = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Patient Summary for.*").Dialog("regexpwndtitle:=Configure Patient Summary").Exist
Call RecordResults(IsPresent, "Configure Patient Summary Dialog is Present")
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Patient Summary for.*").Dialog("regexpwndtitle:=Configure Patient Summary").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Help Button in Patient Summary Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Patient Summary for.*").WinButton("regexpwndtitle:=Help").Click
IsPresent = Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Patient Summary for.*").Dialog("regexpwndtitle:=NexTech Practice").Exist
Call RecordResults(IsPresent, "Nextech Practice Dialog is Present")
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Patient Summary for.*").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Patient Summary for.*").WinButton("regexpwndtitle:=Close").Click
IsPresent = Window("Nextech Main Window").Window("Patients").WinRadioButton("regexpwndtitle:=Patient Prospect").Exist
Call RecordResults(IsPresent, "Patients General 1 Tab is Present")

End Sub

Sub VerifyTracking()
	' Click the Add button

	Window("Nextech Main Window").Window("Patients").WinButton("Add").Click
	IsPresent = Window("Nextech Main Window").Dialog("Add Procedure").Check (CheckPoint("Add Procedure"))
	Call RecordResults(IsPresent, "Tracking: Add button click dialog opened")	
	Window("Nextech Main Window").Dialog("Add Procedure").WinButton("Cancel").Click

'TODO: Add the code to test the procedure information button click.  It requires that the list be populated, and a procedure selected.
	
End Sub

Sub VerifyFollowUp()
	
' Click the Create To-Do Task button

	Window("Nextech Main Window").Window("Patients").WinButton("Follow Up: Create To-Do Task Button").Click
	IsPresent = Window("Nextech Main Window").Dialog("Modify Task").Check (CheckPoint("Modify Task"))
	Call RecordResults(IsPresent, "Follow Up: create todo task button click dialog opened")	
	Window("Nextech Main Window").Dialog("Modify Task").WinButton("Cancel").Click

' Click the View To-Do List button

	Window("Nextech Main Window").Window("Patients").WinButton("View To-Do List").Click
	IsPresent = Dialog("To Do Alarm").Check (CheckPoint("To Do Alarm"))
	Call RecordResults(IsPresent, "Follow Up: view todo list button click dialog opened")	

	Dialog("To Do Alarm").WinButton("Close").Click
' Click the Delete Completed button

	Window("Nextech Main Window").Window("Patients").WinButton("Delete Completed").Click
	If Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndtitle:=NexTech").Static("regexpwndtitle:=Are you sure you wish to delete the 1 completed task\(s\)\?").Exist (1) Then
		Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndtitle:=NexTech").WinButton("regexpwndtitle:=&Yes").Click
		Call RecordResults(True, "Todo(s) were deleted")
	ElseIf Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndtitle:=NexTech Practice").Static("regexpwndtitle:=There are no completed follow.*").Exist (1) Then
		Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
		Call RecordResults(True, "No completed follow-ups")
	End If
	
'	IsPresent = Window("Nextech Main Window").Dialog("NexTech Practice PUP1").Check (CheckPoint("NexTech Practice"))  probably no longer needed because of the if statements above
	'Call RecordResults(IsPresent, "Follow Up: Delete Completed button click dialog opened")	
	'Window("Nextech Main Window").Dialog("NexTech Practice PUP1").WinButton("OK").Click

' Click the Edit Categories button

	Window("Nextech Main Window").Window("Patients").WinButton("Edit Categories").Click
	IsPresent = Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Check (CheckPoint("Note / Follow-Up Categories"))
	Call RecordResults(IsPresent, "Follow Up: view todo list button click dialog opened")	

' Within Note/Follow-up, Click New button


	Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").WinButton("text:=&New").Click
'	IsPresent = Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Dialog("text:=Input").Check (CheckPoint("Input_3"))

	IsPresent = Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Input").Static("Enter Category Name").Check (CheckPoint("Enter Category Name_2"))



	'IsPresent = Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Exist
	Call RecordResults(IsPresent, "Follow Up: Input new category name dialog opened.")	
	
	Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Input").WinButton("Cancel").Click
	
	
	'Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Dialog("regexpwndclass:=#32770","text:=Input").WinButton("&Cancel").Click

' Within Note/Follow-up, Click Adv. EMR Merge button


	Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").WinButton("text:=&Adv\. EMR Merge\.\.\.").Click
	
	
IsPresent = Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").Check (CheckPoint("Advanced Merge - Default Categories_3"))
	Call RecordResults(IsPresent, "Advanced Merge Default Categories dialog opened")	
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Combine Categories").Click


IsPresent = Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").Check (CheckPoint("Combine Note Categories_3"))
	Call RecordResults(IsPresent, "Combine Note Categories opened")	
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").WinButton("Close").Click

Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Close").Click
 
	
	
	
	
'	IsPresent = Window("Nextech Main Window").Dialog("Advanced Merge - Default").Check (CheckPoint("Advanced Merge - Default Categories"))
'	Call RecordResults(IsPresent, "Follow Up: view todo list button click dialog opened")	
'	Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Dialog("Advanced Merge - Default").WinButton("Cancel").Click
'
'' Within Note/Follow-up, Click Combine Categories button
'
'	Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").WinButton("Combine Categories").Click
'	IsPresent = Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Dialog("NexTech Practice").Check (CheckPoint("NexTech Practice_2"))
'	Call RecordResults(IsPresent, "Follow Up: Combine Categories button click dialog opened")	
'	Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Dialog("NexTech Practice").WinButton("OK").Click
'
'	IsPresent = Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Dialog("Combine Note Categories").Check (CheckPoint("Combine Note Categories"))
'	Call RecordResults(IsPresent, "Follow Up: Combine Notes Categories button click dialog opened")	
'	Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").Dialog("Combine Note Categories").WinButton("Close").Click
'	Window("Nextech Main Window").Dialog("text:=Note / Follow-Up Categories").WinButton("text:=Clos&e").Click
'
End Sub

Sub VerifyLabs()
'msgbox ("In VerifyLabs")	

Window("Nextech Main Window").Window("Patients").WinButton("Labs: New Lab button").Click
WinMenu("ContextMenu").Select "Biopsy (Biopsy)"
'Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinCalendar("Date:").SetDate "18-Nov-2015"
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Check (CheckPoint("VAC32, Patient32  (43) - Lab Entry"))
	Call RecordResults(IsPresent, "Labs: Patient Lab Entry dialog opened")	

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("Add New").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").Static("This lab cannot be saved").Exist
	Call RecordResults(IsPresent, "Labs: Labs popup opened.")	

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").WinButton("OK").Click

'if this step fails with a practice db error, it is most likely due to the global practice preference "Default Lab Notes Category" not having a category selected, 
'meaning the manual db setup hasn't been done.  Just select a category for it and re-run the tests.  Same for "Default categoy for Lab to-do tasks"

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("Create To-Do").Click

IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Modify Task").Check (CheckPoint("Modify Task_2"))
	Call RecordResults(IsPresent, "Labs: Modify Task dialog opened")	

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Modify Task").WinButton("Cancel").Click

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("WinButton_6").Click

IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Lab Form Number Editor").Check (CheckPoint("Lab Form Number Editor"))
	Call RecordResults(IsPresent, "Labs: Lab Form Number Editor")	

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Lab Form Number Editor").WinButton("Cancel").Click

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("Create Recall").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Practice - Create Patient").Check (CheckPoint("Practice - Create Patient Recall for VAC32, Patient32 (43)_2"))
	Call RecordResults(IsPresent, "Labs: Practice - Create Patient Recall")	

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Practice - Create Patient").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("Add").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Multiple Selection").Check (CheckPoint("Multiple Selection_2"))
	Call RecordResults(IsPresent, "Labs: Multiple Selection")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Multiple Selection").WinButton("Cancel").Click

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("Preview").Click
IsPresent = IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").Static("This lab cannot be saved").Exist
	Call RecordResults(IsPresent, "Labs popup opened")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").WinButton("OK").Click


Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("WinButton_5").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Lab Diagnoses").Check (CheckPoint("Edit Lab Diagnoses"))
	Call RecordResults(IsPresent, "Labs: Edit Lab Diagnoses")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Lab Diagnoses").WinButton("Close").Click
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("WinButton").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Biopsy Type List").Check (CheckPoint("Edit Biopsy Type List"))
	Call RecordResults(IsPresent, "Labs: Edit Biopsy Type List")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Biopsy Type List").WinButton("Close").Click

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("WinButton_2").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Anatomic Location").Check (CheckPoint("Edit Anatomic Location List"))
	Call RecordResults(IsPresent, "Labs: Edit Anatomic Location List")	

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Anatomic Location").WinButton("Close").Click
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("WinButton_3").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Anatomic Location_2").Check (CheckPoint("Edit Anatomic Location Qualifier List"))
	Call RecordResults(IsPresent, "Labs: LockManager")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Edit Anatomic Location_2").WinButton("Close").Click

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("Button").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").Static("This lab cannot be saved").Exist
	Call RecordResults(IsPresent, "Labs: Edit Anatomic Location Qualifier List")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").WinButton("OK").Click

Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Close
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").Static("This lab cannot be saved").Check (CheckPoint("Do you want to save your changes to this Lab?    'Yes' will save any changes and close the Lab.  'No' will discard any changes and close the Lab.  'Cancel' will cancel this action and leave the Lab open._2"))
	Call RecordResults(IsPresent, "Labs: popup opened")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").WinButton("No").Click
Window("Nextech Main Window").Window("Patients").WinButton("Labs: New Lab button").Click
WinMenu("ContextMenu").Select "Diagnostic (Diagnostics)"
'Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinCalendar("Date:").SetDate "19-Nov-2015"
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").WinButton("WinButton_7").Click
IsPresent = Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Labs To Be Ordered Setup").Check (CheckPoint("Labs To Be Ordered Setup"))
	Call RecordResults(IsPresent, "Labs: Labs To Be Ordered Setup")	
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Labs To Be Ordered Setup").WinButton("Edit Labs To Be Ordered").Click
IsPresent = Dialog("Edit Labs To Be Ordered").Check (CheckPoint("Edit Labs To Be Ordered List"))
	Call RecordResults(IsPresent, "Labs: Edit Labs To Be Ordered")	
Dialog("Edit Labs To Be Ordered").WinButton("Close").Click
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("Labs To Be Ordered Setup").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Close
Window("Nextech Main Window").Dialog("VAC32, Patient32  (43)").Dialog("NexTech Practice").WinButton("No").Click
Window("Nextech Main Window").Window("Patients").WinButton("Labs Needing Attention").Click
IsPresent = Dialog("Labs Needing Attention").Check (CheckPoint("Labs Needing Attention"))
	Call RecordResults(IsPresent, "Labs: Labs Needing Attention")	
Dialog("Labs Needing Attention").WinButton("Close").Click
Window("Nextech Main Window").Window("Patients").WinButton("Preview Results Graph").Click
IsPresent = Window("Nextech Main Window").Dialog("NexTech Practice PUP1").Static("Practice was unable to").Check (CheckPoint("There are no lab results for this patient."))
	Call RecordResults(IsPresent, "Labs: popup opened")	
Window("Nextech Main Window").Dialog("NexTech Practice PUP1").WinButton("OK").Click

End Sub

Sub VerifyGen2()

' Click the rewards point edit button
	Window("Nextech Main Window").Window("regexpwndtitle:=Patients","index:=0").WinButton("regexpwndtitle:=Edit").Click
	'Window("regexpwndtitle:=Nextech \(UFT*.").Dialog("regexpwndtitle:=Patients").WinButton("regexpwndtitle:=Edit").Highlight
	IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Input").Exist
	'TestResult = Call RecordResults(IsPresent, "Gen 2: Edit button pressed, Input dialog present")
	Call RecordResults(IsPresent, "Gen 2: Edit button pressed, Input dialog present")
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Input").WinButton("text:= &Cancel").Click
	
' Click the copy button
	Window("Nextech Main Window").Window("regexpwndtitle:=Patients","index:=0").WinButton("regexpwndtitle:=Copy...").Click
	IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Employer List").Exist
	Call RecordResults(IsPresent, "Gen 2: Copy button pressed, Employer List dialog present")
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Employer List").WinButton("regexpwndtitle:=Cancel").Click
	
' Click the add referral button	
	Window("Nextech Main Window").Window("regexpwndtitle:=Patients","index:=0").WinButton("regexpwndtitle:=Add").Click
	IsPresent = Window("regexpwndtitle:=Nextech \(UFT.....DB\.*\)").Dialog("regexpwndtitle:=Referral Source","object class:=#32770").Exist
'	IsPresent = Window("Nextech Main Window").Dialog("regexpwndtitle:=Referral Source").Exist
	Call RecordResults(IsPresent, "Gen 2: Add button pressed, Referral Source dialog present")
	
' Click the add new top level referral button

	Window("regexpwndtitle:=Nextech \(UFT.....DB\.*\)").Dialog("regexpwndtitle:=Referral Source","object class:=#32770").WinButton("regexpwndtitle:=Add New Top Level Referral").Click @@ hightlight id_;_657056_;_script infofile_;_ZIP::ssf1.xml_;_
	'Window("regexpwndtitle:=Nextech (UFT*.)").Dialog("regexpwndtitle:=Referral Source").WinButton("regexpwndtitle:=Add New Top Level Referral").Click
	IsPresent = Window("regexpwndtitle:=Nextech \(UFT.....DB\.*\)").Dialog("regexpwndtitle:=Referral Source","object class:=#32770").Dialog("regexpwndtitle:=Input").Exist
	Call RecordResults(IsPresent, "Gen 2: Add New Top Level Referral button pressed, Input dialog present")
	Window("regexpwndtitle:=Nextech \(UFT.....DB\.*\)").Dialog("regexpwndtitle:=Referral Source","object class:=#32770").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click	
	
' Click the help button on referral source dialog
	Window("regexpwndtitle:=Nextech \(UFT.....DB\.*\)").Dialog("regexpwndtitle:=Referral Source","object class:=#32770").WinButton("regexpwndtitle:=&Help").Click
	IsPresent = Window("NexTech Practice User").Check (CheckPoint("NexTech Practice User Manual"))
	Call RecordResults(IsPresent, "Gen 2: Help button pressed, Practice User Manual dialog present")
	Window("NexTech Practice User").Close	
	Window("Nextech Main Window").Dialog("Referral Source").WinButton("Cancel").Click

' Click the patient type ellipse button
	Window("Nextech Main Window").Window("Patients").WinButton("PatientTypeEllipse").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Patient Type").Check (CheckPoint("Edit Patient Type"))
	Call RecordResults(IsPresent, "Gen 2: Edit Patient Type pressed, Edit Payment Type dialog present")
	Window("Nextech Main Window").Dialog("Edit Patient Type").WinButton("Close").Click

' Click Race ellipse

	Window("Nextech Main Window").Window("Patients").WinButton("RaceEllipse").Click
	IsPresent = Window("Nextech Main Window").Dialog("Edit Race List").Check (CheckPoint("Edit Race List"))
	Call RecordResults(IsPresent, "Gen 2: Edit Race List dialog present")
	
' Click the Add button on the race dialog
	Window("Nextech Main Window").Dialog("Edit Race List").WinButton("Add").Click
	IsPresent = Dialog("Input").Check (CheckPoint("Edit from Race List dialog"))
	Call RecordResults(IsPresent, "Gen 2: Input dialog present")
	Dialog("Input").WinButton("Cancel").Click
	Window("Nextech Main Window").Dialog("Edit Race List").WinButton("Close").Click
	
' Click the Referring Physician ellipse button

	Window("Nextech Main Window").Window("Patients").WinButton("RefPhysEllipse").Click
	IsPresent = Window("Nextech Main Window").Dialog("Create New Contact").Check (CheckPoint("Create New Contact dialog"))
	Call RecordResults(IsPresent, "Gen 2: Create New Contact dialog present")
	Window("Nextech Main Window").Dialog("Create New Contact").WinButton("Cancel").Click
	
' Click the PCP ellipse button

	Window("Nextech Main Window").Window("Patients").WinButton("WinButton").Click
	IsPresent = Window("Nextech Main Window").Dialog("Create New Contact").Check (CheckPoint("Create New Contact"))
	Call RecordResults(IsPresent, "Gen 2: Create New Contact dialog present")
	Window("Nextech Main Window").Dialog("Create New Contact").WinButton("Cancel").Click

'' Click the New Allocation button
'
'	Window("Nextech Main Window").Window("Patients").WinButton("New Allocation").Click
'	IsPresent = Window("Nextech Main Window").Dialog("Patient Inventory Allocation").Check (CheckPoint("Patient Inventory Allocation"))
'	Call RecordResults(IsPresent, "Gen 2: Pat Inv Alloc dialog present")
'
'' Click the Preview button
'
'	Window("Nextech Main Window").Dialog("Patient Inventory Allocation").WinButton("Preview").Click
'	IsPresent = Dialog("Practice").Check (CheckPoint("Practice"))
'	Call RecordResults(IsPresent, "Gen 2: Practice popup dialog present")
'	Dialog("Practice").WinButton("No").Click
'	Window("Nextech Main Window").Dialog("Patient Inventory Allocation").WinButton("Cancel").Click
'
'' Click the Warrenty Info button

	Window("Nextech Main Window").Window("Patients").WinButton("Warranty Info.").Click
	IsPresent = Window("Nextech Main Window").Dialog("Warranty Information").Check (CheckPoint("Warranty Information"))
	Call RecordResults(IsPresent, "Gen 2: Warrenty Info dialog present")
	Window("Nextech Main Window").Dialog("Warranty Information").WinButton("Close").Click

' Click the Edit NexWeb Login button


	Window("Nextech Main Window").Window("Patients").WinButton("Edit NexWeb Login").Click
	IsPresent = Window("Nextech Main Window").Dialog("NexWeb Login Information").Check (CheckPoint("NexWeb Login Information"))
	Call RecordResults(IsPresent, "Gen 2: NexWeb Login dialog is present")
	Window("Nextech Main Window").Dialog("NexWeb Login Information").WinButton("Close").Click


End Sub




