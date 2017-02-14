

Dim IsPresent
Dim TestDBName : TestDBName = Environment.Value("DBName")
RunAction "LogIn [GlobalPracticeActions]", oneIteration, TestDBName
'Log in to Practice
'
'
'On Error Resume Next


'Navigate to the Scheduler module

wait(4)
Window("Nextech Main Window").WinMenu("Menu").Select "Modules;Scheduler"

'Ensure the scheduler tab is present
'IsPresent = Window("Nextech Main Window").Window("Scheduler").Static("Today").Check (CheckPoint("Today"))
'Call RecordResults(IsPresent, "Scheduler: scheduler dialog opened")	

'Click the event control, pin the dialog and verify that the Event dialog is present
Window("Nextech Main Window").Window("Scheduler").Static("regexpwndtitle:=Event").Click
Window("Nextech Main Window").Dialog("Event").WinCheckBox("Button").Set "ON"
Window("Nextech Main Window").Dialog("Event").Check CheckPoint("Event")

'Click the Create Recall button and verify the Create Patient Recall dialog is open
Window("Nextech Main Window").Dialog("Event").WinButton("Create Recall").Click

'Need to allow for patient not being selected in the Patient dropdown
If Dialog("regexpwndclass:=#32770","regexpwndtitle:=Practice").Static("regexpwndtitle:=Unable to determine patient to create recall for\.  Please select a patient and try again\.").Exist (1) Then
	Dialog("regexpwndclass:=#32770","regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click
	Reporter.ReportEvent micPass, "Popup for no patient selected",""
	 
End If


Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Dialog("regexpwndclass:=#32770","regexpwndtitle:=Event ").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=2").highlight


Dialog("Practice - Create Patient").Check CheckPoint("Practice - Create Patient Recall for ,   (-25)")
Dialog("Practice - Create Patient").WinButton("Cancel").Click  '*** Broke on the line above

'Test that the More Insurance button click works for Add New Insured Party
Window("Nextech Main Window").Dialog("Event").WinButton("More Insurance").Click
Window("Nextech Main Window").Dialog("Event").WinMenu("ContextMenu").Select "Add New Insured Party"
Dialog("Insured Party Info").Check CheckPoint("Insured Party Info")
Dialog("Insured Party Info").WinButton("Cancel").Click

'Test that the More Insurance button click works for Set Additional Insured Parties


'Window("Nextech Main Window").Window("Scheduler").WinButton("Event").Click
Window("Nextech Main Window").Dialog("Event").WinButton("More Insurance").Click
Window("Nextech Main Window").Dialog("Event").WinMenu("ContextMenu").Select "Set Additional Insured Parties"
Dialog("Additional Insured Parties").Check CheckPoint("Additional Insured Parties")
Dialog("Additional Insured Parties").WinButton("Cancel").Click

'Unpin the Event dialog and exit
wait(2)
Window("Nextech Main Window").Dialog("Event").WinCheckBox("Button").Set "OFF"
Window("Nextech Main Window").Dialog("Event").WinButton("Exit without Saving").Click

'Check the Resource Editor dialog
Window("Nextech Main Window").Window("Scheduler").WinButton("regexpwndtitle:=\.\.\.").Click
Window("Nextech (UFT12200DB)").WinMenu("menuobjtype:=3").Select "Edit Resources..."

'Window("Nextech Main Window").WinMenu("ContextMenu").Select "Edit Resources..."
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Resource Editor").Exist Then
	Reporter.ReportEvent micPass, "Resource Editor Dialog Box is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Resource Editor Dialog Box is Not Open"
End If

'Test the View Add Input dialog
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("ViewEllipseButton").Click
Window("Nextech Main Window").Dialog("Resource Editor").WinMenu("ContextMenu").Select "Add"
Dialog("Input").Check CheckPoint("Input")
Dialog("Input").WinButton("Cancel").Click

'Test Copy to specific users
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("Copy to Users").Click
Window("Nextech Main Window").Dialog("Resource Editor").WinMenu("ContextMenu").Select "Copy to Specific Users..."
'Dialog("NexTech Practice").Static("Please select a service").Check CheckPoint("CurrentStandardViewPopup")
IsPresent = VerifyPopups("NexTech Practice", "The current standard view settings will be saved first\. Would you like to continue\?")
Call RecordResults(IsPresent, "Resource Editor Continue popup")

Dialog("NexTech Practice").WinButton("Yes").Click

'Test that the Multiple Selection dialog opens
Dialog("Multiple Selection").Check CheckPoint("Multiple Selection From Resource Editor")
Dialog("Multiple Selection").WinButton("OK").Click

'Negative test when none are selected in the multi selection list 
Dialog("Multiple Selection").Dialog("NexTech Practice").Static("You must select at least").Check CheckPoint("NegativeTestForMultiSelectDialog")
Dialog("Multiple Selection").Dialog("NexTech Practice").WinButton("OK").Click
Dialog("Multiple Selection").WinButton("Cancel").Click

'Test selecting Copy to All Users for the Copy To Users button click
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("Copy to Users").Click
Window("Nextech Main Window").Dialog("Resource Editor").WinMenu("ContextMenu").Select "Copy to All Users"

'Dialog("NexTech Practice").Static("Please select a service").Check CheckPoint("CurrentStandardViewPopup")
IsPresent = VerifyPopups("NexTech Practice", "The current standard view settings will be saved first\. Would you like to continue\?")
Call RecordResults(IsPresent, "Resource Editor Continue popup 2")

Dialog("NexTech Practice").WinButton("No").Click


'Test that the New button opens the Input diaog
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("New").Click
Dialog("Input").Static("Enter a list of required").Check CheckPoint("Enter a new resource_2")
Dialog("Input").WinButton("Cancel").Click

'Test the Delete button
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("Delete").Click
'Dialog("NexTech Practice").Static("Please select a service").Check CheckPoint("AreYouSurePopup")
If Dialog("NexTech Practice").Exist Then
	Call RecordResults(IsPresent, "Appointment schedule pop up apprears")
End If

If Window("Nextech Main Window").Dialog("Resource Editor").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Window("Nextech Main Window").Dialog("Resource Editor").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Else
	Dialog("NexTech Practice").WinButton("No").Click
End If


'Test that the Allowed Purposes...button opens the Schedule Appointes Configuration dialog
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("Allowed Purposes").Click
Dialog("Schedulable Appointments").Check CheckPoint("Schedulable Appointments Configuration")
Dialog("Schedulable Appointments").WinButton("Cancel").Click

'Test that the Merge Resources...button opens the Combine Schedule Resources dialog
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("Merge Resources").Click
Dialog("Combine Scheduler Resources").Check CheckPoint("Combine Scheduler Resources")

'Test the Combine Selected Resources button (PL67820)
Dialog("Combine Scheduler Resources").WinButton("Combine Selected Resources").Click
Dialog("Combine Scheduler Resources").WinButton("Close").Click

'Cancel out of the Resource Editor dialog
Window("Nextech Main Window").Dialog("Resource Editor").WinButton("Cancel").Click

'Test that the Room Manager dialog opens
Window("Nextech Main Window").Window("Scheduler").WinButton("RoomMgrButton").Click
Window("Nextech Main Window").Dialog("Room Manager").Check CheckPoint("Room Manager")

'Test that the Edit Rooms dialog opens and verify the button statuses
Window("Nextech Main Window").Dialog("Room Manager").WinButton("Edit Rooms").Click
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Setup").Check CheckPoint("Room Setup")
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Setup").WinButton("Add Room").Check CheckPoint("Add Room")
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Setup").WinButton("Delete Room").Check CheckPoint("Delete Room")
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Setup").WinButton("Inactivate Room").Check CheckPoint("Inactivate Room")
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Setup").WinButton("Close").Check CheckPoint("Close")


'Test that the input dialog opens for add and then cancel out of add and close Room Setup
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Setup").WinButton("Add Room").Click
Dialog("Input").Static("Enter a list of required").Check CheckPoint("Enter a new room:_2")
Dialog("Input").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Setup").WinButton("Close").Click

'Test that Room Status Setup dialog opens
Window("Nextech Main Window").Dialog("Room Manager").WinButton("Edit Room Statuses").Click
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Status Setup").Check CheckPoint("Room Status Setup")

'Test that the Input dialog opens for Add Status then cancel out and close Room Status Setup

Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Status Setup").WinButton("Add Status").Click
Dialog("Input").Static("Enter a list of required").Check CheckPoint("Enter a new status:")
Dialog("Input").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Room Status Setup").WinButton("Close").Click

'Test that the Configure Columns dialog opens
Window("Nextech Main Window").Dialog("Room Manager").WinButton("Configure Columns").Click
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Configure Columns").Check CheckPoint("Configure Columns")
Window("Nextech Main Window").Dialog("Room Manager").Dialog("Configure Columns").WinButton("Close").Click

'Test the Preview Button
Window("Nextech Main Window").Dialog("Room Manager").WinButton("Preview").Click
'Window("Nextech Main Window").Window("Room Manager (PP)").Static("Static").Check CheckPoint("0 of 0 from RM")
wait(3)
Window("Nextech Main Window").Window("Room Manager (PP)_2").Close
Window("Nextech Main Window").Window("Scheduler").WinButton("RoomMgrButton").Click
Window("Nextech Main Window").Dialog("Room Manager_2").WinButton("Close").Click

'Test that the Recalls Needing Attention dialog opens
Window("Nextech Main Window").Window("Scheduler").WinButton("RNAButton").Click
Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").Check CheckPoint("Recalls Needing Attention_2")

'Test Create Merge Group
Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").WinButton("Create Merge Group").Click

If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Recalls Needing Attention").Dialog("regexpwndtitle:=Practice").Exist Then
	Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Recalls Needing Attention").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click
Else
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Recalls Needing Attention").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click
End If


'Test Merge To Word
Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").WinButton("Merge To Word").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Recalls Needing Attention").Dialog("regexpwndtitle:=Practice").Exist Then
	Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Recalls Needing Attention").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click
Else
	Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").Dialog("Practice").WinButton("No").Click
End If


'Test Create New Recall dialog, cancel out and close Recalls Needing Attention
Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").WinButton("Create New Recall").Click
Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").Dialog("Practice - Create Patient").Check CheckPoint("Practice - Create Patient Recall for VAC32, Patient32 (43)")
Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").Dialog("Practice - Create Patient").WinButton("Cancel").Click
Window("Nextech Main Window").Dialog("Recalls Needing Attention from Schedule button click").WinButton("Close").Click




'*****************************************************************************************
'**** Behavior of the RescheduleAppointments queue is seeminly unpredictable.  It appears
'**** that today, the default is not popped out, so the code here assumes that behavior

Window("Nextech Main Window").Window("Scheduler").WinButton("RQButton").Click
Window("Nextech Main Window").Window("Scheduler").WinButton("Reschedule Appointments").Check CheckPoint("Reschedule Appointments")

'Test that the Reschedule Appointments dialog opens
Window("Nextech Main Window").Window("Scheduler").WinButton("Reschedule Appointments").Click
Window("Nextech Main Window").Dialog("Reschedule Appointments").Check CheckPoint("ReschedulingAptsdialog_opened")
Window("Nextech Main Window").Dialog("Reschedule Appointments").WinButton("Close").Click

'Test the Reschedule Appointments dialog
Window("Nextech Main Window").Window("Scheduler").WinButton("Reschedule Appointments").Click
Window("Nextech Main Window").Dialog("Reschedule Appointments").Check CheckPoint("RSAptsDialog_OpenedAgain")

'Negative test when clicking the Reschedule button with no times selected
Window("Nextech Main Window").Dialog("Reschedule Appointments").WinButton("Reschedule").Click
Dialog("NexTech Practice").Check CheckPoint("NoTimeFramePopup_2")
Dialog("NexTech Practice").WinButton("OK").Click

'Test the Cancellation ellipse button
Window("Nextech Main Window").Dialog("Reschedule Appointments").WinButton("CancellationEllipseButton").Click
Dialog("Edit Cancellation Reasons").Check CheckPoint("Edit Cancellation Reasons")
Dialog("Edit Cancellation Reasons").WinButton("Close").Click

'Test the buttons on Edit Cancellation Reasons and then close.
Window("Nextech Main Window").Dialog("Reschedule Appointments").WinButton("Reschedule").Click

'Dialog("NexTech Practice").Static("Please select a service").Check CheckPoint("You have not selected a time frame from docked access")
IsPresent = VerifyPopups("NexTech Practice", "You have not selected a time frame\.")
Call RecordResults(IsPresent, "voce have not selected a time frame from docked access")
Dialog("NexTech Practice").WinButton("OK").Click

Window("Nextech Main Window").Dialog("Reschedule Appointments").WinButton("CancellationEllipseButton").Click
Dialog("Edit Cancellation Reasons").WinButton("Add").Check CheckPoint("CancellationReasonsAdd")
Dialog("Edit Cancellation Reasons").WinButton("Edit").Check CheckPoint("Edit_2")
Dialog("Edit Cancellation Reasons").WinButton("Delete").Check CheckPoint("Delete")
Dialog("Edit Cancellation Reasons").WinButton("Close").Click
Window("Nextech Main Window").Dialog("Reschedule Appointments").WinButton("Close").Click

'******************************************************************************************
'Window("Nextech Main Window").Window("Scheduler").WinButton("RQButton").Click
'
'Window("Nextech Main Window").Dialog("Rescheduling Queue").WinButton("Reschedule Appointments").Click
'Window("Nextech Main Window").Dialog("Rescheduling Queue").Dialog("Reschedule Appointments").WinButton("Cancellation Reason Ellipse").Click
'Dialog("Edit Cancellation Reasons").Check CheckPoint("Edit Cancellation Reasons From Popout Access")
'Dialog("Edit Cancellation Reasons").WinButton("Close").Click
'Window("Nextech Main Window").Dialog("Rescheduling Queue").Dialog("Reschedule Appointments").WinButton("Reschedule").Click
'
'Dialog("NexTech Practice").Static("Please select a service").Check CheckPoint("You have not selected a time frame from RS popout")
'Dialog("NexTech Practice").WinButton("OK").Click
'Window("Nextech Main Window").Dialog("Rescheduling Queue").Dialog("Reschedule Appointments").WinButton("Close").Click
'



'Test that the RecallScheduling Queue dialog opens and dock it to the Scheduler
'Window("Nextech Main Window").Dialog("Rescheduling Queue").WinButton("Dock to Scheduler").Click
'Window("Nextech Main Window").Dialog("Reschedule Appointments").WinButton("Close").Click
'Window("Nextech Main Window").Window("Scheduler").WinButton("Pop-out").Click
'Window("Nextech Main Window").Dialog("Rescheduling Queue").WinButton("Close").Click


Window("Nextech Main Window").Window("Scheduler").WinButton("Pop-out").Check CheckPoint("Pop-out")

'Click the pop-out button 


Window("Nextech Main Window").Window("Scheduler").WinButton("Pop-out").Click
wait(.5)
Window("Nextech Main Window").Dialog("Rescheduling Queue").Check CheckPoint("Rescheduling Queue Dialog Is Present")
Window("Nextech Main Window").Dialog("Rescheduling Queue").WinButton("Reschedule Appointments").Click
wait(.5)
Window("Nextech Main Window").Dialog("Rescheduling Queue").Dialog("Reschedule Appointments").Check CheckPoint("Reschedule Apts from rs queue dialog is present")
Window("Nextech Main Window").Dialog("Rescheduling Queue").Dialog("Reschedule Appointments").WinButton("Cancellation Reason Ellipse").Click
wait(.5)
Dialog("Edit Cancellation Reasons").Check CheckPoint("Edit Cancellation Reasons dialog is present")
Dialog("Edit Cancellation Reasons").WinButton("Close").Click
wait(.5)
Window("Nextech Main Window").Dialog("Rescheduling Queue").Dialog("Reschedule Appointments").WinButton("Reschedule").Click

'Dialog("NexTech Practice").Static("Please select a service").Check CheckPoint("You have not selected a time frame popup from popout")
IsPresent = VerifyPopups("NexTech Practice", "You have not selected a time frame.")
Call RecordResults(IsPresent, "You have not selected a time frame popup from popout popup")

Dialog("NexTech Practice").WinButton("OK").Click
Window("Nextech Main Window").Dialog("Rescheduling Queue").Dialog("Reschedule Appointments").WinButton("Close").Click
Window("Nextech Main Window").Dialog("Rescheduling Queue").WinButton("Dock to Scheduler").Click

'log out of practice


Window("Nextech Main Window").WinToolbar("ModuleButtons").Press 1
'Window("Nextech Main Window").Check CheckPoint("PatientsModuleOpened")


wait(2)


RunAction "LogOut [GlobalPracticeActions]", oneIteration




