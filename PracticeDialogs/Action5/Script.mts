Dim TestDBName : TestDBName = "UFT12200DB"
Dim IsPresent
RunAction "LogIn [GlobalPracticeActions]", oneIteration, TestDBName

' Verify if Contacts Module opens

Window("Nextech Main Window").WinMenu("Menu").Select "Modules;Contacts"

wait(2)

Call ChangeTabs (Window("Nextech Main Window").Window("Contacts").ActiveX("NxTab Control"), "General")
Call VerTabOpened ("General")
Call VerifyGeneral()

Call ChangeTabs (Window("Nextech Main Window").Window("Contacts").ActiveX("NxTab Control"), "Notes")
Call VerTabOpened("Notes")
Call VerifyNotes()


Call ChangeTabs (Window("Nextech Main Window").Window("Contacts").ActiveX("NxTab Control"), "Follow Up")
Call VerTabOpened("Follow Up")
Call VerifyFollowUp()

Call ChangeTabs (Window("Nextech Main Window").Window("Contacts").ActiveX("NxTab Control"), "History")
Call VerTabOpened("History")
Call VerifyHistory()

Sub VerifyGeneral()
 @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf66.xml_;_
'Checking the ellipsis button
Window("Nextech Main Window").Window("Contacts").WinButton("WinButton").Click
'=========
' IsPresent = CheckIfPresent("regexwintitle:=Nextech Main Window").Dialog("regexpwintitle:=Edit Prefixes").WinButton("text:=OK"))

'If IsPresnet Then
' CheckpointValue("Edit checkpoint present")
'If passed then RecordResults(pass,PresentString)
' click something
' verify values on screen, blah blah blah
'else
' RecordResults(fail,PresentString)
'End if




'End If



'Checkpointvalue(IsPresnet,PresentString)
'If passed then RecordResults(pass,PresentString)
'else
' RecordResults(fail,PresentString)
'End if

' CheckIfPresent(DPRecognProps)
'	If DPRecognProps.exists Then
'		IsPresent = True
'		Exit Function
'	IsPresent = False	
'	End If
'=========

'Window("Nextech Main Window").Dialog("Edit Prefixes").Check CheckPoint("Edit Prefixes") @@ hightlight id_;_1311794_;_script infofile_;_ZIP::ssf2.xml_;_
Window("Nextech Main Window").Dialog("Edit Prefixes").WinButton("OK").Click @@ hightlight id_;_7539610_;_script infofile_;_ZIP::ssf3.xml_;_
Dialog("NexTech Practice").WinButton("No").Click @@ hightlight id_;_1381792_;_script infofile_;_ZIP::ssf4.xml_;_
Window("Nextech Main Window").Dialog("Edit Prefixes").WinButton("Cancel").Click @@ hightlight id_;_396388_;_script infofile_;_ZIP::ssf5.xml_;_

' Verifying the Permission Groups Tab
Window("Nextech Main Window").Window("Contacts").WinButton("Permission Groups").Click @@ hightlight id_;_2099164_;_script infofile_;_ZIP::ssf6.xml_;_
'Window("Nextech Main Window").Dialog("Permission Groups").Check Checkpoint("Permission Groups") @@ hightlight id_;_1377330_;_script infofile_;_ZIP::ssf7.xml_;_

'Verifying the New Group within Permission Groups
Window("Nextech Main Window").Dialog("Permission Groups").WinButton("New Group").Click @@ hightlight id_;_461940_;_script infofile_;_ZIP::ssf8.xml_;_

'Verifying the New Permission Group
'Dialog("New Permission Group").Check CheckPoint("New Permission Group") @@ hightlight id_;_1709472_;_script infofile_;_ZIP::ssf9.xml_;_
Dialog("New Permission Group").WinButton("OK").Click @@ hightlight id_;_789428_;_script infofile_;_ZIP::ssf10.xml_;_
Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_461716_;_script infofile_;_ZIP::ssf11.xml_;_
Dialog("New Permission Group").WinButton("Cancel").Click @@ hightlight id_;_1313600_;_script infofile_;_ZIP::ssf12.xml_;_
Window("Nextech Main Window").Dialog("Permission Groups").WinButton("Close").Click @@ hightlight id_;_7605146_;_script infofile_;_ZIP::ssf13.xml_;_

'Verifying the User Properties
Window("Nextech Main Window").Window("Contacts").WinButton("User Properties").Click @@ hightlight id_;_856598_;_script infofile_;_ZIP::ssf14.xml_;_
'Window("Nextech Main Window").Dialog("User Properties").Check CheckPoint("User Properties") @@ hightlight id_;_1442866_;_script infofile_;_ZIP::ssf15.xml_;_

'Configure Groups Dialog
Window("Nextech Main Window").Dialog("User Properties").WinButton("Configure Groups").Click @@ hightlight id_;_461330_;_script infofile_;_ZIP::ssf16.xml_;_
'Dialog("Configure Users").Check CheckPoint("Configure Users") @@ hightlight id_;_920368_;_script infofile_;_ZIP::ssf17.xml_;_
Dialog("Configure Users").WinButton("Close").Click @@ hightlight id_;_2623836_;_script infofile_;_ZIP::ssf18.xml_;_

' Verifying the Configure Password Strength
Window("Nextech Main Window").Dialog("User Properties").WinButton("Configure Password Strength").Click @@ hightlight id_;_1049174_;_script infofile_;_ZIP::ssf19.xml_;_
'Dialog("Login Configuration").Check CheckPoint("Login Configuration") @@ hightlight id_;_1049646_;_script infofile_;_ZIP::ssf20.xml_;_
Dialog("Login Configuration").WinButton("Cancel").Click @@ hightlight id_;_2688028_;_script infofile_;_ZIP::ssf21.xml_;_
Window("Nextech Main Window").Dialog("User Properties").WinButton("Cancel").Click @@ hightlight id_;_1707272_;_script infofile_;_ZIP::ssf22.xml_;_
End Sub
'Clicking Notes on the Contacts Tab
 @@ hightlight id_;_1970348_;_script infofile_;_ZIP::ssf24.xml_;_
Sub VerifyNotes()
	
'Clicking on Add Note
Window("Nextech Main Window").Window("Contacts").WinButton("Add Note").Click @@ hightlight id_;_1642738_;_script infofile_;_ZIP::ssf25.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("Add Note").Click

'Clicking on Delete Note
Window("Nextech Main Window").Window("Contacts").WinButton("Delete Note").Click @@ hightlight id_;_1249520_;_script infofile_;_ZIP::ssf26.xml_;_
'Window("Nextech Main Window").Dialog("Delete?").Check CheckPoint("Delete?") @@ hightlight id_;_1770586_;_script infofile_;_ZIP::ssf27.xml_;_
Window("Nextech Main Window").Dialog("Delete?").WinButton("Yes").Click @@ hightlight id_;_8915586_;_script infofile_;_ZIP::ssf28.xml_;_

'Click on Edit Categories
Window("Nextech Main Window").Window("Contacts").WinButton("Edit Categories").Click @@ hightlight id_;_1837436_;_script infofile_;_ZIP::ssf29.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Check CheckPoint("Note / Follow-Up Categories") @@ hightlight id_;_1836122_;_script infofile_;_ZIP::ssf30.xml_;_

'Click on Adv EMR Merge
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Adv. EMR Merge").Click @@ hightlight id_;_1573934_;_script infofile_;_ZIP::ssf31.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").Check CheckPoint("Advanced Merge - Default Categories") @@ hightlight id_;_9178018_;_script infofile_;_ZIP::ssf32.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").WinButton("Cancel").Click @@ hightlight id_;_3017512_;_script infofile_;_ZIP::ssf33.xml_;_

'Click on Combine Categories
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Combine Categories").Click @@ hightlight id_;_2034952_;_script infofile_;_ZIP::ssf34.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").Check CheckPoint("Combine Note Categories") @@ hightlight id_;_9243554_;_script infofile_;_ZIP::ssf35.xml_;_

'Click on Combine Selected Categories
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").WinButton("Combine Selected Categories").Click @@ hightlight id_;_1376850_;_script infofile_;_ZIP::ssf36.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").WinButton("Close").Click @@ hightlight id_;_985894_;_script infofile_;_ZIP::ssf37.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Close").Click @@ hightlight id_;_1311300_;_script infofile_;_ZIP::ssf38.xml_;_
End Sub

'Verifying the Follow-up Tab

'Window("Nextech Main Window").Window("Contacts").ActiveX("NxTab Control").Click 307,19

Sub VerifyFollowUp()

'Click Create T0-Do Task
Window("Nextech Main Window").Window("Contacts").WinButton("Create To-Do Task").Click @@ hightlight id_;_1184096_;_script infofile_;_ZIP::ssf40.xml_;_
'Window("Nextech Main Window").Dialog("Modify Task").Check CheckPoint("Modify Task") @@ hightlight id_;_461876_;_script infofile_;_ZIP::ssf59.xml_;_
Window("Nextech Main Window").Dialog("Modify Task").WinButton("Cancel").Click @@ hightlight id_;_8653722_;_script infofile_;_ZIP::ssf41.xml_;_

'Click the View To-Do List
Window("Nextech Main Window").Window("Contacts").WinButton("View To-Do List").Click @@ hightlight id_;_1313154_;_script infofile_;_ZIP::ssf42.xml_;_
'Dialog("To Do Alarm - 0 items").Check CheckPoint("To Do Alarm - 0 items") @@ hightlight id_;_6886210_;_script infofile_;_ZIP::ssf60.xml_;_
Dialog("To Do Alarm - 0 items").WinButton("Close").Click @@ hightlight id_;_1053178_;_script infofile_;_ZIP::ssf43.xml_;_

'Click the Delete Completed
Window("Nextech Main Window").Window("Contacts").WinButton("Delete Completed").Click @@ hightlight id_;_1839216_;_script infofile_;_ZIP::ssf44.xml_;_
'Window("Nextech Main Window").Dialog("NexTech Practice").Check CheckPoint("NexTech Practice") @@ hightlight id_;_11930546_;_script infofile_;_ZIP::ssf45.xml_;_
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_2427406_;_script infofile_;_ZIP::ssf46.xml_;_

'Click the Edit Categories
Window("Nextech Main Window").Window("Contacts").WinButton("Edit Categories").Click @@ hightlight id_;_2166668_;_script infofile_;_ZIP::ssf47.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Check CheckPoint("Note / Follow-Up Categories_2") @@ hightlight id_;_11996082_;_script infofile_;_ZIP::ssf48.xml_;_

'Click New
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("New").Click @@ hightlight id_;_2362178_;_script infofile_;_ZIP::ssf49.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Input").Check CheckPoint("Input") @@ hightlight id_;_1377988_;_script infofile_;_ZIP::ssf50.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Input").WinButton("Cancel").Click @@ hightlight id_;_1707124_;_script infofile_;_ZIP::ssf51.xml_;_

'Click Adv. EMR Merge
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Adv. EMR Merge").Click @@ hightlight id_;_1968948_;_script infofile_;_ZIP::ssf52.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").Check CheckPoint("Advanced Merge - Default Categories_2") @@ hightlight id_;_1443524_;_script infofile_;_ZIP::ssf53.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").WinButton("Cancel").Click @@ hightlight id_;_2165590_;_script infofile_;_ZIP::ssf54.xml_;_

'Click Combine Categories
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Combine Categories").Click @@ hightlight id_;_3475804_;_script infofile_;_ZIP::ssf55.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").Check CheckPoint("Combine Note Categories_2") @@ hightlight id_;_1509060_;_script infofile_;_ZIP::ssf56.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").WinButton("Close").Click @@ hightlight id_;_2034712_;_script infofile_;_ZIP::ssf57.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Close").Click @@ hightlight id_;_2100032_;_script infofile_;_ZIP::ssf58.xml_;_

End Sub

'Verifying the History Tab

Sub VerifyHistory()


'Click on New Button
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf63.xml_;_

'Click on the WinMenu Attach Existing Folder
WinMenu("ContextMenu").Select "Attach Existing Folder"
'Window("Nextech Main Window").Dialog("Browse for Folder").Check CheckPoint("Browse for Folder") @@ hightlight id_;_1641456_;_script infofile_;_ZIP::ssf64.xml_;_
Window("Nextech Main Window").Dialog("Browse for Folder").WinButton("Cancel").Click @@ hightlight id_;_592944_;_script infofile_;_ZIP::ssf65.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf66.xml_;_

'Click on Import and Attach existing File
WinMenu("ContextMenu").Select "Import and Attach Existing File"
Window("Nextech Main Window").Dialog("Open").WinButton("Cancel").Click @@ hightlight id_;_1115220_;_script infofile_;_ZIP::ssf67.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf68.xml_;_

'Click on WinMenu Import from Scanner
WinMenu("ContextMenu").Select "Import from Scanner/Camera;Scan as Image..."
'Window("Nextech Main Window").Dialog("NexTech Practice").Check CheckPoint("NexTech Practice_2") @@ hightlight id_;_6556292_;_script infofile_;_ZIP::ssf69.xml_;_
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_920326_;_script infofile_;_ZIP::ssf70.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf71.xml_;_

'Click on WinMenu Import from PDA
WinMenu("ContextMenu").Select "Import from PDA"
'Window("Nextech Main Window").Dialog("NexTech Practice").Check CheckPoint("NexTech Practice_3") @@ hightlight id_;_1573986_;_script infofile_;_ZIP::ssf72.xml_;_
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_3541826_;_script infofile_;_ZIP::ssf73.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf74.xml_;_

'Click on WinMenu Import from Device
WinMenu("ContextMenu").Select "Import from Device"
'Dialog("Device Import").Check CheckPoint("Device Import") @@ hightlight id_;_791162_;_script infofile_;_ZIP::ssf75.xml_;_
Dialog("Device Import").WinButton("Close").Click @@ hightlight id_;_922382_;_script infofile_;_ZIP::ssf76.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf77.xml_;_

'Click on WinMenu Create new Document
WinMenu("ContextMenu").Select "Create New Document"
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf78.xml_;_

'Click on Merge New Document
WinMenu("ContextMenu").Select "Merge New Document"
Window("Nextech Main Window").Dialog("Open").WinButton("Cancel").Click @@ hightlight id_;_1182470_;_script infofile_;_ZIP::ssf79.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("New").Click @@ hightlight id_;_2168316_;_script infofile_;_ZIP::ssf80.xml_;_

'Click on Mege New Packet
WinMenu("ContextMenu").Select "Merge New Packet"
Window("Nextech Main Window").Dialog("Select Packet").WinButton("WinButton").Click @@ hightlight id_;_1248304_;_script infofile_;_ZIP::ssf81.xml_;_

'Clicking on the ellipsis button
'Dialog("Configure Letter Writing").Check CheckPoint("Configure Letter Writing Packets") @@ hightlight id_;_1313542_;_script infofile_;_ZIP::ssf82.xml_;_

'Click on Add
Dialog("Configure Letter Writing").WinButton("Add").Click @@ hightlight id_;_1573988_;_script infofile_;_ZIP::ssf83.xml_;_
'Dialog("Input").Check CheckPoint("Input_2") @@ hightlight id_;_1115270_;_script infofile_;_ZIP::ssf84.xml_;_
Dialog("Input").WinButton("Cancel").Click @@ hightlight id_;_2296932_;_script infofile_;_ZIP::ssf85.xml_;_

'Click on Add Copy
Dialog("Configure Letter Writing").WinButton("Add Copy").Click @@ hightlight id_;_1901742_;_script infofile_;_ZIP::ssf86.xml_;_
'Dialog("NexTech Practice").Check CheckPoint("NexTech Practice_4") @@ hightlight id_;_1180806_;_script infofile_;_ZIP::ssf87.xml_;_
Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_4128922_;_script infofile_;_ZIP::ssf88.xml_;_
Dialog("Configure Letter Writing").WinButton("Close").Click @@ hightlight id_;_1377344_;_script infofile_;_ZIP::ssf89.xml_;_

'Click on Merge on the select packet Dialog
Window("Nextech Main Window").Dialog("Select Packet").WinButton("Merge").Click @@ hightlight id_;_4461184_;_script infofile_;_ZIP::ssf90.xml_;_
Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_2032814_;_script infofile_;_ZIP::ssf91.xml_;_
Window("Nextech Main Window").Dialog("Select Packet").WinButton("Cancel").Click @@ hightlight id_;_2229486_;_script infofile_;_ZIP::ssf92.xml_;_

'Click on Import from Scanner on the History tab
Window("Nextech Main Window").Window("Contacts").WinButton("Import from Scanner/Camera").Click @@ hightlight id_;_1446512_;_script infofile_;_ZIP::ssf93.xml_;_
WinMenu("ContextMenu").Select "Scan as Image..."
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_1311784_;_script infofile_;_ZIP::ssf94.xml_;_

'Checking different WinMenu on the Import from Scanner Button
Window("Nextech Main Window").Window("Contacts").WinButton("Import from Scanner/Camera").Click @@ hightlight id_;_1446512_;_script infofile_;_ZIP::ssf95.xml_;_
WinMenu("ContextMenu").Select "Scan as PDF..."
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_2884616_;_script infofile_;_ZIP::ssf96.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("Import from Scanner/Camera").Click @@ hightlight id_;_1446512_;_script infofile_;_ZIP::ssf97.xml_;_
WinMenu("ContextMenu").Select "Scan as Multi-Page PDF..."
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_3081338_;_script infofile_;_ZIP::ssf98.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("Import from Scanner/Camera").Click @@ hightlight id_;_1446512_;_script infofile_;_ZIP::ssf99.xml_;_

'Clicking on Twain Input Source within the Import From Scanner
WinMenu("ContextMenu").Select "Select TWAIN Input Source..."
'Dialog("Select Source").Check CheckPoint("Select Source") @@ hightlight id_;_2884846_;_script infofile_;_ZIP::ssf100.xml_;_
Dialog("Select Source").WinButton("Cancel").Click @@ hightlight id_;_1641520_;_script infofile_;_ZIP::ssf101.xml_;_

'Clicking on Record Audio Tab
Window("Nextech Main Window").Window("Contacts").WinButton("Record Audio").Click @@ hightlight id_;_1642884_;_script infofile_;_ZIP::ssf102.xml_;_
'Window("Nextech Main Window").Dialog("NexTech Practice").Check CheckPoint("NexTech Practice_5") @@ hightlight id_;_2950382_;_script infofile_;_ZIP::ssf103.xml_;_
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_2032794_;_script infofile_;_ZIP::ssf104.xml_;_

'Clicking on Open Default Folder
Window("Nextech Main Window").Window("Contacts").WinButton("Open Default Folder").Click @@ hightlight id_;_1971074_;_script infofile_;_ZIP::ssf105.xml_;_
Window(".*").Close

'Clicking on Edit Categories
Window("Nextech Main Window").Window("Contacts").WinButton("Edit Categories_2").Click @@ hightlight id_;_4065652_;_script infofile_;_ZIP::ssf106.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Check CheckPoint("Note / Follow-Up Categories_3") @@ hightlight id_;_3474554_;_script infofile_;_ZIP::ssf107.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("New").Click @@ hightlight id_;_1838128_;_script infofile_;_ZIP::ssf108.xml_;_

'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Input").Check CheckPoint("Input_3") @@ hightlight id_;_1641444_;_script infofile_;_ZIP::ssf109.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Input").WinButton("Cancel").Click @@ hightlight id_;_3212362_;_script infofile_;_ZIP::ssf110.xml_;_

'Clicking on Adv EMR Merge
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Adv. EMR Merge").Click @@ hightlight id_;_3345302_;_script infofile_;_ZIP::ssf111.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").Check CheckPoint("Advanced Merge - Default Categories_3") @@ hightlight id_;_1706980_;_script infofile_;_ZIP::ssf112.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Advanced Merge - Default").WinButton("Cancel").Click @@ hightlight id_;_1836072_;_script infofile_;_ZIP::ssf113.xml_;_

'Clicking On Combine Catogeries within Edit Categories
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Combine Categories").Click @@ hightlight id_;_1508394_;_script infofile_;_ZIP::ssf114.xml_;_
'Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").Check CheckPoint("Combine Note Categories_3") @@ hightlight id_;_1772516_;_script infofile_;_ZIP::ssf115.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").Dialog("Combine Note Categories").WinButton("Close").Click @@ hightlight id_;_3542074_;_script infofile_;_ZIP::ssf116.xml_;_
Window("Nextech Main Window").Dialog("Note / Follow-Up Categories").WinButton("Close").Click @@ hightlight id_;_1836160_;_script infofile_;_ZIP::ssf117.xml_;_


Window("Nextech Main Window").Window("Contacts").ActiveX("NexTech DataList Control").Click 36,29
Window("Nextech Main Window").Window("Contacts").WinButton("Detach File(s)").Click @@ hightlight id_;_1641938_;_script infofile_;_ZIP::ssf119.xml_;_
WinMenu("ContextMenu").Select "Detach File"
 @@ hightlight id_;_1836190_;_script infofile_;_ZIP::ssf120.xml_;_
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("Yes").Click @@ hightlight id_;_3278086_;_script infofile_;_ZIP::ssf121.xml_;_
Window("Nextech Main Window").Window("Contacts").WinButton("Import from Device").Click @@ hightlight id_;_1249396_;_script infofile_;_ZIP::ssf122.xml_;_
'Dialog("Device Import").Check CheckPoint("Device Import_2") @@ hightlight id_;_791162_;_script infofile_;_ZIP::ssf123.xml_;_
Dialog("Device Import").WinButton("Close").Click @@ hightlight id_;_922382_;_script infofile_;_ZIP::ssf124.xml_;_

End Sub


Sub VerTabOpened (TabName)
	Select Case TabName
		Case "General"
		
		Case "Notes"
		
		Case "Follow Up"
		
		Case "History"
	End Select
End Sub




