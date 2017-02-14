Dim TestDBName : TestDBName = Environment.Value("DBUnderTest")
'Dim TestDBName : TestDBName = "UFT12200DB"
Dim IsPresent

RunAction "LogIn [GlobalPracticeActions]", oneIteration, TestDBName

' Verify if Contacts Module opens
If Dialog("regexpwndtitle:= 0 items").WinButton("regexpwndtitle:=Close").Exist(2) Then
'Dialog("regexpwndtitle:= 0 items").WinButton("regexpwndtitle:=Print Preview").Click
Dialog("regexpwndtitle:= 0 items").WinButton("regexpwndtitle:=Close").Click 
	
End If



Window("Nextech.*").WinMenu("Menu").Select "Modules;Contacts"

wait(2)


Call ChangeTabs (Window("Nextech.*").Window("Contacts").ActiveX("NxTab Control"), "General")
Call VerTabOpened ("General")
Call VerifyGeneral()

Call ChangeTabs (Window("Nextech.*").Window("Contacts").ActiveX("NxTab Control"), "Notes")
Call VerTabOpened("Notes")
Call VerifyNotes()


Call ChangeTabs (Window("Nextech.*").Window("Contacts").ActiveX("NxTab Control"), "Follow Up")
Call VerTabOpened("Follow Up")
Call VerifyFollowUp()

Call ChangeTabs (Window("Nextech.*").Window("Contacts").ActiveX("NxTab Control"), "History")
Call VerTabOpened("History")
Call VerifyHistory()

Window("Nextech.*").WinMenu("Menu").Select "Modules;Patients"

wait(1)
'RunAction "LogOut [GlobalPracticeActions]", oneIteration

RunAction "LogOut [GlobalPracticeActions]", oneIteration

Sub VerifyGeneral()

If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").Static("regexpwndtitle:=Contact Type:  ").Exist(2) Then
	Reporter.ReportEvent micPass, "General Tab Open", ""
Else
	Reporter.ReportEvent micFail, "General Tab failed to open",""
End If	

If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndclass:=Button", "index:=0").Exist Then
	Reporter.ReportEvent micPass, "Ellipsis Present", ""
Else
	Reporter.ReportEvent micFail, "Ellipsis Missing", ""
End If

'Clicking on the ... button
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndclass:=Button", "index:=0").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Edit Prefixes").Exist Then
	Reporter.ReportEvent micPass, "Edit Prefixes Dialog is Present", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Prefixes Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Edit Prefixes").WinButton("regexpwndtitle:=OK").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Exist Then
	CheckNo = Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Click
Else
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
End If
	
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Edit Prefixes").WinButton("regexpwndtitle:=Add").Click
'Both adds needed
wait(1)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Edit Prefixes").WinButton("regexpwndtitle:=Add").Click	

If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist(2) Then
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
End If

'Clicking on Remove Button
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Edit Prefixes").WinButton("regexpwndtitle:=Remove").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Edit Prefixes").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").click
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Edit Prefixes").WinButton("regexpwndtitle:=Cancel").Click

'12/29/16 Modifying the script to include Active Directory Setup functionality
'Clicking on Active Directory Setup

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Active Directory Setup").Click

If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").Exist Then
	Reporter.ReportEvent micPass, "Active Directory Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Active Directory Dialog is not Present"
End If

'This is Enable the 3 Buttons in the Active Directory Authentication Setup Dialog box
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").WinCheckBox("regexpwndtitle:=Enable Active Directory Authentication").Click

'Clicking on Configure Group Permissions...

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").WinButton("regexpwndtitle:=Configure &Group Permissions\.\.\.").Click

If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").Dialog("regexpwndtitle:=Configure Active Directory Group Permissions").Exist Then
	Reporter.ReportEvent micPass, "Configure Active Directory Group Permissions is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure Active Directory Group Permissions is not Present"
End If

'Selecting the Domain User from the datalist
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").Dialog("regexpwndtitle:=Configure Active Directory Group Permissions").ActiveX("acx_name:=NexTech DataList Control 2\.0").Type "Domain Users"

'Clicking on the Add Permission Groups Within the Configure Active Directory Group Dialog
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").Dialog("regexpwndtitle:=Configure Active Directory Group Permissions").WinButton("regexpwndtitle:=&Add Permission Groups").Click

If Dialog("regexpwndtitle:=Multiple Selection").Exist Then
	Reporter.ReportEvent micPass, "Multiple Selection Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Multiple Selection Dialog is not Present"
End If

Dialog("regexpwndtitle:=Multiple Selection").ActiveX("acx_name:=NexTech DataList Control 2\.0").Type "Auditors"

Dialog("regexpwndtitle:=Multiple Selection").WinButton("regexpwndtitle:=OK").Click

'Removing the Permission Group from Domain Users Group
Dialog("regexpwndtitle:=Configure Active Directory Group Permissions").ActiveX("acx_name:=NexTech DataList Control 2\.0").Type "Domain Users"
Dialog("regexpwndtitle:=Configure Active Directory Group Permissions").WinButton("regexpwndtitle:=&Remove Permission Groups").Click

If Dialog("regexpwndtitle:=Multiple Selection").Exist Then
	Reporter.ReportEvent micPass, "Multiple Selection Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Multiple Selection Dialog is not Present"
End If

'Unselecting the Check box next to the selected group
Dialog("regexpwndtitle:=Multiple Selection").ActiveX("acx_name:=NexTech DataList Control 2\.0").click 15,20

Dialog("regexpwndtitle:=Multiple Selection").WinButton("regexpwndtitle:=OK").Click

'Closing the configure Active Directory Group Permissions Dialog box
Dialog("regexpwndtitle:=Configure Active Directory Group Permissions").WinButton("regexpwndtitle:=&Close").Click

'Clicking on Configure User Login Settings
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").WinButton("regexpwndtitle:=Configure &User Login Settings\.\.\.").Click

If Dialog("regexpwndtitle:=Configure User Login Settings").Exist Then
	Reporter.ReportEvent micPass, "Configure User Login Settings Dialog is Present",""
Else	
	Reporter.ReportEvent micFail, "Test Failed", "Configure User Loging Setting Dialog is not Present"
End If

'Closing the Configure User Login Dialog 
Dialog("regexpwndtitle:=Configure User Login Settings").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Link Existing Users to Windows Users Button
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").WinButton("regexpwndtitle:=&Link Existing Users to Windows Users\.\.\.").Click
If Dialog("regexpwndtitle:=Link Existing Users To Active Directory Users").WinButton("regexpwndtitle:=&Link Selected Usernames").Exist Then
	Reporter.ReportEvent micPass, "Link Existing Users to Active Directory Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Link Existing Users to Active Directory Dialog is not Present"
End If

Dialog("regexpwndtitle:=Link Existing Users To Active Directory Users").ActiveX("acx_name:=NexTech DataList Control 2\.0","index:=0").Type "Administrator"

'Verifying if the Link Selected Username Button works
Dialog("regexpwndtitle:=Link Existing Users To Active Directory Users").WinButton("regexpwndtitle:=&Link Selected Usernames").Click

If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=Link Usernames").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If

Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=Link Usernames").Click

'Closing the Link Existing Users To Active Directory Dialog
Dialog("regexpwndtitle:=Link Existing Users To Active Directory Users").WinButton("regexpwndtitle:=Close").Click

'Closing the Active Directory Authentication Setup Dialog Box
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").WinButton("regexpwndtitle:=Close").Click

'Unlinked the previously linked user from User Properties
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=User Properties").Click

If Dialog("regexpwndtitle:=User Properties").Exist Then
	Reporter.ReportEvent micPass, "User Properties Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "User Properties Dialog is not Present"
End If

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=User Properties","index:=1").WinButton("regexpwndtitle:=Unlink Active Directory User\.\.\.").Click

If Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If

Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=Unlink User from Active Directory").Click

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=User Properties","index:=1").WinButton("regexpwndtitle:=OK").Click

'Unchecking the Enable Active Directory Authentication Check box
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Active Directory Setup").Click

If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").Exist Then
	Reporter.ReportEvent micPass, "Active Directory Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Active Directory Dialog is not Present"
End If

'This is to unable the 3 Buttons in the Active Directory Authentication Setup Dialog box
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").WinCheckBox("regexpwndtitle:=Enable Active Directory Authentication").Click

If Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
	
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=Disable Active Directory Authentication").Click

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Active Directory Authentication Setup").WinButton("regexpwndtitle:=Close").Click


'Clicking on Permission Groups

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Permission Groups").Click

If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Permission Groups").Exist Then
	Reporter.ReportEvent micPass, "Permission Group Dialog Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Permission Group Dialog Not Open"	
End If

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Permission Groups", "index:=1").WinButton("regexpwndtitle:=New Group").Click

If Dialog("regexpwndtitle:=New Permission Group").Exist Then
	Reporter.ReportEvent micPass, "New Permission Group Dialog Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "New Permission Group Dialog Not Open"		
End If

' Verifying New Permission group
Dialog("regexpwndtitle:=New Permission Group").WinButton("regexpwndtitle:=Cancel").Click
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Permission Groups", "index:=1").WinButton("regexpwndtitle:=Help").Click
	If Window("regexpwndtitle:=NexTech Practice User Manual").Exist Then
		Reporter.ReportEvent micPass, "Nextech Practice User Manual Window is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice User Manual Window is not Present"
	End If
	
Window("regexpwndtitle:=NexTech Practice User Manual").Close
wait(.5)

'Closing Permission Group
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Permission Groups", "index:=1").WinButton("regexpwndtitle:=Close").Click
wait(.5)

'Clicking User properties
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=User Properties").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=User Properties", "index:=1").Exist Then
		Reporter.ReportEvent micPass, "User Properties Dialog is Present",""
Else
		Reporter.ReportEvent micFail, "Test Failed","User Properties Dialog is not Present"
End If
	
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=User Properties", "index:=1").WinButton("regexpwndtitle:=Configure Groups").Click
If Dialog("regexpwndtitle:=Configure Users").WinButton("regexpwndtitle:=Group Configuration").Exist Then
	Reporter.ReportEvent micPass, "Configure User Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure User Dialog is not Present"
End If

wait(.5)
Dialog("regexpwndtitle:=Configure Users").WinButton("regexpwndtitle:=Group Configuration").Click
wait(.5)
Dialog("regexpwndtitle:=Configure Permission Groups").WinButton("regexpwndtitle:=User Configuration").Click
wait(.5)
Dialog("regexpwndtitle:=Configure Users").WinButton("regexpwndtitle:=Group Configuration").Click
wait(.5)
Dialog("regexpwndtitle:=Configure Permission Groups").WinButton("regexpwndtitle:=Close").Click

'Clicking on Configure Password Strength
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=User Properties", "index:=1").WinButton("regexpwndtitle:=Configure Password Strength and Login Options").Click
If Dialog("regexpwndtitle:=Login Configuration").WinButton("regexpwndtitle:=Cancel").Exist Then
	Reporter.ReportEvent micPass, "Login Configuration Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Login Configuration Dialog is not Present"
End If	
Dialog("regexpwndtitle:=Login Configuration").WinButton("regexpwndtitle:=Cancel").Click

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=User Properties", "index:=1").WinButton("regexpwndtitle:=Cancel").Click

End Sub

Sub VerifyNotes()
	
'Verifying the Notes Tab in Contacts Module
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Add &Note").Exist Then
	Reporter.ReportEvent micPass, "Notes Tab is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Notes Tab Not Open"
End If

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Add &Note").Click
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Add &Note").Click
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=&Delete Note").Click
If Dialog("regexpwndtitle:=Delete\?").Exist Then
	Reporter.ReportEvent micPass, "Delete Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Delete Dialog is not Present"
End If
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on the Edit Categories
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=&Edit Categories").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&New").Exist Then
	Reporter.ReportEvent micPass, "Note/Follow-Up Categories Dialog Box is Open", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Note/Follow-up Categories Dialog Box is Not Open"
End If

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&New").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories", "index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories", "index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Adv. EMR Merge...
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&Adv\. EMR Merge\.\.\.").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:= Default Categories", "index:=1").Exist Then
	Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:= Default Categories", "index:=1").WinButton("regexpwndtitle:=OK", "index:=1").Click
	Reporter.ReportEvent micPass, "Advanced Merge - Default Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Advanced Merge - Default Categories Dialog is not Present"
End If
'Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:= Default Categories", "index:=1").WinButton("regexpwndtitle:=OK", "index:=1").Click

'Clicking on Combine Categories..
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&Combine Categories\.\.\.").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Combine Note Categories").WinButton("regexpwndtitle:=Close").Exist Then
	Reporter.ReportEvent micPass, "Combine Note Categories Dialog Box Opened",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Combine Note Categories Dialog Box Not Found"
End If

wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Combine Note Categories").WinButton("regexpwndtitle:=Close").Click
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=Clos&e").Click

End Sub

Sub VerifyFollowUp()

'Verifying the Follow Up tab in Contact module	
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Create &To-Do Task").Exist Then
	Reporter.ReportEvent micPass, "Follow Up Tab is open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Follow Up Tab Not Found"
End If	

'Clicking on the Create To-Do Task
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Create &To-Do Task").Click
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Modify Task").Exist Then
	wait(.5)
	Reporter.ReportEvent micPass, "Modify Task Dialog is Present",""
	Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Modify Task").WinButton("regexpwndtitle:=Cancel").Click
	wait(.5)
Else
	Reporter.ReportEvent micFail, "Test Failed","Modify Task Dialog is not Present"
End If

'Clicking on View To-Do List
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=View To-&Do List").Click
If Dialog("regexpwndtitle:= 0 items").Exist Then
	Reporter.ReportEvent micPass, "To Do Alarm Dialog is Present",""
	wait(.5)
'Clicking on Print Preview Button
	Dialog("regexpwndtitle:= 0 items").WinButton("regexpwndtitle:=Print Preview").Click
	wait(.5)
	If window("Nextech.*").Window("regexpwndtitle:=To-Do Alarm Preview \(PP\)").Exist Then
		Reporter.ReportEvent micPass, "To do Alarm PP Window is Present",""
		wait(.5)
		Window("Nextech.*").Window("regexpwndtitle:=To-Do Alarm Preview \(PP\)").Close
	Else
		Reporter.ReportEvent micFail, "Test Failed", "To do Alarm PP Window is not Present"
	End If
Else
	Reporter.ReportEvent micFail, "Test Failed", "To Do Alarm is not Present"
End If

'Dialog("regexpwndtitle:= 0 items").WinButton("regexpwndtitle:=Close").Click

'Clicking on the Delete Completed Button
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=&Delete Completed").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "NexTech Practice Dialog is Present",""
	wait(.5)
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Else
	Reporter.ReportEvent micFail, "NexTech Practice Dialog was expected but wasn't Present",""
End If

Call VerifyEditCategories()


End Sub

Sub VerifyHistory()
	
'Verifying the History Tab in Contacts Module
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Exist(2) Then
	Reporter.ReportEvent micPass, "History Tab is open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "History Tab Not Open"
End If

'Clicking on the WinMenu content of New Button
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Click
Window("Nextech.*").WinMenu("ContextMenu").Select "Attach Existing Folder"
wait(.1)
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Browse for Folder").Exist Then
	Reporter.ReportEvent micPass, "Browse for folder Dialog is Present",""
	Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Browse for Folder").WinButton("regexpwndtitle:=Cancel").Click
Else
	Reporter.ReportEvent micFail, "Test Failed", "Browse for Folder Dialog is not Present"
End If

'Clicking On WinMenu Import and Attach Existing File
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Click
Window("Nextech.*").WinMenu("ContextMenu").Select "Import and Attach Existing File"
wait(.1)
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Open Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Open Dialog box is not Present"
End If

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Open","index:=1").WinButton("regexpwndtitle:=Cancel").Click
	wait(2)
	
'Clicking on WinMenu Import from PDA
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Click
wait(1)
Window("Nextech.*").WinMenu("ContextMenu").Select "Import from PDA"
wait(1)

' <<<< Practice has stopped working here >>>>

If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
' <<<< Practice has stopped working here >>>>
wait(1)
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Import From Device
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Click
Window("Nextech.*").WinMenu("ContextMenu").Select "Import from Device"
wait(.5)
If Dialog("regexpwndtitle:=Device Import").Exist(2) Then
	Reporter.ReportEvent micPass, "Device Import Dialog Opens",""
	Dialog("regexpwndtitle:=Device Import").WinButton("regexpwndtitle:=&Close").Click
Else
	Reporter.ReportEvent micFail, "Test Failed", "Device Import Dialog does Not Open"
End If
wait(1)

'Clicking on WinMenu Create New Document
'needs to be modified for both word and remoteapp
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Click
wait(1)
Window("Nextech.*").WinMenu("ContextMenu").Select "Create New Document"
wait(5)
If Window("regexpwndtitle:= Word").Exist Then
	Reporter.ReportEvent micPass, "Word Document is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Word Document is not Open"
End If
Window("regexpwndtitle:= Word").Close
'
'Clicking on WinMenu Merge New Document
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Click
wait(1)
Window("Nextech.*").WinMenu("ContextMenu").Select "Merge New Document"
wait(1)
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Open Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Open Dialog is not Present"
End If
wait(1)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on WinMenu Merge New Packet
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=New").Click
Window("Nextech.*").WinMenu("ContextMenu").Select "Merge New Packet"
wait(.5)
If Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Select Packet").Exist Then
	Reporter.ReportEvent micPass, "Select packet Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Select Packet Dialog is not Present"
End If

'Verifying the Select Packet Dialog works properly
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Select Packet").WinButton("regexpwndtitle:=\.\.\.").Click
If Dialog("regexpwndtitle:=Configure Letter Writing Packets").WinButton("regexpwndtitle:=Add").Exist Then
	Reporter.ReportEvent micPass, "Configure Letter Writing Packets Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure Letter Writing Packets Dialog is not Present"
End If
Dialog("regexpwndtitle:=Configure Letter Writing Packets").WinButton("regexpwndtitle:=Add").Click
If Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click
Dialog("regexpwndtitle:=Configure Letter Writing Packets").WinButton("regexpwndtitle:=Add Copy").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'ManagePopUp CheckYes, CheckNo, CheckOK
'Clciking the Select Packet Dialog
Dialog("regexpwndtitle:=Configure Letter Writing Packets").WinButton("regexpwndtitle:=Close").Click
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Select Packet").WinButton("regexpwndtitle:=Merge").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
wait(.5)
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("regexpwndtitle:=Select Packet").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Import From Scanner/Camera Tab
wait(.5)
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
wait(.5)
Window("Nextech.*").WinMenu("ContextMenu").Select "Scan as Image..."
wait(.1)

Call VerScanner()
	
'Clicking on WinMenu Scan as PDF
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
wait(.5)
Window("Nextech.*").WinMenu("ContextMenu").Select "Scan as PDF..."
wait(.1)

Call VerScanner()


'Clicking on WinMenu Multi-Page PDF...
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
wait(.5)
Window("Nextech.*").WinMenu("ContextMenu").Select "Scan as Multi-Page PDF..."
wait(.1)

Call VerScanner()

''Clicking on WinMenu Select TWAIN Input Source
'Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Import from Scanner/Camera").Click
'Window("Nextech.*").WinMenu("ContextMenu").Select "Select TWAIN Input Source..."
'wait(.1)
'
'Call VerScanner()
'
'Clicking on Record Audio
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Record Audio").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist(2) Then
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
ElseIf Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("text:=NexTech Contact Audio Recording for.*").Exist(2) Then
	Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Dialog("text:=NexTech Contact Audio Recording for.*").WinButton("regexpwndtitle:=Discard").Click
	Reporter.ReportEvent micPass, "Audio Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Unexpected Audio problem"
End If

'Clicking on Open Default Folder
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Open Default Folder").Click
If Window("regexpwndclass:=CabinetWClass").Exist Then
	Reporter.ReportEvent micPass, "Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Dialog is not Present"
End If
Window("regexpwndclass:=CabinetWClass","index:=0").Close


'Clicking on Edit Categories
Call VerifyEditCategories()
wait(2)

'Clicking on the Detach File(s) Button
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinObject("regexpwndclass:=Afx:").Click 450,26
If Window("Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Exist Then
	Reporter.ReportEvent micPass, "History Item Notes Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "History Item Notes Dialog is not Present"
End If

'Clicking on Add Macro in History Item Notes Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").WinButton("regexpwndtitle:=A&dd Macro").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Exist Then
	Reporter.ReportEvent micPass, "Add Macro Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Add Macro Dialog is not Present"
End If

'Clicking on Edit Macros Button in Add Macro Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").WinButton("regexpwndtitle:=Edit Macros").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Edit Macros Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Macros Dialog is not Present"
End If

'Clicking on Add macros button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Add Macro").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Type "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

'Clicking on Delete macros Button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Delete Macro").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed","Nextech Practice Dialog is not Present"
End If

Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Closing the Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Closing the Add Macro Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Add Macro").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Macros on Histoy Item Notes Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").WinButton("regexpwndtitle:=Edit Macros").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Edit Macro Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Macro Dialog is not Present"
End If

'Clicking on Add macros button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Add Macro").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Type "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

'Clicking on Delete macros Button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Delete Macro").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed","Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Closing the Edit Macros Dialog box
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Categories Button

'Call VerifyEditCategories()

Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").WinButton("regexpwndtitle:=Edit Cate&gories").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Exist Then
	Reporter.ReportEvent micPass, "Note/ Follow-Up Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Note/ Follow-Up Categories Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&New").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Type "ABCD"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").WinObject("object class:=Afx:.*:.*").Click 11,33

'Clicking on the Delete Button in Note/Follow Up Categories
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=&Delete").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Adv. EMR Merge...
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=&Adv\. EMR Merge\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:= Default Categories").Exist Then
	Reporter.ReportEvent micPass, "Default Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Default Categories Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:= Default Categories").WinButton("regexpwndtitle:=Cancel","index:=1").Click

'Clicking on Combine Categories...
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=&Combine Categories\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=Combine Note Categories").Exist Then
	Reporter.ReportEvent micPass, "Combine Note Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Combine Note Categories Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=Combine Note Categories").WinButton("regexpwndtitle:=Close").Click

'Closing the Note/Follow-Up Categories
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=Clos&e").Click
'
'Clicking on Search Notes
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=History Item Notes").WinButton("regexpwndtitle:=Search Notes").Click

If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If

Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Search Notes For:.*").Exist Then
	Reporter.ReportEvent micPass, "Search Notes for Patient Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Notes for Patient Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=History Item Notes").Dialog("regexpwndtitle:=Search Notes For:.*").WinButton("regexpwndtitle:=Close").Click

'Closing the Bill Notes Dialog 
Set dr = CreateObject("Mercury.DeviceReplay")
    dr.KeyDown 56 'PRESS THE CONTROL KEY DOWN & HOLD "56" IS ASCII CODE FOR "Left ALT"
    dr.KeyDown 62 'PRESS THE SHIFT KEY DOWN & HOLD "42" IS ASCII CODE FOR "F4"
    wait 2
    dr.KeyUp 56 'RELEAS THE CONTROL KEY "29" IS ASCII CODE FOR "Left ALT"
    dr.KeyUp 62 'PRESS THE SHIFT KEY "42" IS ASCII CODE FOR "F4"
    wait 2 ' MAKE SURE THE KEYS ARE RELEASED
Set dr = Nothing
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Detach File\(s\)").Click
Window("Nextech.*").WinMenu("ContextMenu").Select "Detach And Delete File"
	If Dialog("regexpwndtitle:=NexTech Practice").Exist Then
		Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click
		Reporter.ReportEvent micPass, "Nextech practice Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
	End If

'Clicking on Import from Device
Window("regexpwndtitle:=Nextech \(UFT.*DB\)").Window("regexpwndtitle:=Contacts").WinButton("regexpwndtitle:=Import from Device").Click
If Dialog("regexpwndtitle:=Device Import").Exist Then
	Reporter.ReportEvent micPass, "Device Import Dialog Opens",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Device Import Dialog does Not Open"
End If
Dialog("regexpwndtitle:=Device Import").WinButton("regexpwndtitle:=&Close").Click

End Sub

Sub VerTabOpened (TabName)
	Select Case TabName
		Case "General"
		
		Case "Notes"
		
		Case "Follow Up"
		
		Case "History"
	End Select
End Sub

Sub VerScanner()

If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist(2) Then
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
	
ElseIf	Dialog("regexpwndtitle:=Integrated Camera TWAIN").Exist(2) Then
	Dialog("regexpwndtitle:=Integrated Camera TWAIN").WinButton("regexpwndtitle:=Exit").Click
	Reporter.ReportEvent micPass, "Twain opened", "Twain found"
	
ElseIf Dialog("regexpwndtitle:=HP TWAIN.*").Exist Then
	Dialog("regexpwndtitle:=HP TWAIN.*").WinButton("regexpwndtitle:=Cancel").Click
	Reporter.ReportEvent micPass,"HP Twain Scanner is Found",""
		
Else
	Reporter.ReportEvent micFail, "unexpect result for import device", "not train or nt practice?"
End If

If Dialog("NexTech Practice").Exist Then
	Dialog("NexTech Practice").WinButton("OK").Click
End If

End Sub










