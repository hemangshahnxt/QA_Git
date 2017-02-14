Dim TestDBName : TestDBName = Environment.Value("DBUnderTest")

Dim IsPresent
Dim StartUpMethod
StartUpMethod = Environment.Value("StartUpMethod")
'Log in to practice


RunAction "LogIn [GlobalPracticeActions]", oneIteration
wait (2)

If Not Window("regexpwndtitle:=Nextech.*").Exist(60) Then
	err
End If

Window("regexpwndtitle:=Nextech.*").WinMenu("menuobjtype:=2").Select "Modules;Surgery Center"
	If Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").Exist(10) Then
		Call RecordResults(True, "Surgery Center module opened!")
	Else
		Call RecordResults(False, "Surgery Center module did not open!***")
	End If

Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinRadioButton("regexpwndtitle:=All Dates").Click
	
	Call ChangeTabs (Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").ActiveX("acx_name:=NxTab Control"), "Case Histories")
	Call VerFTabOpened ("Case Histories")
	Call VerifyCaseHistories()
	wait (1)
	
	Call ChangeTabs (Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").ActiveX("acx_name:=NxTab Control"), "Doctor Prefs.")
	Call VerFTabOpened ("Doctor Prefs.")
	Call VerifyDoctorPrefs()
	wait (1)	

	Call ChangeTabs (Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").ActiveX("acx_name:=NxTab Control"), "Credentials")
	Call VerFTabOpened ("Credentials")
	Call VerifyCredentials()
	wait (1)	

	
Sub VerifyCaseHistories()
	Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinRadioButton("regexpwndtitle:=All Dates").Click
			
End Sub

Sub VerifyDoctorPrefs()
'Clicking on the Add Button
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Add").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").Exist Then
		Reporter.ReportEvent micPass, "Input Dialog box is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
	End If
	
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Type "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

'Clicking on the Rename Button
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Rename").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").Exist Then
		Reporter.ReportEvent micPass, "Input Dialog box is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Type "UFT Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click	
	
'Clicking on the Delete Button
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Delete").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").Static("text:=Are you SURE you wish to remove this preference card\?").Exist Then
		Reporter.ReportEvent micPass, "Practice Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Save Copy Button
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Save Copy").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").Exist Then
		Reporter.ReportEvent micPass, "Input Dialog box is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Type "UFT Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click	

'Clicking on Delete Button again
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Delete").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").Static("text:=Are you SURE you wish to remove this preference card\?").Exist Then
		Reporter.ReportEvent micPass, "Practice Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Link To Providers
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Link To Providers").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Link Providers To Preference Cards").Exist Then
		Reporter.ReportEvent micPass, "Link Providers To Preference Cards Dialog is Present",""
	Else	
		Reporter.ReportEvent micFail, "Test Failed", "Link Providers To Preference Cards Dialog is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Link Providers To Preference Cards").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Advanced Editing
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Advanced Editing").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Advanced Preference Card Editing").Exist Then
		Reporter.ReportEvent micPass, "Advanced Preference Card Editing Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Advanced Preference Card Editing Dialog is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Advanced Preference Card Editing").WinButton("regexpwndtitle:=Apply").Click

str1 = "You must have at least one Inventory Item, Personnel, or Service Code selected\."
	If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Advanced Preference Card Editing").Dialog("regexpwndtitle:=NexTech Practice").Static("text:="&str1).Exist Then
		Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
	End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Advanced Preference Card Editing").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Advanced Preference Card Editing").WinButton("regexpwndtitle:=Close").Click

End Sub

Sub VerifyCredentials()

'Clicking on Configure Licenses
Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Configure Licenses").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").Exist Then
		Reporter.ReportEvent micPass, "License Setup Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "License Setup Dialog is not Present"
	End If

'Clicking on Add New License Button in License Setup Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").WinButton("regexpwndtitle:=Add New License").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").Dialog("regexpwndtitle:=Input").Exist Then
		Reporter.ReportEvent micPass, "Input Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Type "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").WinButton("regexpwndtitle:=Add New License").Click

'Clicking the Delete License Button in License Setup Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").WinButton("regexpwndtitle:=Delete License").Click
	If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").Dialog("regexpwndtitle:=Practice").Exist Then
		Reporter.ReportEvent micPass, "Practice Dialog is Present",""
	Else
		Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
	End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Closing the License Setup Dialog box
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=License Setup").WinButton("regexpwndtitle:=Close").Click
	
End Sub
	
	
	Sub VerFTabOpened(TabName)

Select Case TabName

	Case "Case Histories"
		If Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinRadioButton("regexpwndtitle:=All Dates").Exist(3) Then
			Call RecordResults(True, "Surgery Center Tab opened!")
		Else
			Call RecordResults(False, "<<< Surgery Center tab did not open! >>>")
		End If

	Case "Doctor Prefs."
		If Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Add").Exist(3) Then
			Call RecordResults(True, "Doctor Prefs. tab opened!")			
		Else
			Call RecordResults(False, "<<< Doctor Prefs. tab did not open! >>>")
		End If

	Case "Credentials"
		If Window("regexpwndtitle:=Nextech.*").Window("regexpwndtitle:=Surgery Center").WinButton("regexpwndtitle:=Configure Licenses").Exist(3) Then
			Call RecordResults(True, "Crendentials Batch tab opened!")	
		Else
			Call RecordResults(False, "<<< Credentials tab did not open! >>>")
		End If

		Case Else
		Call RecordResults(False, "Bad call to case statement: Bad name was: "& TabName)	
End Select

End Sub

Window("regexpwndtitle:=Nextech.*").WinMenu("menuobjtype:=2").Select "Modules;Patients"


RunAction "LogOut [GlobalPracticeActions]", oneIteration






