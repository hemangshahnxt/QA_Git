'On Error Resume Next

Dim logmsg
Dim IsPresent
Dim TheCurrentDateStr
Dim NumPassedInt, NumFailedInt
Dim DBUnderTest
NumPassedInt = 0
NumFailedInt = 0
Dim MsgTextStr


DBUnderTest = "UFT12200DB"
TheCurrentDateStr = CStr(date)

'Log in to practice Dialog

RunAction "LogIn [GlobalPracticeActions]", oneIteration, DBUnderTest


'Verifying the Medication Tab on Patients Module
Call ChangeTabs (Window("Nextech (UFT12200DB)").Window("Patients").ActiveX("acx_name:=NxTab Control"), "Medications")
Call VerTabOpened ("Medications")
Call VerifyMedications()
'--------------UnComment these codes when the test is complete

'WIndow("Nextech (UFT12200DB)").Window("Patients").WinObject("regexpwndtitle:=NxTabView").Highlight



Sub VerifyMedications()
	
'Clicking on the first ellipsis Button in the Medication tab
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx License Configuration").Exist Then
	Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx License Configuration").WinButton("regexpwndtitle:=&No").Click
End If

	
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=\.\.\.", "index:=0").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").Exist Then
	Reporter.ReportEvent micPass, "Edit Medication List is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Medication List is not displayed"
End If

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinMenu("menuobjtype:=3").Select "Import Medication"

'Verifying the Meidcation Import Dialog
Dialog("regexpwndtitle:=Medication Import").WinEdit("regexpwndclass:=Edit").Type "alv"
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Search").Click
wait 2
Dialog("regexpwndtitle:=Medication Import").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Help Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=&Help").Click
Window("NexTech Practice User").Close

'Clicking on Inactive Medications button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Inactive Medications").Click
Dialog("regexpwndtitle:=Inactive Medications").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Latin Notation button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Edit Latin Notation").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Delete").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Close").Click

'Clicking on the Interactions Button
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Interactions").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Drug Interactions.*").Exist Then
	Reporter.ReportEvent micPass, "Interactions Dialog is open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Interaction Dialog is Missing"	
End If

'Clicking on Configure Severity Filters
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Drug Interactions.*").WinButton("regexpwndtitle:=Configure Severity Filters").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Drug Interactions.*").Dialog("regexpwndtitle:=Configure Drug Interaction Severity Filters").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Drug Interactions.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on Med History Button
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Med History").Click
If Window("Nextech (UFT12200DB)").Dialog("text:=Medication History.*").Exist Then
	Reporter.ReportEvent micPass, "Medication History Dialog appears",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Medication History Dialog missing"	
End If
Window("Nextech (UFT12200DB)").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=Request Med History").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on View Eligibility Details
Window("Nextech (UFT12200DB)").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=View Eligibility Details").Click
If Dialog("regexpwndtitle:=Eligibility Details").Exist Then
	Reporter.ReportEvent micPass, "Eligibility Details Dialog displays",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Eligibility Details Dialog not displayed"	
End If
Dialog("regexpwndtitle:=Eligibility Details").WinButton("regexpwndtitle:=Close").Click

'Clicking on Import into Current Meds
Window("Nextech (UFT12200DB)").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=Import into Current Meds").Click
Dialog("regexpwndtitle:=No Historical Medications Checked").WinButton("regexpwndtitle:=OK").Click
Window("Nextech (UFT12200DB)").Dialog("text:=Medication History.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on the ellipsis button on the bottom
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=\.\.\.", "index:=1").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").Exist Then
	Reporter.ReportEvent micPass, "Edit Allergy List Dialog displays",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Allergy List Dialog is not displayed"	
End If
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinMenu("menuobjtype:=3").Select "Import Allergy"

'Try to search the searchbar with just 2 letters
Dialog("regexpwndtitle:=Allergy Import").WinEdit("regexpwndclass:=Edit").Type "Po"
Dialog("regexpwndtitle:=Allergy Import").WinButton("regexpwndtitle:=Search").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "This Test is passes",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Needs Attention!!!!"
End If
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=Allergy Import").WinEdit("regexpwndclass:=Edit").Set ""

Dialog("regexpwndtitle:=Allergy Import").WinEdit("regexpwndclass:=Edit").Type "Poll"
Dialog("regexpwndtitle:=Allergy Import").WinButton("regexpwndtitle:=Search").Click
Dialog("regexpwndtitle:=Allergy Import").WinButton("regexpwndtitle:=Cancel").Click

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinMenu("menuobjtype:=3").Select "Add Free Text Allergy"
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&No").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Edit").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Delete").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Update All Allergies").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Allergy List").WinButton("regexpwndtitle:=Close").Click

'Clicking on ellipsis Button in the middle of the screen

Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=\.\.\.", "index:=2").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").Exist Then
	Reporter.ReportEvent micPass, "Edit Medication List is displayed", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Medication List is not displayed"
End If

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Add").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinMenu("menuobjtype:=3").Select "Import Medication"

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
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=&Help").Click
Window("NexTech Practice User").Close

'Clicking on Inactive Medications button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Inactive Medications").Click
Dialog("regexpwndtitle:=Inactive Medications").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Latin Notation button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Edit Latin Notation").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Delete").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Close").Click

'Clicking on Write From Quick List

Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Write From Quick List").Click
Window("Nextech (UFT12200DB)").Window("Patients").ActiveX("NexTech DataList Control").WinObject("Afx:11580000:8b").Click


'Cliking on the Add Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=0").Click
Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=0").Type "Cipro 250 mg tablet"
Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=0").Click
Dialog("regexpwndtitle:=Add Quick List Item").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Button

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Edit").Click
If Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Exist Then
	Reporter.ReportEvent micPass, "Edit Dialog Box displayed",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Dialog Box not displayed"
End If
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click


'Clicking on Delete Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
End If

'Clicking on Import Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Import\.\.\.").Click
If Dialog("regexpwndtitle:=Multiple Selection").Exist Then
	Reporter.ReportEvent micPass, "Multiple Selection Dialog displayed",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Multiple Selection Dialog not displayed"
End If
Dialog("regexpwndtitle:=Multiple Selection").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Close").Click

'Clicking on Rx Print Setup
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Rx Print\nSetup").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescription Template Setup").Exist Then
	Reporter.ReportEvent micPass, "Prescription Template Dialog Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Prescription Template Dialog Not Present"
End If

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Add New Default Template").Click
Dialog("regexpwndtitle:=Select a New Prescription Template").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Remove Selected Default Template").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Favorite Pharmacies
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Edit Favorite\nPharmacies").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Favorite Pharmacies.*").Exist Then
	Reporter.ReportEvent micPass, "Edit Favorite Pharmacies Dialog Box is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Favorite Pharmacies Dialog Box Not Found"
End If
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Edit Favorite Pharmacies.*").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Word Template
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Edit Word\nTemplate").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Select a template to edit").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Rx Needing Attention Button
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Rx Needing\nAttention").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx License Configuration").Exist Then
	Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx License Configuration").WinButton("regexpwndtitle:=&No").Click
End If

If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Exist Then
	Reporter.ReportEvent micPass, "Prescriptions Needing Attention Dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Prescriptions Needing Attention Dialog is Not Open"
End If

'Clicking on Rx Print Setup in Prescription Needing Attention Dialog
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").WinButton("regexpwndtitle:=Rx Print\nSetup").Click

If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").Exist Then
	Reporter.ReportEvent micPass, "Prescription Template Dialog Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Prescription Template Dialog Not Present"
End If

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Add New Default Template").Click
Dialog("regexpwndtitle:=Select a New Prescription Template").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Remove Selected Default Template").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Prescription Template Setup").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Word Template in the Prescription Needing Attention Dialog
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").WinButton("regexpwndtitle:=Edit Word\nTemplate").Click
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Dialog("regexpwndtitle:=Select a template to edit").WinButton("regexpwndtitle:=Cancel").Click

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Prescriptions Needing Attention").Close

'Clicking on View Formulary Information
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=View Formulary Information").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Add Button in Medication Schedule
Window("Nextech (UFT12200DB)").Window("Patients").WinButton("regexpwndtitle:=Add").Click
If Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").Exist Then
	Reporter.ReportEvent micPass, "Medication Schedule Setup Dialog Opens",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Medication Schedule Setup Dialog did not Open"
End If

'Clicking on the Add Details Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Add Detail").Click
Dialog("regexpwndtitle:=Medication Schedule Detail").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=Medication Schedule Detail").WinButton("regexpwndtitle:=Write Prescription").Click
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
Window("NexTech Practice User").Close

'Clicking on Inactive Medications button
Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Inactive Medications").Click
Dialog("regexpwndtitle:=Inactive Medications").WinButton("regexpwndtitle:=Close").Click

'Clicking on Edit Latin Notation button
Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Edit Latin Notation").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Delete").Click
Dialog("regexpwndtitle:=Latin Prescription Settings").WinButton("regexpwndtitle:=Cancel").Click
Dialog("regexpwndtitle:=Edit Medication List").WinButton("regexpwndtitle:=Close").Click
Dialog("regexpwndtitle:=Medication Schedule Detail").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Buttong in Medication Schedule Setup
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Edit Detail").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Remove Detail Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Remove Detail").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Load Form Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Load From").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Save To Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Save To").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Save New Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Save New").Click
If Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is Not Open"
End If

Dialog("regexpwndtitle:=Practice").WinEdit("regexpwndclass:=Edit").Type "Test"
Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Delete Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Delete","index:=0").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on the Preview Button
Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Preview").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=Medication Schedule Setup").WinButton("regexpwndtitle:=Cancel").Click

If Dialog("regexpwndtitle:=Practice").Exist Then
	Dialog("regexpwndtitle:=Practice").Winbutton("regexpwndtitle:=&Yes").Click
End If

End Sub








'
'Window("Nextech (UFT12200DB)").Dialog("regexpwndtitle:=NexERx User Quick List Configuration").WinButton("regexpwndtitle:=Add").Click
'Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=0").Click
'Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=0").Set "Cipro 250 mg tablet"
'Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=0").Click 94, 86
'
'Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=1").Click
'wait (2)
'Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=1").Type"2"
'Dialog("regexpwndtitle:=Add Quick List Item").WinEdit("regexpwndclass:=Edit", "index:=2").Set "20"
'Dialog("regexpwndtitle:=Add Quick List Item").WinButton("regexpwndtitle:=OK").Click
'
'
'
'
'
'
''Dialog("regexpwndtitle:=Add Quick List Item").ActiveX("acx_name:=NexTech DataList Control 2\.0", "index:=1").DblClick 84,96




