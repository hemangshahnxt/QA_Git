
'On Error Resume Next

Dim logmsg
Dim IsPresent
Dim TheCurrentDateStr
Dim NumPassedInt, NumFailedInt
Dim DBUnderTest
NumPassedInt = 0
NumFailedInt = 0
Dim MsgTextStr
Dim VerificationResult

Dim TheMessage
'On Error Resume Next
'Msgbox Environment.Value("DBUnderTest")

DBUnderTest = Environment.Value("DBUnderTest")
TheCurrentDateStr = CStr(date)

'Log in to practice
'
Set oShell=CreateObject("WScript.Shell")


'Log in to practice

RunAction "LogIn [GlobalPracticeActions]", oneIteration, DBUnderTest

wait (2)

If Not Window("Nextech Main Window").Exist(60) Then
	err
End If

Window("Nextech Main Window").WinToolbar("ModuleButtons").Press 7
If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").Exist(10) Then
	Call RecordResults(True, "Financial module opened!")
Else
	Call RecordResults(False, "Financial module did not open!***")
End If

Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinRadioButton("regexpwndtitle:=All Dates").Click
	
	Call ChangeTabs (Window("Nextech Main Window").Window("regexpwndtitle:=Financial").ActiveX("acx_name:=NxTab Control"), "Banking")
	Call VerFTabOpened ("Banking")
	Call VerifyBanking()
	wait (1)
'	
	Call ChangeTabs (Window("Nextech Main Window").Window("regexpwndtitle:=Financial").ActiveX("acx_name:=NxTab Control"), "Paper Batch")
	Call VerFTabOpened ("Paper Batch")
	Call VerifyPaperBatch()
'	wait (1)	
'
	Call ChangeTabs (Window("Nextech Main Window").Window("regexpwndtitle:=Financial").ActiveX("acx_name:=NxTab Control"), "EBilling Batch")
	Call VerFTabOpened ("EBilling Batch")
	Call VerifyEBillingBatch()
'	wait (1)	

	Call ChangeTabs (Window("Nextech Main Window").Window("regexpwndtitle:=Financial").ActiveX("acx_name:=NxTab Control"), "E-Eligibility")
	Call VerFTabOpened ("E-Eligibility")
	Call VerifyEEligibility()
'	wait (1)	

	Call ChangeTabs (Window("Nextech Main Window").Window("regexpwndtitle:=Financial").ActiveX("acx_name:=NxTab Control"), "Batch Payments")
	Call VerFTabOpened ("Batch Payments")
	Call VerifyBatchPayments()
	wait (1)	

	Call ChangeTabs (Window("Nextech Main Window").Window("regexpwndtitle:=Financial").ActiveX("acx_name:=NxTab Control"), "Billing Followup")
	Call VerFTabOpened ("Billing Followup")
	Call VerifyBillingFollowup()

'Play Area

'
'	Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Send Selected Items\nTo Quickbooks").Click
'	
'	'Click the send to quickbooks button with nothing selected
'	VerificationResult = VerifyPopups("Nextech Practice","There are no payments in the 'Selected' lists\.")
'	If VerificationResult Then
'		Call RecordResults(True,"Expected popup displayed")
'		Window("regexpwndtitle:=Nextech Practice").winbutton("regexpwndtitle:=OK").Click
'	Else
'		Call RecordResults(False,"<<< Expected a popup that didn't display >>>")
'	End If
'
	'Move payments to selected by clicking All Dates and the send to quickbooks button.  different popups expected

'	Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Send Selected Items\nTo Quickbooks").Click
'	
'Window("Nextech Main Window").Dialog("Practice").WinButton("No").Click @@ hightlight id_;_984552_;_script infofile_;_ZIP::ssf8.xml_;_
'Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("No").Click @@ hightlight id_;_461282_;_script infofile_;_ZIP::ssf7.xml_;_
'	VerificationResult = Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").Exist(2)
	
'VerificationResult = Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").Static("text:=.*Quickbooks Link.*").Exist(2)
''Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").Static.GetProperty("text","TheMessage")
'Dim ThehwndValue
'TheMessage = Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").GetROProperty("hwnd")
'MsgBox (TheMessage)


Sub VerifyBanking()
	

Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Send Selected Items\nTo Quickbooks").Click
	
	'Click the send to quickbooks button with nothing selected
a = "The Quickbooks Link files are not installed on this computer\.\n\n" &_
"You can install these files from your NexTech CD or NxCD folder on the server, under the ThirdParty\\QBSetup folder\.\n\n"&_
"Would you like to auto-install the Quickbooks Link files now\?\n\(After installing, you will need to restart both Practice and Quickbooks\.\)"
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Practice").Static("text:="&a).Exist Then
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click
Else
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click
End If


b = "The Quickbooks Link setup files must be installed before the link can be utilized."

If Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").Static("text:="&b).Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
	
'Clicking on the Print Deposit Slip
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Print Deposit Slip").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Print").Exist(5) Then
	Reporter.ReportEvent micPass, "Print Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Print Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Print").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Deposit Selected Items
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Deposit Selected Items").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click

'Clicking on the Prepare Refund Checks
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Prepare Refund Checks").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Prepare Refund Checks dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Prepare Refund Checks dialog is not Present"
End If

'Clicking on Auto-Number Checks inside the Prepare Refund Checks
Window("Nextech Main Window").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").WinButton("regexpwndtitle:=Auto-Number Checks").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Clear Check Numbers inside the Prepare Refund Checks
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").WinButton("regexpwndtitle:=Clear Check Numbers").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Save and Preview Report Button
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").WinButton("regexpwndtitle:=Save and Preview Report").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click

'Clicking on Save and Close Button inside the Prepare Refund Checks
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").WinButton("regexpwndtitle:=Save and Close").Click
If Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click

'Clicking on the Cancel Button inside the Prepare Refund Checks
Window("regexpwndtitle:=Nextech.*","index:=0").Dialog("regexpwndtitle:=Prepare Refund Checks","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Restore Past Deposits
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Restore Past Deposits").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Retrieve Past Deposits").Exist Then
	Reporter.ReportEvent micPass, "Retrieve Past Deposits dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Retrieve Past Deposits dialog is not Present"
End If

Window("Nextech Main Window").Dialog("regexpwndtitle:=Retrieve Past Deposits").WinButton("regexpwndtitle:=Close").Click



End Sub

Sub VerifyPaperBatch()
'Clicking on the Configure Claim Validation in Paper Batch 
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Configure Claim Validation").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Claim Validation Configuration").Exist Then
	Reporter.ReportEvent micPass, "Claim Validation Configuration Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Claim Validation Configuration Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Claim Validation Configuration").WinButton("regexpwndtitle:=Cancel").Click
	
'Clicking on Unbatch Unselected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch\nUnselected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").Static("regexpwndtitle:=Are you sure you want to Unbatch.*").Exist Then
	Reporter.ReportEvent micPass, "Question Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Question Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").WinButton("regexpwndtitle:=&No").Click

'Clicking on Preview Unselected HCFA claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview\nUnselected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("regexpwndtitle:=The are no claims selected\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Validate Unselected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Validate\nUnselected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("regexpwndtitle:=There are no unselected claims to validate!").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is Not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Unbatch All HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch\nAll\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").Static("regexpwndtitle:=Are you sure you want to Unbatch.*").Exist Then
	Reporter.ReportEvent micPass, "Question Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Question Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").WinButton("regexpwndtitle:=&No").Click

'Clicking on Preview All Batched HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview\nAll Batched\nHCFA Claims").Click
If Window("Nextech Main Window").Window("regexpwndtitle:=Paper Claim List \(PP\)").Exist Then
	Reporter.ReportEvent micPass, "Paper Batch List Print Preview is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Pater Batch List Print Preview is not Present"
End If
Window("Nextech Main Window").Window("regexpwndtitle:=Paper Claim List \(PP\)").Close

'Clicking on Validate all HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Validate\nAll\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("regexpwndtitle:=There are no batched claims to validate!").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Unbatch Selected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch\nSelected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").Static("text:=Are you sure you want to Unbatch all selected.*").Exist Then
	Reporter.ReportEvent micPass, "Question Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Question Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").WinButton("regexpwndtitle:=&No").Click

'Clicking on Preview Selected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview\nSelected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=The are no claims selected\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Validate selected HCFA claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Validate\nSelected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=There are no selected claims to validate!").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Print HCFA Batch
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Print\nHCFA Batch").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("regexpwndtitle:=The are no claims selected\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click


End Sub

Sub VerifyEBillingBatch()
	'Clicking on the Configure Claim Validation in Paper Batch 
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Configure Claim Validation").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Claim Validation Configuration").Exist Then
	Reporter.ReportEvent micPass, "Claim Validation Configuration Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Claim Validation Configuration Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Claim Validation Configuration").WinButton("regexpwndtitle:=Cancel").Click
	
'Clicking on Unbatch Unselected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch\nUnselected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").Static("regexpwndtitle:=Are you sure you want to Unbatch.*").Exist Then
	Reporter.ReportEvent micPass, "Question Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Question Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").WinButton("regexpwndtitle:=&No").Click

'Clicking on Preview Unselected HCFA claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview\nUnselected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("regexpwndtitle:=The are no claims selected\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Validate Unselected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Validate\nUnselected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("regexpwndtitle:=There are no unselected claims to validate!").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is Not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Unbatch All HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch\nAll\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").Static("regexpwndtitle:=Are you sure you want to Unbatch.*").Exist Then
	Reporter.ReportEvent micPass, "Question Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Question Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").WinButton("regexpwndtitle:=&No").Click

'Clicking on Preview All Batched HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview\nAll Batched\nHCFA Claims").Click
If Window("Nextech Main Window").Window("regexpwndtitle:=Ebilling Batch List \(PP\)").Exist Then
	Reporter.ReportEvent micPass, "Electronic Billing Claim List Print Preview is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Electronic Billing Claim List Print Preview is not Present"
End If
Window("Nextech Main Window").Window("regexpwndtitle:=Ebilling Batch List \(PP\)").Close

'Clicking on Validate all HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Validate\nAll\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("regexpwndtitle:=There are no batched claims to validate!").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Unbatch Selected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch\nSelected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").Static("text:=Are you sure you want to Unbatch all selected.*").Exist Then
	Reporter.ReportEvent micPass, "Question Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Question Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Question").WinButton("regexpwndtitle:=&No").Click

'Clicking on Preview Selected HCFA Claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview\nSelected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=The are no claims selected\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Validate selected HCFA claims
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Validate\nSelected\nHCFA Claims").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=There are no selected claims to validate!").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click
	
'Clicking on Export Batch
If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Export Batch").Exist Then
'
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Export Batch").Click
str1 = "There are no claims in the selected list\. Please batch some claims before exporting\."
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:="&str1).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Format 997 Report
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Format 997 Report").Click
wait(.5)
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Format ANSI 997 Report File").Exist Then
	Reporter.ReportEvent micPass, "Format ANSI 997 Report file Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Format ANSI 997 Report file Dialog is not Present"
End If

'Clicking on the TWO browse button in the Format ANSI 997
Window("Nextech Main Window").Dialog("regexpwndtitle:=Format ANSI 997 Report File").WinButton("regexpwndtitle:=Browse","index:=0").Click
wait(.5)
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 997 Report File").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Browse for folder dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Browse for Folder dialog is not Open"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 997 Report File").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 997 Report File").WinButton("regexpwndtitle:=Browse","index:=1").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 997 Report File").Dialog("regexpwndtitle:=Save As").Exist Then
	Reporter.ReportEvent micPass, "Save the File dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save the File dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 997 Report File").Dialog("regexpwndtitle:=Save As").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Format report File in the Format ANSI 997 Report
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 997 Report File").WinButton("regexpwndtitle:=Format Report File").Click
str2 = "The input file could not be found or opened\. Please double-check your path and filename\."
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Format ANSI 997 Report File").Dialog("regexpwndtitle:=NexTech Practice").Static("text:="&str2).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Format ANSI 997 Report File").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 997 Report File").WinButton("regexpwndtitle:=Close").Click

'Clciking on Format 277 Report
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Format 277 Report").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").Exist Then
	Reporter.ReportEvent micPass, "Format ANSI 277 Report file Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Format ANSI 277 Report file Dialog is not Present"
End If

'Clicking on the TWO browse button in the Format ANSI 997
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").WinButton("regexpwndtitle:=Browse","index:=0").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Browse for folder dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Browse for Folder dialog is not Open"
End If
' <<<< Practice has stopped working here >>>>
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click
wait(.5)
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").WinButton("regexpwndtitle:=Browse","index:=1").Click
wait(2)
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").Dialog("regexpwndtitle:=Save As").Exist Then
	wait(1)
	Reporter.ReportEvent micPass, "Save the File dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Save the File dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").Dialog("regexpwndtitle:=Save As").WinButton("regexpwndtitle:=Cancel").Click
wait(.5)

'Clicking on Format report File in the Format ANSI 997 Report
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").WinButton("regexpwndtitle:=Format Report File").Click
wait(.5)

str3 = "The input file could not be found or opened\. Please double-check your path and filename\."
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Format ANSI 277 Report File").Dialog("regexpwndtitle:=NexTech Practice").Static("text:="&str3).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Format ANSI 277 Report File").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Format ANSI 277 Report File").WinButton("regexpwndtitle:=Close").Click

'Clicking on Clearinghouse Integration
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Clearinghouse Integration").Click

If Window("Nextech Main Window").Dialog("regexpwndtitle:=E-Billing Clearinghouse Integration Setup").Exist Then
	Reporter.ReportEvent micPass, "Clearinghouse Integration Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Clearinghouse Intergration Dialog is not Present"
End If

Window("Nextech Main Window").Dialog("regexpwndtitle:=E-Billing Clearinghouse Integration Setup").WinButton("regexpwndtitle:=Cancel").Click


'Clicking on ANSI properties
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=ANSI Properties").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Exist Then
	Reporter.ReportEvent micPass, "ANSI properties Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "ANSI properties Dialog is not Present"
End If

'Clicking on Add Button inside the ANSI properties Dialog
Window("Nextech Main Window").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Add").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click


If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Set "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

'Clicking on Delete Button inside the ANSI properties Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Delete").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").Static("text:=Are you sure you wish to permanently delete this vendor\?").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Show Advanced Options
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Show Advanced Options").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Hide Advanced Options").Exist Then
	Reporter.ReportEvent micPass, "ANSI Properties Dialog does extend",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "ANSI Properties Dialog does not extend"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Hide Advanced Options").Click

'Clicking on Cancel Button in the ANSI properties Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Retrieve Past Batches Button
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Retrieve Past Batches").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Retrieve Past Electronic Batches").WinButton("regexpwndtitle:=Close").Exist Then
	Reporter.ReportEvent micPass, "Retrieve Past Electronic Batches Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Retrieve past Electronic Batches Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Retrieve Past Electronic Batches").WinButton("regexpwndtitle:=Close").Click
	
End Sub

Sub VerifyEEligibility()

'Clicking on the Create Single Eligibility Request
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Create Single Eligibility Request").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Eligibility Request").WinButton("regexpwndtitle:=Cancel").Exist Then
	Reporter.ReportEvent micPass, "Eligibility Request Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Eligibility Request Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Eligibility Request").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Create Requests for Scheduled Patients
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Create Requests For Scheduled Patients").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Create E-Eligibility Requests For Scheduled Patients").Exist Then
	Reporter.ReportEvent micPass, "Create E-Eligibility Reqests Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Create E-Eligibility Requests Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Create E-Eligibility Requests For Scheduled Patients").WinButton("regexpwndtitle:=Create Requests").Click
str4 = "There are no insured patients that match your filters and have appointments in the range given and do not already have an E-Eligiblity request batched\."

If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Create E-Eligibility Requests For Scheduled Patients").Dialog("regexpwndtitle:=NexTech Practice").Static("text:="&str4).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Create E-Eligibility Requests For Scheduled Patients").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Create E-Eligibility Requests For Scheduled Patients").WinButton("regexpwndtitle:=Close").Click

'Clicking on Unbatch Unselected Requests 
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch \nUnselected \nRequests").Click
If Window("Nextech Main Window").Dialog("Practice").Static("text:=Are you sure you want to unbatch the unselected Eligibility Requests\?").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("Practice").WinButton("No").Click

'Clicking on Review Past Eiligibility Requests
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Review Past Eligibility Requests").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Eligibility Review").Exist Then
	Reporter.ReportEvent micPass, "Eligibility Review Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Eligibility Review Dialog is not Present"
End If

'Clicking on Re-batch Selected Requests in Eligibility Review Dialog
Window("Nextech Main Window").Dialog("regexpwndtitle:=Eligibility Review").WinButton("regexpwndtitle:=Re-batch Selected Requests").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Eligibility Review").Dialog("regexpwndtitle:=NexTech Practice").Static("text:=The selected requests have been re-batched\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Eligibility Review").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Closing the Eligibility Review Dialog
Window("Nextech Main Window").Dialog("regexpwndtitle:=Eligibility Review").WinButton("regexpwndtitle:=Close").Click

'Clicking on Unbatch All Requests
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch \nAll Requests").Click
If Window("Nextech Main Window").Dialog("Practice").Static("text:=Are you sure you want to unbatch all Eligibility Requests\?").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("Practice").WinButton("No").Click

'Clicking on Unbatch Selected Requests
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Unbatch \nSelected \nRequests").Click
If Window("Nextech Main Window").Dialog("Practice").Static("text:=Are you sure you want to unbatch the selected Eligibility Requests\?").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("Practice").WinButton("No").Click

'Clicking on Export Requests
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Export Requests").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=There are no eligibility requests in the selected list\. Please select some requests before exporting\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on the Import Responses Button
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Import Responses").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Import dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed","Import Dialog is not Open"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Configure Response Filtering
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Configure Response Filtering").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Configure E-Eligibility Response Filtering").Exist Then
	Reporter.ReportEvent micPass, "Configure E-Eligibility Response Filtering Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure E-Eligibility Response Filtering Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Configure E-Eligibility Response Filtering").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Real-Time Setting button
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Real-Time Settings").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=E-Eligibility Real-Time Response Setup").Exist Then
	Reporter.ReportEvent micPass, "E-Eligibility Real-Time Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "E-Eligibility Real-Time Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=E-Eligibility Real-Time Response Setup").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on ANSI properties
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=ANSI Properties").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Exist Then
	Reporter.ReportEvent micPass, "ANSI properties Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "ANSI properties Dialog is not Present"
End If

'Clicking on Add Button inside the ANSI properties Dialog
Window("Nextech Main Window").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Add").Click

'Verifying if the Practice dialog is present
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Set "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

'Clicking on Delete Button inside the ANSI properties Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Delete").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").Static("text:=Are you sure you wish to permanently delete this vendor\?").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Show Advanced Options
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Show Advanced Options").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Hide Advanced Options").Exist Then
	Reporter.ReportEvent micPass, "ANSI Properties Dialog does extend",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "ANSI Properties Dialog does not extend"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Hide Advanced Options").Click

'Clicking on Cancel Button in the ANSI properties Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=ANSI Properties","index:=1").WinButton("regexpwndtitle:=Cancel").Click


End Sub

Sub VerifyBatchPayments()
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Add").Click

'Selecting Medical Batch Payment from WinMenu 
Set bc = CreateObject("Mercury.DeviceReplay")
bc.PressKey 208
bc.PressKey 28
Set bc = Nothing
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Medical Batch Payment").Exist Then
	Reporter.ReportEvent micPass, "Medical Batch Payment Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Medical Batch Payment Dialog is not Present"
End If

'Clicking on the first ellipsis Button
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").WinButton("regexpwndtitle:=\.\.\.", "index:=0").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Exist Then
	Reporter.ReportEvent micPass, "Payment Categories Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Payment Categories Dialog box is not Present"
End If

'Clicking on the ADD button in the Payment Catogeries Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Add").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinEdit("regexpwndclass:=Edit").Set "abc"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

'Closing the Payment Categories Dialog box
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Close").Click

'Clicking on the second ellipsis Button in the Batch Payment Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").WinButton("regexpwndtitle:=\.\.\.", "index:=1").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").Exist Then
	Reporter.ReportEvent micPass, "Edit Combo Box Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Combo Box Dialog is not Present"
End If

'Clicking on the Add Button in the Edit Combo Box Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Add").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Edit").Click

'Clicking on the Delete Button
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Delete").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Reporter.ReportEvent micPass, "Delete Dialog is Present", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Delete Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Close").Click

'Clicking on the Create Payment Button
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").WinButton("regexpwndtitle:=Create Payment").Click
Str1 = "You have not selected a provider for this batch payment\.\nAll new batch payments must have providers selected\."
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Practice").Static("text:="&str1).Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click

'Closing the Medical Batch Payment Dialog box
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Medical Batch Payment").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Add to open the Vision Batch Payment dialog
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Add").Click

'Selecting Vision Batch Payment from WinMenu 
Set bc = CreateObject("Mercury.DeviceReplay")
bc.PressKey 208
bc.PressKey 208
bc.PressKey 28
Set bc = Nothing

If Window("Nextech Main Window").Dialog("regexpwndtitle:=Vision Batch Payment").Exist Then
	Reporter.ReportEvent micPass, "Vision Batch Payment Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Vision Batch Payment Dialog is not Present"
End If

'Clicking on the first ellipsis Button
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").WinButton("regexpwndtitle:=\.\.\.", "index:=0").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Exist Then
	Reporter.ReportEvent micPass, "Payment Categories Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Payment Categories Dialog box is not Present"
End If

'Clicking on the ADD button in the Payment Catogeries Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Add").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinEdit("regexpwndclass:=Edit").Set "abc"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

'Closing the Payment Categories Dialog box
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Close").Click

'Clicking on the second ellipsis Button in the Batch Payment Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").WinButton("regexpwndtitle:=\.\.\.", "index:=1").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").Exist Then
	Reporter.ReportEvent micPass, "Edit Combo Box Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Combo Box Dialog is not Present"
End If

'Clicking on the Add Button in the Edit Combo Box Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Add").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Edit").Click

'Clicking on the Delete Button
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Delete").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Reporter.ReportEvent micPass, "Delete Dialog is Present", ""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Delete Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Close").Click

'Clicking on the Create Payment Button
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").WinButton("regexpwndtitle:=Create Payment").Click
Str1 = "You have not selected a provider for this batch payment\.\nAll new batch payments must have providers selected\."
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Practice").Static("text:="&str1).Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click

'Closing the Vision Batch Payment Dialog box
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Vision Batch Payment").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the Edit Button
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Edit").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=Please select a payment from the list\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click

'Clicking on the Delete Button
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Delete").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=Please select a payment from the list\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click

'Clicking on Import Electronic Payment and E-Remittance Payment File
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Import Electronic Payment").Click
Set bc = CreateObject("Mercury.DeviceReplay")
bc.PressKey 208
bc.PressKey 28
Set bc = Nothing
'Window("regexpwndtitle:=Nextech.*").WinMenu("menuobjtype:=3").Select "E-Remittance Payment File"
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Browser Open Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Browser Open Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Import Electronic Remittance file
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=Import Electronic Remittance File").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Browser Open Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Browser Open Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the first ellipsis button
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=\.\.\.","index:=0").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").Exist Then
	Reporter.ReportEvent micPass, "Edit Combo Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Combo Box is Not Present"
End If

'Clicking on Add Button in the Edit combo box
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Edit").Click

'Clicking on the Delete button in the Edit Combo Box
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Reporter.ReportEvent micPass, "Delete dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Delete Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Close").Click

'Clicking on the second ellipsis Button
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=\.\.\.","index:=1").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").Exist Then
	Reporter.ReportEvent micPass, "Payment Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Payment Categories Dialog is not Present"
End If

'Clicking on the Add Button
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Add").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present with Win edit Field",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Close").Click

'Clicking on Third ellipsis Button
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=\.\.\.","index:=2").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").Exist Then
	Reporter.ReportEvent micPass, "Edit Combo Box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Combo Box is Not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Edit").Click

'Clicking on the Delete button in the Edit Combo Box
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Reporter.ReportEvent micPass, "Delete dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Delete Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Close").Click

'Clicking on the fourth ellipsis Button
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=\.\.\.","index:=3").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").Exist Then
	Reporter.ReportEvent micPass, "Payment Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Payment Categories Dialog is not Present"
End If

'Clicking on the Add Button
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Add").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present with Win edit Field",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Close").Click

'Clicking on Skip All Duplicates
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=Skip All Duplicates").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Practice").Static("text:=You have not imported a remittance file, there are no charges to skip\.").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Print Button
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=Print").Click
str2 = "You have not imported a remittance file that applies to any existing bills\."
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Practice").Static("text:="&str2).Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Process
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=Process").Click
str3 = "You have not imported a remittance file that applies to any existing bills\."
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Practice").Static("text:="&str3).Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on COnfigure Adjustment codes To Ignore
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=Adjustment Code Settings").Click
wait(1)
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Adjustment Code Settings","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Configure E-Remittance Adj Codes to Ignore Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure E-Remittance Adj Codes to Ignore Dialog is not Present"
End If

'Clicking on the Add and Remove Button
'Modified for 12300 which added a new panel to the dialog that has the same add/remove buttons. scrript needed to be modified to recognize both panels

' Adjustment codes to ignore panel
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Adjustment Code Settings","index:=1").WinButton("regexpwndtitle:=Add","index:=0").Click
'oShell.SendKeys "{ESC}"
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Adjustment Code Settings","index:=1").WinButton("regexpwndtitle:=Remove","index:=0").Click
Reporter.ReportEvent micPass, "Adjustment codes to ignore Add worked",""
wait(1)
' Adjustment codes allowing negative adjustments to be posted panel
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Adjustment Code Settings","index:=1").WinButton("regexpwndtitle:=Add","index:=1").Click
'oShell.SendKeys "{ESC}"
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Adjustment Code Settings","index:=1").WinButton("regexpwndtitle:=Remove","index:=1").Click
Reporter.ReportEvent micPass, "Adjustment codes allowing negative adjustments to be posted Add worked",""
wait(1)

'Closing the Configure E-Remittance Adjustment Dialog box
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Adjustment Code Settings","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Configure EOB Import Filters
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=Configure EOB Import Filters").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Configure EOB Import Filters is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure EOB Import Filters is not Present"
End If

'Clicking on Add and Delete Button in Configure EOB Import Filters
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").WinButton("regexpwndtitle:=Add","index:=0").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Set "test"
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").WinObject("object class:=Afx:42000000:8b","index:=0").Click

Set bc = CreateObject("Mercury.DeviceReplay")
bc.PressKey 208
'bc.PressKey 28
Set bc = Nothing

Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").WinButton("regexpwndtitle:=Delete","index:=0").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on the second Add and Delete Button in COnfigure EOB Import Filters
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").WinButton("regexpwndtitle:=Add","index:=1").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Set "test"
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").WinObject("object class:=Afx:42000000:8b","index:=1").Click

Set bc = CreateObject("Mercury.DeviceReplay")
bc.PressKey 208
'bc.PressKey 28
Set bc = Nothing

Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").WinButton("regexpwndtitle:=Delete","index:=1").Click
If Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click

'Closing the Configure EOB Import Filters
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").Dialog("regexpwndtitle:=Configure EOB Import Filters","index:=1").Winbutton("regexpwndtitle:=Close").Click

'Closing the Electronic Remittance/EOD Processing
Dialog("regexpwndtitle:=Electronic Remittance / EOB Processing").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Import Electronic Payment and Lockbox Payment file
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Import Electronic Payment").Click
'Window("regexpwndtitle:=Nextech.*").WinMenu("menuobjtype:=3").Select "Lockbox Payment File"
Set bc = CreateObject("Mercury.DeviceReplay")
bc.PressKey 208
bc.PressKey 208
bc.PressKey 28
Set bc = Nothing
wait(.5)

If Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Choose a file dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Choose a file Dialog is not Open"
End If
wait(.5)
Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click
If Dialog("regexpwndtitle:=Import Lockbox Payment").Exist Then
	Reporter.ReportEvent micPass, "Import Lockbox Payment Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Import Lockbox Payment Dialog is not Present"
End If

'Clicking on Import Lockbox Payment File Button
Dialog("regexpwndtitle:=Import Lockbox Payment").WinButton("regexpwndtitle:=Import Lockbox Payment File").Click
wait(.5)
If Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Open").Exist Then
	Reporter.ReportEvent micPass, "Browse for Folder dialog is Open",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Browse for Folder Dialog is not Open"
End If
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Open").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on the first ellipsis Button
Dialog("regexpwndtitle:=Import Lockbox Payment").WinButton("regexpwndtitle:=\.\.\.","index:=0").Click
If Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Edit Combo Box").Exist Then
	Reporter.ReportEvent micPass, "Edit Combo Box Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Combo Box Dialog is not Present"
End If

'Clicking on the Add Button in Edit Combo Box
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Add").Click
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Edit").Click
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Delete").Click
If Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Reporter.ReportEvent micPass, "Delete dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Delete Dialog is not Present"
End If
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Close").Click

'Clicking on the second ellipsis button
Dialog("regexpwndtitle:=Import Lockbox Payment").WinButton("regexpwndtitle:=\.\.\.","index:=1").Click
If Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Payment Categories").Exist Then
	Reporter.ReportEvent micPass, "Payment Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Payment Categories Dialog is not Present"
End If

'Clicking on the Add Button
wait(.5)
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Add").Click
If Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present with Win edit Field",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
wait(.5)
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click
Dialog("regexpwndtitle:=Import Lockbox Payment").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Close").Click

'Closing the Import Lockbox Payment Dialog
Dialog("regexpwndtitle:=Import Lockbox Payment").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Create Refund
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Create Refund").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").Static("text:=Please select a payment from the list to refund\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
wait(.5)
Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Create Adjustment
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Create Adjustment").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").Static("text:=Please select a payment from the list to adjust\.").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
wait(.5)
Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on View Refunded/Adjusted Payments
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=View Refunded / Adjusted Payments").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Refunded / Adjusted Batch Payments").Exist Then
	Reporter.ReportEvent micPass, "Refunded/Adjusted Batch Payment Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Refunded/Adjusted Batch Payment Dialog is not Present"
End If

'Clicking on Unapply Item Button
wait(.5)
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Refunded / Adjusted Batch Payments").WinButton("regexpwndtitle:=Unapply Item").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Refunded / Adjusted Batch Payments").Dialog("regexpwndtitle:=Nextech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Refunded / Adjusted Batch Payments").Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click

Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Refunded / Adjusted Batch Payments").WinButton("regexpwndtitle:=Close").Click

'Clicking on Manage Lockbox Deposits
wait(.5)
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Manage Lockbox Deposits").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Manage Lockbox Deposits","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Manage Lockbox Deposits Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Manage Lockbox Deposits Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Manage Lockbox Deposits","index:=1").WinButton("regexpwndtitle:=Close").Click

'Clicking on Filter By Bill ID
wait(.5)
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Filter By Bill ID").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Go To Patient ID
wait(.5)
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Go To Patient ID").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&Cancel").Click

'Clicking on Post to Patient
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Post To Patient").Click

'Clicking on Post reversals
wait(.5)
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Post Reversals").Click
Str5 = "This feature should be used when the payment has been reversed by the insurance company\.\r\n\r\nThe following actions will occur:\r\n\r\n- "&_ 
"A reverse payment will be created and applied to the original payment to offset its amount\.\r\n- "&_
"The payment amount will be credited towards the selected batch payment\.\r\n- If a payment is linked with a quote, it will be unlinked\."
If Dialog("regexpwndtitle:=Nextech Practice").Static("text:="&Str5).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If

Dialog("regexpwndtitle:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Insurance Reversal").Exist Then
	Reporter.ReportEvent micPass, "Insurance Reversal Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Insurance Reversal Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("regexpwndtitle:=Insurance Reversal").WinButton("regexpwndtitle:=Cancel").Click

End Sub

Sub VerifyBillingFollowup()

'Clicking on Search Payemnts Under Allowed Amount
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Search Payments\nUnder Allowed\nAmount").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").Exist Then
	Reporter.ReportEvent micPass, "Payments Under Allowed Amount Search Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Payments Under Allowed Amount Search Dialog is not Present"
End If

'Clicking on Search Button in Payemnts Under Allowed Amount
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").WinButton("regexpwndtitle:=Search").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Preview button in Payments Under Allowed Amount
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").WinButton("regexpwndtitle:=Preview").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Closing the Payments Under Allowed Amount Dialog
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Payments Under Allowed Amount Search Utility").WinButton("regexpwndtitle:=Close").Click

'Clicking on Write Off Amounts
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Write Off Accounts").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Exist Then
	Reporter.ReportEvent micPass, "Account Write Off Utility Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Account Write Off Utility Dialog is not Present"
End If
	
'Clicking on Display Results Button in Account Write off Utility
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").WinButton("regexpwndtitle:=Display Results").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").WinButton("regexpwndtitle:=Adjust Accounts To Zero").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&No").Click

'Clicking on the first ellipsis Button in the Account write off Utility
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").WinButton("regexpwndtitle:=\.\.\.","index:=0").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Edit Combo Box").Exist Then
	Reporter.ReportEvent micPass, "Edit Combo Box Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Combo Box Dialog is not Present"
End If

'Clicking on the Add, Edit and Delete Button in Edit Combo box
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Add").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Edit").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Delete").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").Exist Then
	Reporter.ReportEvent micPass, "Delete Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Delete Dialog box is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Edit Combo Box").Dialog("regexpwndtitle:=Delete\?").WinButton("regexpwndtitle:=&Yes").Click
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Edit Combo Box").WinButton("regexpwndtitle:=Close").Click

'Clicking on the second ellipsis Button in the Account Write off Utility
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").WinButton("regexpwndtitle:=\.\.\.","index:=1").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Payment Categories").Exist Then
	Reporter.ReportEvent micPass, "Payment Categories Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Payment Categories Dialog box is not Present"
End If

'Clicking on Add Button in Payment Categories
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Add").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Payment Categories").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

'Closing the Payment Categories Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").Dialog("regexpwndtitle:=Payment Categories").WinButton("regexpwndtitle:=Close").Click

'Closing the Account Write Off Utility Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Account Write Off Utility").WinButton("regexpwndtitle:=Close").Click

'Clicking on Send to Paper Batch
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Send To Paper Batch").Click
str1="As this feature could potentally change the batch of large amounts of bills, \nplease click on 'Display Results' "&_
"on the left panel to preview the claims first\."

If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:="&str1).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Send to Ebilling Batch
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Send To Ebilling Batch").Click
str2="As this feature could potentally change the batch of large amounts of bills, \nplease click on 'Display Results' "&_
"on the left panel to preview the claims first\."

If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:="&str2).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Preview Tracer Forms
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview Tracer\nForms").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=There are no results to display!").Exist Then
	Reporter.ReportEvent micPass, "NexTech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Preview List
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Preview List").Click
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:=There are no results to display!").Exist Then
	Reporter.ReportEvent micPass, "NexTech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Merge Charges To Word
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Merge Charges\nTo Word").Click
str3 = "Please click on 'Display Results' on the left panel to preview the claims before merging\."
If Window("Nextech Main Window").Dialog("NexTech Practice").Static("text:="&str3).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=OK").Click

'Clicking on Display Results
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Display Results").Click

'When clicking display result List is displayed in the field below
'clicking on the notes(paper icon) on the far left side
wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinObject("object class:=Afx:42.*:.*").Click 11,53
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Exist Then
	Reporter.ReportEvent micPass, "Bill Notes Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Bill Notes Dialog is not Present"
End If

'Clicking on Add Macro in Bill Notes Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").WinButton("regexpwndtitle:=A&dd Macro").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Exist Then
	Reporter.ReportEvent micPass, "Add Macro Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Add Macro Dialog is not Present"
End If

'Clicking on Edit Macros Button in Add Macro Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").WinButton("regexpwndtitle:=Edit Macros").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Edit Macros Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Macros Dialog is not Present"
End If

'Clicking on Add macros button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Add Macro").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Set "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

'Clicking on Delete macros Button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Delete Macro").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed","Nextech Practice Dialog is not Present"
End If

Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Closing the Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Closing the Add Macro Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Add Macro").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Macros on Bill Notes Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").WinButton("regexpwndtitle:=Edit Macros").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Edit Macro Dialog box is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Edit Macro Dialog is not Present"
End If

'Clicking on Add macros button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Add Macro").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Set "Test"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click

'Clicking on Delete macros Button in Edit Macros Dialog
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Delete Macro").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed","Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Closing the Edit Macros Dialog box
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Edit Macros","index:=1").WinButton("regexpwndtitle:=Cancel").Click

'Clicking on Edit Categories Button
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").WinButton("regexpwndtitle:=Edit Cate&gories").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Exist Then
	Reporter.ReportEvent micPass, "Note/ Follow-Up Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Note/ Follow-Up Categories Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").WinButton("regexpwndtitle:=&New").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").Exist Then
	Reporter.ReportEvent micPass, "Input Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Input Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").Set "ABCD"
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note / Follow-Up Categories").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").ActiveX("acx_name:=NexTech DataList Control").WinObject("object class:=Afx:.*:.*").Click 11,33

'Clicking on the Delete Button in Note/Follow Up Categories
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=&Delete").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=NexTech Practice").Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click

'Clicking on Adv. EMR Merge...
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=&Adv\. EMR Merge\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:= Default Categories").Exist Then
	Reporter.ReportEvent micPass, "Default Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Default Categories Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:= Default Categories").WinButton("regexpwndtitle:=Cancel","index:=1").Click

'Clicking on Combine Categories...
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=&Combine Categories\.\.\.").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=Combine Note Categories").Exist Then
	Reporter.ReportEvent micPass, "Combine Note Categories Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Combine Note Categories Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").Dialog("regexpwndtitle:=Combine Note Categories").WinButton("regexpwndtitle:=Close").Click

'Closing the Note/Follow-Up Categories
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Note /.*").WinButton("regexpwndtitle:=Clos&e").Click

'Clicking on Search Notes
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Bill Notes").WinButton("regexpwndtitle:=Search Notes").Click
str4 = "You are currently filtering the notes on financial information\. However, the search will look through all the notes for this patient\."
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=NexTech Practice").Static("text:="&str4).Exist Then
	Reporter.ReportEvent micPass, "Nextech Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Nextech Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=OK").Click
If Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Search Notes For:.*").Exist Then
	Reporter.ReportEvent micPass, "Search Notes for Patient Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Search Notes for Patient Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Dialog("regexpwndtitle:=Search Notes For:.*").WinButton("regexpwndtitle:=Close").Click

'Closing the Bill Notes Dialog 
Window("regexpwndtitle:=Nextech.*","index:=1").Dialog("regexpwndtitle:=Bill Notes").Close

'Clicking on Configure Columns
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Configure Columns").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Configure Columns","index:=1").Exist Then
	Reporter.ReportEvent micPass, "Configure Columns Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Configure Columns Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Configure Columns","index:=1").WinButton("regexpwndtitle:=OK").Click

'Clicking on Create Merge Group
Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Create Merge Group").Click
If Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").Exist Then
	Reporter.ReportEvent micPass, "Practice Dialog is Present",""
Else
	Reporter.ReportEvent micFail, "Test Failed", "Practice Dialog is not Present"
End If
Window("regexpwndtitle:=Nextech.*").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=Cancel").Click

	
End Sub



Sub VerFTabOpened(TabName)

Select Case TabName

	Case "Banking"
		If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Send Selected Items\nTo Quickbooks").Exist(3) Then
			Call RecordResults(True, "Banking Tab opened!")
		Else
			Call RecordResults(False, "<<< Banking tab did not open! >>>")
		End If

	Case "Paper Batch"
		If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Configure Claim Validation").Exist(3) Then
			Call RecordResults(True, "Paper Batch tab opened!")			
		Else
			Call RecordResults(False, "<<< Paper Batch tab did not open! >>>")
		End If

	Case "EBilling Batch"
		If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=ANSI Properties").Exist(3) Then
			Call RecordResults(True, "EBilling Batch tab opened!")	
		Else
			Call RecordResults(False, "<<< EBilling Batch tab did not open! >>>")
		End If

	Case "E-Eligibility"
		If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Create Single Eligibility Request").Exist(3) Then
			Call RecordResults(True, "E-Eligibility tab opened!")	
		Else
			Call RecordResults(False, "<<< E-Eligibility tab did not open! >>>")
		End If

	Case "Batch Payments"
		If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Import Electronic Payment").Exist(3) Then
			Call RecordResults(True, "Batch Payments tab opened!")	
		Else
			Call RecordResults(False, "<<< Batch Payments tab did not open! >>>")
		End If

	Case "Billing Followup"
		If Window("Nextech Main Window").Window("regexpwndtitle:=Financial").WinButton("regexpwndtitle:=Search Payments\nUnder Allowed\nAmount").Exist(3) Then
			Call RecordResults(True, "Billing Followup tab opened!")
		Else
			Call RecordResults(False, "<<< Billing Followup tab did not open! >>>")
		End If

	Case Else
		Call RecordResults(False, "Bad call to case statement: Bad name was: "& TabName)	
End Select

End Sub

Window("regexpwndtitle:=Nextech.*").WinMenu("menuobjtype:=2").Select "Modules;Patients"

RunAction "LogOut [GlobalPracticeActions]", oneiteration










