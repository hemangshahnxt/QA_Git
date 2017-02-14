'On Error Resume Next
' Log in to practice

''========================

''currently UFT opens practice attached to the USA63DB.  See Record>Record and Run Settings>Windows Applications
''SystemUtil.Run "C:\PracStation\Practice.exe"
Dim TestDBName : TestDBName = Environment.Value("DBUnderTest")
Dim SUMethod : SUMethod = Environment.Value("StartUpMethod")
''=================
''Dim VCTests
'Set VCTests = CreateObject("QuickTest.Application")
'App.Options.Run.RunMode = "Normal"
'App.Options.Run.StepExecutionDelay = 0
'
'Set App = Nothing 
'This is just a test

' UFTODO
'	Use checkpoints for the verification results as well
'   Clean up what should be functions/subs
'   Create a nightly run test suite
'	Determine best way to handle expected results, i.e. spreadsheet row or import to data table (probably the latter)
'	Need ability to switch patients, test cases, get the values from the balances section
'	Which to use, .txt, .csv, import into data table?
'   Added ***debug*** code, commented out.  Search for that token '***debug***' to see what needs to be done to invoke 'debug'  It will allow faster debugging
'   Add code to read in the ledger once per cell ledger, not once per cell verify *should be a big perf boost


Dim ActRowBillIDStr, ExpRowBillIDStr, ActDateStr, ExpDateStr, ActInputDateStr, ExpInputDateStr, InputDateSubStr
Dim ActOnHoldStr, ExpOnHoldStr, ActDescriptionStr, ExpDescriptionStr, ActChargeStr, ExpChargeStr
Dim ActTotAmountStr, ExpTotAmountStr, ActPaymentStr, ExpPaymentStr, ActAdjustmentStr, ExpAdjustmentStr
Dim ActRefundStr, ExpRefundStr, ActBalanceStr, ExpBalanceStr
Dim xlBookStr, xlAppStr, xlSheetStr, TestCaseNumStr, PilotFyleStr, TestDirStr, StepNoStr, OutputStr
Dim RowInx, ColumnInx, NumPassedInt, NumFailedInt, StepNumInt
Dim EndOfTestBool
Dim PatientTabNameStr
Dim BalBalanceDueStr, BalPrePaymentsStr, BalTotalBalancesStr, BalChargesStr, BalAdjustmentsStr
Dim BalNetChargesStr, BalPaymentsStr, BalRefundsStr, BalNetPaymentsStr
Dim TheCurrentDateStr
Dim chk_PassFail


TheCurrentDateStr = CStr(date)
	Print "Only columns that are validated"
	Print Time
	

NumPassedInt = 0
NumFailedInt = 0


RunAction "LogIn [GlobalPracticeActions]", oneIteration, TestDBName

'Window("Nextech Main Window").Window("Patients").WinEdit("Edit").Set
'UFTODO: Here is an example of something that needs to be fixed


'Window("regexpwndtitle:=.*UFT.....DB.*").Window("regexpwndtitle:=Patients","text:=Patients").WinEdit("regexpwndtitle:=Patient32").Highlight


'Window("regexpwndtitle:=Nextech \(.*\)").Window("regexpwndtitle:=Patients").WinEdit("nativeclass:=Edit","regexpwndtitle:=Patient32").Highlight
'Window("regexpwndtitle:=Nextech \(.*\)").Window("regexpwndtitle:=Patients","index:=0").Highlight


''===========
'Window("regexpwndtitle:=Nextech \(.*\)").Window("regexpwndtitle:=Patients","nativeclass:=Afx:00400000:b:00010003:00000006:.*").Highlight
'
'	If Window("regexpwndtitle:=.*UFT.....DB.*").Window("regexpwndtitle:=Patients").WinEdit("nativeclass:=Edit","regexpwndtitle:=Patient32").Exist Then
'		msgbox ("found")
'	End If
''===============
'
' StepNoStr is used to identify which tab in the spreadsheet to pull the verification data from
StepNumInt = 1
StepNoStr = "Step" + CStr(StepNumInt)

EndOfTestBool = False

While Not EndOfTestBool
			
	
	' This is the spreadsheet that contains the expected results.  It may prove easier to have a comma delimited text file.  TBD
	Set xlAppStr = CreateObject("Excel.Application")
	Set xlBookStr = xlAppStr.Workbooks.Open("C:\code\trunk\QA\UFT\Solutions\VoidAndCorrectComplexBill01\InputFiles\12000_US63_05_32A.xlsx")
	Set xlSheetStr = xlBookStr.WorkSheets("Step3")
	
	' Verify it is the correct patient.  Probably want to do this from the patient dropdown.  Patient is VAC32, Patient32 (43)
	'Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "General 1")
	Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "General 1")
	
	'*****
	'Window("Nextech Main Window").Window("Patients").WinEdit("First, Middle, Last").Check CheckPoint("CP_Gen1FName") @@ hightlight id_;_659186_;_script infofile_;_ZIP::ssf7.xml_;_
''=================	
'	If Window("regexpwindtitle:=.*UFT.....DB.*").Window("regexpwndtitle:=Patients").WinEdit("regexpwndtitle:=Patient32").exist(3) Then
'		msgbox ("found")
'	End If
''==================	
'	'Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("NxTab Control"), "Billing")
	
	'RunAction "NavigateTabs", oneIteration, "Billing"
	
	Call ChangeTabs (Window("Nextech Main Window").Window("Patients").ActiveX("PatientTabsControl"), "Billing")
	
	Set ledger = Window("Nextech Main Window").Window("Patients").ActiveX("NxDL_BillingLedger").Object
	
	int rowCount
	int myRow
	
	rowCount = ledger.GetRowCount()	
	set row = ledger.GetFirstRow()
	myRow = rowCount
		
	
	' Start of ledger verify loop.  The loop needs to consist of:
	'	1) Collapse Ledger (assumed)*
	'	2) Select hide voided items
	'	3) Verify ledger
	'	4) Expand ledger
	'	5) Verify ledger
	'	6) Unselect hide voided items
	'	7) Verify ledger
	'	8) Collapse ledger
	'	9) Verify ledger
	'
	' * Assumption is that the ledger always is collapsed when the ledger is first accessed.
	
		
	'	1) Collapse Ledger (assumed it is collapsed)
	
	'	2) Select hide voided items
	
		Call SelectHideVoidedItems("Select") @@ hightlight id_;_986752_;_script infofile_;_ZIP::ssf9.xml_;_
		'***debug*** comment out the code between the '======='
'=============================		
	'	3) Verify ledger
		Set xlSheetStr = xlBookStr.WorkSheets("Step3")
		Call VerifyLedger()
		
	'	4) Expand ledger
		Call ExpandLedger()
		
	'	5) Verify ledger
		Set xlSheetStr = xlBookStr.WorkSheets("Step4")
		Call VerifyLedger()
		
	'	6) Unselect hide voided items
		Call SelectHideVoidedItems("Unselect")
	
	'	7) Verify ledger
		Set xlSheetStr = xlBookStr.WorkSheets("Step5")
		Call VerifyLedger()
		
	'	8) Collapse ledger
		Call CollapseLedger()
	
	'	9) Verify ledger
		Set xlSheetStr = xlBookStr.WorkSheets("Step6")
		Call VerifyLedger()
	
		Call GetTotals()
			
		Call ValidateTotals("$0.00", "Balance Due", BalBalanceDueStr)
		Call ValidateTotals("$0.00", "PrePayment", BalPrePaymentsStr)
		Call ValidateTotals("$0.00", "Total Balance", BalTotalBalancesStr)
		Call ValidateTotals("$1,200.00", "Charges", BalChargesStr)
		Call ValidateTotals("$0.00", "Adjustments", BalAdjustmentsStr)
		Call ValidateTotals("$1,200.00", "Net Charges", BalNetChargesStr)
		Call ValidateTotals("$1,200.00", "Payments", BalPaymentsStr)
		Call ValidateTotals("$0.00", "Refunds", BalRefundsStr)
		Call ValidateTotals("$1,200.00", "Net Payments", BalNetPaymentsStr)
'==========================================	'***debug***
	
		Call VoidBill()
		Call VoidAndCorrectBill()
		
			
		'=========================================================
		'==================  Void and Correct  ===================
		'=========================================================
		
		
		
		
		
		EndOfTestBool = True
	Wend
	
' close Practice here **

'close the spreadsheet  Add datapool code here **
xlBookStr.Close
xlAppStr.Quit
Set xlSheetStr = Nothing
Set xlBookStr = Nothing
Set xlAppStr = Nothing

	Print Time
	

RunAction "LogOut [GlobalPracticeActions]", oneIteration
'
'msgbox ("break here")
'err  '***debug***
'App.Options.Run.StepExecutionDelay = 100
'Set App = Nothing
'
'================================================================================
'====================== Local Functions and Subs ================================
'================================================================================

Sub VoidBill()

	Call SelectHideVoidedItems("Select")
	Call Void("Bill")
	
'	10) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step7")
	Call VerifyLedger()

'	11) Expand ledger
	Call ExpandLedger()
	
'	12) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step8")
	Call VerifyLedger()
	
'	13) Unselect hide voided items
	Call SelectHideVoidedItems("Unselect")

'	14) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step9")
	Call VerifyLedger()
	
'	Expand the ledger
	Call ExpandLedger()
	Set xlSheetStr = xlBookStr.WorkSheets("Step10")
	Call VerifyLedger()

'	16) Verify ledger


	Call GetTotals()
	wait (1)
	Call ValidateTotals("($455.00)", "Balance Due", BalBalanceDueStr) '????
'	Call ValidateTotals("$455.00", "Balance Due", BalBalanceDueStr)
	Call ValidateTotals("$745.00", "PrePayment", BalPrePaymentsStr)
	Call ValidateTotals("($1,200.00)", "Total Balance", BalTotalBalancesStr)
	Call ValidateTotals("$0.00", "Charges", BalChargesStr)
	Call ValidateTotals("$0.00", "Adjustments", BalAdjustmentsStr)
	Call ValidateTotals("$0.00", "Net Charges", BalNetChargesStr)
	Call ValidateTotals("$1,200.00", "Payments", BalPaymentsStr)
	Call ValidateTotals("$0.00", "Refunds", BalRefundsStr)
	Call ValidateTotals("$1,200.00", "Net Payments", BalNetPaymentsStr)

'	1.6) Undo Correction	
	Call UndoCorrection("VBill")
	
'	17) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step3")
	Call VerifyLedger()

'	18) Expand ledger
	Call ExpandLedger()
	
'	19) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step4")
	Call VerifyLedger()
	
'	20) Unselect hide voided items
	Call SelectHideVoidedItems("Unselect")

'	21) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step5")
	Call VerifyLedger()
	
'	22) Collapse ledger
	Call CollapseLedger()

'	23) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step6")
	Call VerifyLedger()

	Call GetTotals()
	Call ValidateTotals("$0.00", "Balance Due", BalBalanceDueStr)
	Call ValidateTotals("$0.00", "PrePayment", BalPrePaymentsStr)
	Call ValidateTotals("$0.00", "Total Balance", BalTotalBalancesStr)
	Call ValidateTotals("$1,200.00", "Charges", BalChargesStr)
	Call ValidateTotals("$0.00", "Adjustments", BalAdjustmentsStr)
	Call ValidateTotals("$1,200.00", "Net Charges", BalNetChargesStr)
	Call ValidateTotals("$1,200.00", "Payments", BalPaymentsStr)
	Call ValidateTotals("$0.00", "Refunds", BalRefundsStr)
	Call ValidateTotals("$1,200.00", "Net Payments", BalNetPaymentsStr)
End Sub

Sub VoidAndCorrectBill()

	Call SelectHideVoidedItems("Select") @@ hightlight id_;_986752_;_script infofile_;_ZIP::ssf9.xml_;_
	Call VoidAndCorrect ("Bill")
	
'	10) Verify ledger collapsed
	Set xlSheetStr = xlBookStr.WorkSheets("Step12")
	Call VerifyLedger() ' Excel sheet tab Step110
	
'	11) Expand ledger
	Call ExpandLedger()
	
'	12) Verify ledger expanded
	Set xlSheetStr = xlBookStr.WorkSheets("Step13")
	Call VerifyLedger()
	
'	13) Unselect hide voided items
	Call SelectHideVoidedItems("Unselect")
	
	

'	14) Expand the ledger before verifying

	Call ExpandLedger()
	Call ExpandLedger()

	Set xlSheetStr = xlBookStr.WorkSheets("Step14")
	Call VerifyLedger()
	
'	Expand the ledger
	Call CollapseLedger()
	Set xlSheetStr = xlBookStr.WorkSheets("Step15")
	Call VerifyLedger()

'	16) Verify ledger


	Call GetTotals()
	
	Call ValidateTotals("$0.00", "Balance Due", BalBalanceDueStr)
	Call ValidateTotals("$0.00", "PrePayment", BalPrePaymentsStr)
	Call ValidateTotals("$0.00", "Total Balance", BalTotalBalancesStr)
	Call ValidateTotals("$1,200.00", "Charges", BalChargesStr)
	Call ValidateTotals("$0.00", "Adjustments", BalAdjustmentsStr)
	Call ValidateTotals("$1,200.00", "Net Charges", BalNetChargesStr)
	Call ValidateTotals("$1,200.00", "Payments", BalPaymentsStr)
	Call ValidateTotals("$0.00", "Refunds", BalRefundsStr)
	Call ValidateTotals("$1,200.00", "Net Payments", BalNetPaymentsStr)

'	1.6) Undo Correction	
	Call UndoCorrection("VCBill")
	
'	17) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step3")
	Call VerifyLedger()

'	18) Expand ledger
	Call ExpandLedger()
	
'	19) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step4")
	Call VerifyLedger()
	
'	20) Unselect hide voided items
	Call SelectHideVoidedItems("Unselect")

'	21) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step5")
	Call VerifyLedger()
	
'	22) Collapse ledger
	Call CollapseLedger()

'	23) Verify ledger
	Set xlSheetStr = xlBookStr.WorkSheets("Step6")
	Call VerifyLedger()

	Call GetTotals()
	Call ValidateTotals("$0.00", "Balance Due", BalBalanceDueStr)
	Call ValidateTotals("$0.00", "PrePayment", BalPrePaymentsStr)
	Call ValidateTotals("$0.00", "Total Balance", BalTotalBalancesStr)
	Call ValidateTotals("$1,200.00", "Charges", BalChargesStr)
	Call ValidateTotals("$0.00", "Adjustments", BalAdjustmentsStr)
	Call ValidateTotals("$1,200.00", "Net Charges", BalNetChargesStr)
	Call ValidateTotals("$1,200.00", "Payments", BalPaymentsStr)
	Call ValidateTotals("$0.00", "Refunds", BalRefundsStr)
	Call ValidateTotals("$1,200.00", "Net Payments", BalNetPaymentsStr)	
	
'=========================================================	
	
End Sub

Sub SelectHideVoidedItems(SelUnselStr)
	
	If SelUnselStr = "Select" Then
		Window("Nextech Main Window").Window("Patients").WinCheckBox("Hide Voided Items").Set "ON"
		Window("Nextech Main Window").Window("Patients").WinCheckBox("Hide Voided Items").Check CheckPoint("Hide Voided Items Selected")
	ElseIf SelUnselStr = "Unselect" Then
		Window("Nextech Main Window").Window("Patients").WinCheckBox("Hide Voided Items").Set "OFF" @@ hightlight id_;_1250178_;_script infofile_;_ZIP::ssf10.xml_;_
		Window("Nextech Main Window").Window("Patients").WinCheckBox("Hide Voided Items").Check CheckPoint("Hide Voided Items Unselected") @@ hightlight id_;_1250178_;_script infofile_;_ZIP::ssf11.xml_;_
	Else
		MsgBox ("Bad Value Passed to select hide voided items")
		Err
	End If

End Sub
'
Function CellValueToStr(Value)
    If IsNull(Value) Then
        CellValueToStr = ""
    Else
        CellValueToStr = CStr(Value)
    End If
End Function
'


Function GetColumnIndex(dl, name)
	dim c 
	c = dl.ColumnCount
	Print CStr(c)
	For i = 0 To c-1
		dim col
		set col = dl.GetColumn(i)
		If (col.ColumnTitle = name) Then
			GetColumnIndex = i
 		Print "Column Used"
			Print name
			Print CStr(i)

			exit function 
		End If
	Next
	GetColumnIndex = -1
	
End Function

Sub VerifyLedger()
	rowCount = ledger.GetRowCount()	
	set row = ledger.GetFirstRow()
	myRow = rowCount
	
	For RowInx = 2 To rowCount + 1 'Ignore the header row in the spreadsheet
'	If rowCount < 4 Then ***debug*** remove the comment in comment 1and make this a statement
		'Print CStr(RowInx)	
		Call GetLedgerValues()
		Call GetExpectedValues()
		Call ValLedgerRow()
		Set row = row.GetNextRow()	
'	End If ***debug*** remove the comment in column 1 and make this a statement
	Next
End Sub

Sub GetExpectedValues()
		'Get the expected results for that row in the spreadsheet
	ExpRowBillIDStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(1).Value)
	ExpDateStr= CellValueToStr(xlSheetStr.Rows(RowInx).Columns(2).Value)
	ExpInputDateStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(3).Value)
	If ExpInputDateStr = "<TD>" Then
		ExpInputDateStr = TheCurrentDateStr
	End If
	'ExpOnHoldStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(4).Value)
	ExpDescriptionStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(5).Value)
	ExpChargeStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(6).Value)
	ExpTotAmountStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(7).Value)
	ExpPaymentStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(8).Value)
	ExpAdjustmentStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(9).Value)
	ExpRefundStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(10).Value)
	ExpBalanceStr = CellValueToStr(xlSheetStr.Rows(RowInx).Columns(11).Value)

End Sub

Sub GetLedgerValues()
		'Get the ledger values
	ActRowBillIDStr = CellValueToStr(row.Value(8)) 	'charge
	ActDateStr = CellValueToStr(row.Value(9))		'total amount
	If Len(CellValueToStr(row.Value(10))) > 0 Then
		InputDateSubStr = Split(CellValueToStr(row.Value(10)))
		ActInputDateStr = InputDateSubStr(0)
	Else
		ActInputDateStr = ""
	End If

	'ActOnHoldStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "OnHold")))
	ActDescriptionStr = LTrim(CellValueToStr(row.Value(14))) 'payment
	ActChargeStr = CellValueToStr(row.Value(15))		'adjustment
	ActTotAmountStr = CellValueToStr(row.Value(16))		'refund
	ActPaymentStr = CellValueToStr(row.Value(17))		'balance
	

	'Print Time	
	
	ActAdjustmentStr = CellValueToStr(row.Value(18))
	ActRefundStr = CellValueToStr(row.Value(19))
	ActBalanceStr = CellValueToStr(row.Value(24))

'
'	'Get the ledger values
'	ActRowBillIDStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Bill ID")))
'	ActDateStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Date")))
'	If Len(CellValueToStr(row.Value(GetColumnIndex(ledger, "Input Date")))) > 0 Then
'		InputDateSubStr = Split(CellValueToStr(row.Value(GetColumnIndex(ledger, "Input Date"))))
'		ActInputDateStr = InputDateSubStr(0)
'	Else
'		ActInputDateStr = ""
'	End If
'
'	'ActOnHoldStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "OnHold")))
'	ActDescriptionStr = LTrim(CellValueToStr(row.Value(GetColumnIndex(ledger, "Description"))))
'	ActChargeStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Charge")))
'	ActTotAmountStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Total Amt.")))
'	ActPaymentStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Payment")))
'	
'
'	'Print Time	
'	
'	ActAdjustmentStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Adjustment")))
'	ActRefundStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Refund")))
'	ActBalanceStr = CellValueToStr(row.Value(GetColumnIndex(ledger, "Balance")))
'
End Sub

Sub ValLedgerRow()

	Call ValidateLedgerCell(ExpBillRowID, ActBillRowID, "Bill ID")
	Call ValidateLedgerCell(ExpDateStr, ActDateStr, "Date")
	Call ValidateLedgerCell(ExpInputDateStr, ActInputDateStr, "Input Date")
	Call ValidateLedgerCell(ExpOnHoldStr, ActOnHoldStr, "On Hold")
	Call ValidateLedgerCell(ExpDescriptionStr, ActDescriptionStr, "Description")
	Call ValidateLedgerCell(ExpChargeStr, ActChargeStr, "Charge")
	Call ValidateLedgerCell(ExpTotAmountStr, ActTotAmountStr, "Total Amt.")
	Call ValidateLedgerCell(ExpPaymentStr, ActPaymentStr, "Payment")
	Call ValidateLedgerCell(ExpAdjustmentStr,ActAdjustmentStr,"Adjustment")
	Call ValidateLedgerCell(ExpRefundStr,ActRefundStr,"Refund")
	Call ValidateLedgerCell(ExpBalanceStr,ActBalanceStr,"Balance")

End Sub

Sub ValidateTotals(ExpResultStr, TotalBeingTestedStr, ActResultStr)
	Dim ReportEventStr
	
 	If  ExpResultStr = ActResultStr Then
		NumPassedInt = NumPassedInt + 1
		ReportEventStr = TotalBeingTestedStr + " validated"
		Reporter.ReportEvent micPass, ReportEventStr, TotalBeingTestedStr
	Else
		ReportEventStr = "<<< " + TotalBeingTestedStr + " failed validation! >>> [Exp: " + ExpResultStr + "] [Act: " + ActResultStr + "]"
		Reporter.ReportEvent micFail, ReportEventStr, TotalBeingTestedStr
		NumFailedInt = NumFailedInt + 1
	End If
End Sub

Sub ValidateLedgerCell(ExpCellValStr, ActCellValStr, ColumnBeingTestedStr)
	Dim ReportEventStr
	
	' This sub simply checks expected cell value (from an excel spreadsheet sheet) and compares it to the actual cell value found at run time
	' and then writes it to both the UFT Run Results module and to a log file created in folder C:\QTP Logs folder with a date/time stamped filename
	' to keep the run results in unique files.  It's useful when debugging the script
	
	If ExpCellValStr = ActCellValStr Then
		NumPassedInt = NumPassedInt + 1
		ReportEventStr = ColumnBeingTestedStr + " validated"
		Reporter.ReportEvent micPass, ReportEventStr, ColumnBeingTestedStr
	Else
		ReportEventStr = "<<< " + ColumnBeingTestedStr + " failed validation! >>> [Exp: " + ExpCellValStr + "] [Act: " + ActCellValStr + "]"
		Reporter.ReportEvent micFail, ReportEventStr, ColumnBeingTestedStr
		NumFailedInt = NumFailedInt + 1
	End If
End Sub

Sub GetTotals()

	' All of these window id's referenced in "Static" are the window ids defined in PracticeRc.h
	' The GetROProperty method gets the value of the object at run time
	
	BalBalanceDueStr = Window("Nextech Main Window").Window("Patients").Static("BalanceDueLabel").GetROProperty("text")
	BalPrePaymentsStr  = Window("Nextech Main Window").Window("Patients").Static("PrepaymentsLabel").GetROProperty("text")
	BalTotalBalancesStr = Window("Nextech Main Window").Window("Patients").Static("TotalBalancesLabel").GetROProperty("text")
	BalChargesStr = Window("Nextech Main Window").Window("Patients").Static("ChargesAmountLabel").GetROProperty("text")
	BalAdjustmentsStr = Window("Nextech Main Window").Window("Patients").Static("AdjustmentAmountLabel").GetROProperty("text")
	BalNetChargesStr  = Window("Nextech Main Window").Window("Patients").Static("NetChargesLabel").GetROProperty("text")
	BalPaymentsStr  = Window("Nextech Main Window").Window("Patients").Static("PaymentsLabel").GetROProperty("text")
	BalRefundsStr  = Window("Nextech Main Window").Window("Patients").Static("RefundsLabel").GetROProperty("text")
	BalNetPaymentsStr = Window("Nextech Main Window").Window("Patients").Static("NetPaymentsLabel").GetROProperty("text")
End Sub

Sub ExpandLedger()
' Expands the ledger
	Window("Nextech Main Window").Window("Patients").ActiveX("NxDL_BillingLedger").WinObject("Afx:42000000:8b").Click 13,5
End Sub

Sub CollapseLedger()
' Collapse the ledger
	Window("Nextech Main Window").Window("Patients").ActiveX("NxDL_BillingLedger").WinObject("Afx:42000000:8b").Click 13,5
End Sub

Sub Void (VoidTypeStr)

Select Case VoidTypeStr

	Case "Bill" 
	
		' ugly, but I'm sure it can be refactored into cleaner code
		Window("Nextech Main Window").Window("Patients").ActiveX("NxDL_BillingLedger").WinObject("Afx:42000000:8b").Click 10, 26, micRightBtn @@ hightlight id_;_528584_;_script infofile_;_ZIP::ssf12.xml_;_
		WinMenu("ContextMenu").Select "Void Bill"
		
		chk_PassFail = Window("Nextech Main Window").Dialog("NexTech Practice").Static("This feature should be").Check (CheckPoint("z_Popup1")) @@ hightlight id_;_332618_;_script infofile_;_ZIP::ssf45.xml_;_
		Call Checkpoint_PassFail(chk_PassFail,"z_Popup1")	
		Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_3934062_;_script infofile_;_ZIP::ssf46.xml_;_

		chk_PassFail = Window("Nextech Main Window").Dialog("NexTech Practice").Static("This feature should be").Check (CheckPoint("z_Popup2")) @@ hightlight id_;_3999598_;_script infofile_;_ZIP::ssf47.xml_;_
		Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_398154_;_script infofile_;_ZIP::ssf48.xml_;_
		Call Checkpoint_PassFail(chk_PassFail,"z_Popup2")
		
		chk_PassFail = Window("Nextech Main Window").Dialog("NexTech Practice").Static("This feature should be").Check (CheckPoint("z_Popup3")) @@ hightlight id_;_463690_;_script infofile_;_ZIP::ssf49.xml_;_
		Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("Yes").Click @@ hightlight id_;_4065134_;_script infofile_;_ZIP::ssf50.xml_;_
		Call Checkpoint_PassFail(chk_PassFail,"z_Popup3")
'	
	Case "Charge" MsgBox ("Charge selected to Void")
	
	Case "Adjustment" MsgBox ("Adjustment selected to Void")
	
	Case "Payment" MsgBox ("Payment selected to Void")
	Case "Refund" MsgBox ("Refund selected to Void")
	
	
	Case Else MsgBox ("Bad VOID Type supplied")
	err
End Select

End Sub

Sub VoidAndCorrect (VoidTypeStr)

Select Case VoidTypeStr

	Case "Bill" 
	
		' ugly, but I'm sure it can be refactored into cleaner code
		Window("Nextech Main Window").Window("Patients").ActiveX("NxDL_BillingLedger").WinObject("Afx:42000000:8b").Click 10, 26, micRightBtn @@ hightlight id_;_528584_;_script infofile_;_ZIP::ssf12.xml_;_
		WinMenu("ContextMenu").Select "Void and Correct Bill"
		
		chk_PassFail = Window("Nextech Main Window").Dialog("NexTech Practice").Static("This feature should be").Check (CheckPoint("z_vcp1")) @@ hightlight id_;_5312994_;_script infofile_;_ZIP::ssf60.xml_;_
		Call Checkpoint_PassFail(chk_PassFail,"z_vcp1")		
		Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_2951334_;_script infofile_;_ZIP::ssf61.xml_;_
		chk_PassFail = Window("Nextech Main Window").Dialog("NexTech Practice").Static("This feature should be").Check (CheckPoint("z_vcp2")) @@ hightlight id_;_3016870_;_script infofile_;_ZIP::ssf62.xml_;_
		Call Checkpoint_PassFail(chk_PassFail,"z_vcp2")		
		Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_5378530_;_script infofile_;_ZIP::ssf63.xml_;_
		chk_PassFail = Window("Nextech Main Window").Dialog("NexTech Practice").Static("This feature should be").Check (CheckPoint("VCPopup3")) @@ hightlight id_;_3148976_;_script infofile_;_ZIP::ssf43.xml_;_
		Call Checkpoint_PassFail(chk_PassFail,"VCPopup3")		
		Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("Yes").Click @@ hightlight id_;_2623274_;_script infofile_;_ZIP::ssf44.xml_;_
	
	Case "Charge" MsgBox ("Charge selected to Void")
	
	Case "Adjustment" MsgBox ("Adjustment selected to Void")
	
	Case "Payment" MsgBox ("Payment selected to Void")
	Case "Refund" MsgBox ("Refund selected to Void")
	
	
	Case Else MsgBox ("Bad VOID Type supplied")
	err
End Select

End Sub

Sub UndoCorrection (CorrTypeStr)


Select Case CorrTypeStr
	Case "VBill"
		Window("Nextech Main Window").Window("Patients").ActiveX("NxDL_BillingLedger").WinObject("Afx:42000000:8b").Click 13, 26, micRightBtn
		WinMenu("ContextMenu").Select "Undo Correction"
		

		If Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").Exist (5) Then
			Call RecordResults(True, "Undo Correction")	
			Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click
		Else
			Call RecordResults(False, "Undo Correction FAILED!!")	
		End If
			
		If Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").Exist (5) Then
			Call RecordResults(True, "Undo Correction, yes absolutely sure")	
			Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click
		Else
			Call RecordResults(False, "Undo Correction FAILED!!")	
		End If

	Case "VCBill"
	
		Window("Nextech Main Window").Window("Patients").ActiveX("NxDL_BillingLedger").WinObject("Afx:42000000:8b").Click 13, 26, micRightBtn
		WinMenu("ContextMenu").Select "Undo Correction"


		If Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").Exist (5) Then
			Call RecordResults(True, "Undo Correction")	
			Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click
		Else
			Call RecordResults(False, "Undo Correction FAILED!!")	
		End If
			
		If Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").Exist (5) Then
			Call RecordResults(True, "Undo Correction, yes absolutely sure")	
			Window("Nextech Main Window").Dialog("regexpwndtitle:=Practice").WinButton("regexpwndtitle:=&Yes").Click
		Else
			Call RecordResults(False, "Undo Correction FAILED!!")	
		End If
 @@ hightlight id_;_1838962_;_script infofile_;_ZIP::ssf14.xml_;_
		
	Case "Charge"
	
	Case "Adjustment"
	
	Case "Payment"
	
	Case "Refund"
	
	Case Else MsgBox ("Bad Undo Correction Type supplied")
	err
End Select

End Sub

Sub Checkpoint_PassFail(Rslt,CPName)

	If Rslt Then
		NumPassedInt = NumPassedInt + 1
		RsltStr = "Passed popup " + CPName
		
	Else
		NumFailedInt = NumFailedInt + 1
		RsltStr = "Failed popup " + CPName
	End If
	
End Sub











