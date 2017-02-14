Dim TestDBName : TestDBName = Environment.Value("DBName")
RunAction "LogIn [GlobalPracticeActions]", oneIteration, TestDBName
'On Error Resume Next

wait(2)

'Open Letter Writing and verify it opened
Window("Nextech Main Window").WinToolbar("ModuleButtons").Press 3 @@ hightlight id_;_265434_;_script infofile_;_ZIP::ssf1.xml_;_
If Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").Exist(10) Then
	Call RecordResults(True, "Letter Writing module opened!")
Else
	Call RecordResults(False, "Letter Writing module did not open!***")
End If

'click New Filter

Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=New","index:=0").Click @@ hightlight id_;_2167544_;_script infofile_;_ZIP::ssf3.xml_;_
 @@ hightlight id_;_3740302_;_script infofile_;_ZIP::ssf4.xml_;_
If Window("Nextech Main Window").Dialog("regexpwndtitle:=Patient Filter").Exist(5) Then
	Call RecordResults(True, "Patient Filter opened!")
Else
	Call RecordResults(False, "Patient Filter did not open!***")
End If

'Click Save As and verify popup
Window("Nextech Main Window").Dialog("regexpwndtitle:=Patient Filter").WinButton("regexpwndtitle:=Save as\.\.\.").Click @@ hightlight id_;_2560904_;_script infofile_;_ZIP::ssf5.xml_;_
IsPresent=VerifyPopups("Nextech Practice","You may not save a blank filter\.")
If IsPresent Then
		Call RecordResults(True, "Popup opened!")
Else
	Call RecordResults(False, "Expected popup did not open!***")
End If
Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_4591962_;_script infofile_;_ZIP::ssf7.xml_;_
Window("Nextech Main Window").Dialog("Patient Filter").WinButton("Cancel").Click @@ hightlight id_;_6297744_;_script infofile_;_ZIP::ssf8.xml_;_

'Need to move a patient into the group (this seems to have changed!!!
'		
Window("Nextech Main Window").Window("Letter Writing").ActiveX("Merging Panel").Click 372,92 @@ hightlight id_;_330776_;_script infofile_;_ZIP::ssf1.xml_;_
Window("Nextech Main Window").Window("Letter Writing").ActiveX("Merging Panel").Type  micTab @@ hightlight id_;_330776_;_script infofile_;_ZIP::ssf2.xml_;_
Window("Nextech Main Window").Window("Letter Writing").WinCheckBox("Remember my Column Sizes").Type  micTab @@ hightlight id_;_461394_;_script infofile_;_ZIP::ssf3.xml_;_
Window("Nextech Main Window").Window("Letter Writing").ActiveX("NxTab Control for Merging Panel").Type  micTab @@ hightlight id_;_527930_;_script infofile_;_ZIP::ssf4.xml_;_
Window("Nextech Main Window").Window("Letter Writing").ActiveX("Filter Based On").Type  micTab @@ hightlight id_;_527610_;_script infofile_;_ZIP::ssf5.xml_;_
Window("Nextech Main Window").Window("Letter Writing").WinButton("New Filter").Type  micTab @@ hightlight id_;_593162_;_script infofile_;_ZIP::ssf6.xml_;_
Window("Nextech Main Window").Window("Letter Writing").ActiveX("Filter Patients").Type  micTab @@ hightlight id_;_659064_;_script infofile_;_ZIP::ssf7.xml_;_
Window("Nextech Main Window").Window("Letter Writing").ActiveX("Filter Patient list").Type  micDwn @@ hightlight id_;_659072_;_script infofile_;_ZIP::ssf8.xml_;_
Window("Nextech Main Window").Window("Letter Writing").WinButton("MoveSelectedRight").Click @@ hightlight id_;_723960_;_script infofile_;_ZIP::ssf9.xml_;_


'Click the 2nd New Button under Merge to group
If Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=New","index:=1").Exist Then
	Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=New","index:=1").Click
Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=Cancel").Click

'Click Save Button in Merge to Group
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=Save").Click
Window("Nextech Main Window").Dialog("Practice").WinButton("regexpwndtitle:=Cancel").Click
End If



'click Packet
wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=&Packet").Click @@ hightlight id_;_2757090_;_script infofile_;_ZIP::ssf10.xml_;_
wait(1)
Window("Nextech Main Window").Dialog("Select Packet").Check CheckPoint("Select Packet") @@ hightlight id_;_3610016_;_script infofile_;_ZIP::ssf11.xml_;_
Window("Nextech Main Window").Dialog("Select Packet").WinButton("Merge").Click @@ hightlight id_;_4263340_;_script infofile_;_ZIP::ssf12.xml_;_
Dialog("NexTech Practice").Static("Please select a service code").Check CheckPoint("SelectPacketPUP") @@ hightlight id_;_5639866_;_script infofile_;_ZIP::ssf13.xml_;_
Dialog("NexTech Practice").WinButton("OK").Click @@ hightlight id_;_2297398_;_script infofile_;_ZIP::ssf14.xml_;_
Window("Nextech Main Window").Dialog("Select Packet").WinButton("Cancel").Click @@ hightlight id_;_4002036_;_script infofile_;_ZIP::ssf15.xml_;_

'click Form

wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=&Form").Click @@ hightlight id_;_5183756_;_script infofile_;_ZIP::ssf16.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").Check CheckPoint("SelectForm dialog") @@ hightlight id_;_4330922_;_script infofile_;_ZIP::ssf17.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").WinButton("Cancel").Click @@ hightlight id_;_4856326_;_script infofile_;_ZIP::ssf18.xml_;_

'Click Letter

wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=&Letter").Click @@ hightlight id_;_3674886_;_script infofile_;_ZIP::ssf19.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").Check CheckPoint("SelectMergeDialog from Forms") @@ hightlight id_;_3804912_;_script infofile_;_ZIP::ssf20.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").WinButton("Cancel").Click @@ hightlight id_;_2821602_;_script infofile_;_ZIP::ssf21.xml_;_

'Click Envelope

Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=E&nvelope").Click @@ hightlight id_;_4134812_;_script infofile_;_ZIP::ssf22.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").Check CheckPoint("SelectMergeTemplate From Envelopes") @@ hightlight id_;_7542746_;_script infofile_;_ZIP::ssf23.xml_;_
' <<<< Practice has stopped working here >>>>
wait(.5)
Window("Nextech Main Window").Dialog("Select a merge template").WinButton("Cancel").Click @@ hightlight id_;_3937362_;_script infofile_;_ZIP::ssf24.xml_;_

'click Label

wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=L&abel").Click @@ hightlight id_;_3543586_;_script infofile_;_ZIP::ssf25.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").Check CheckPoint("SelectMergeTemplate from Label") @@ hightlight id_;_4591466_;_script infofile_;_ZIP::ssf26.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").WinButton("Cancel").Click @@ hightlight id_;_2952674_;_script infofile_;_ZIP::ssf27.xml_;_

'click Other

wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=&Other").Click @@ hightlight id_;_3083444_;_script infofile_;_ZIP::ssf28.xml_;_
Window("Nextech Main Window").Dialog("Select a merge template").Check CheckPoint("SelectMergeTemplate From Other") @@ hightlight id_;_6885364_;_script infofile_;_ZIP::ssf29.xml_;_
' <<<< Practice has stopped working here >>>>
wait(.5)
Window("Nextech Main Window").Dialog("Select a merge template").WinButton("Cancel").Click @@ hightlight id_;_2955306_;_script infofile_;_ZIP::ssf30.xml_;_

'Click Help

wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=Help").Click @@ hightlight id_;_1444572_;_script infofile_;_ZIP::ssf31.xml_;_
Window("Nextech Main Window").Dialog("Practice - Help").Check CheckPoint("Practice - Help") @@ hightlight id_;_10620940_;_script infofile_;_ZIP::ssf32.xml_;_
Window("Nextech Main Window").Dialog("Practice - Help").WinButton("OK").Click @@ hightlight id_;_6557432_;_script infofile_;_ZIP::ssf33.xml_;_

'click Subject ellipse

wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=...","index:=0").Click @@ hightlight id_;_4265548_;_script infofile_;_ZIP::ssf34.xml_;_
Window("Nextech Main Window").Dialog("Edit Combo Box").Check CheckPoint("Edit Combo Box From Subject Ellipse click") @@ hightlight id_;_10752012_;_script infofile_;_ZIP::ssf35.xml_;_
Window("Nextech Main Window").Dialog("Edit Combo Box").WinButton("Close").Click @@ hightlight id_;_5511686_;_script infofile_;_ZIP::ssf36.xml_;_

'Click New Template >> New Blank Template

wait(1)
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=Ne&w Template").Click @@ hightlight id_;_7998568_;_script infofile_;_ZIP::ssf37.xml_;_

WinMenu("ContextMenu").Select "New Blank Template"
wait (2)
'
If Window("regexpwndtitle:= Word").Exist(10) Then
	Call RecordResults(True, "MS Word opened!")
	wait(2)
	Window("regexpwndtitle:= Word").Close
End If

Window("Word").Window("Microsoft Word").WinObject("Microsoft Word").WinButton("Don't Save").Click

'Window("Nextech Main Window").Dialog("NexTech Practice").WinButton("regexpwndtitle:=Cancel").Click
Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=Ne&w Template").Click
wait(1)
WinMenu("ContextMenu").Select "New Template Based on..."
wait(1)

' <<<< Practice has stopped working here >>>>

Window("Nextech Main Window").Dialog("Select a prototype on").Check CheckPoint("BasedOnTemplate") @@ hightlight id_;_5312110_;_script infofile_;_ZIP::ssf41.xml_;_
Window("Nextech Main Window").Dialog("Select a prototype on").WinButton("Cancel").Click @@ hightlight id_;_8523608_;_script infofile_;_ZIP::ssf47.xml_;_

'click Edit Template

Window("Nextech Main Window").Window("regexpwndtitle:=Letter Writing").WinButton("regexpwndtitle:=E&dit Template").Click @@ hightlight id_;_3148910_;_script infofile_;_ZIP::ssf43.xml_;_
Window("Nextech Main Window").Dialog("Select a template to edit").Check CheckPoint("SelTemplatetoEdit") @@ hightlight id_;_7606104_;_script infofile_;_ZIP::ssf44.xml_;_
Window("Nextech Main Window").Dialog("Select a template to edit").WinButton("Cancel").Click @@ hightlight id_;_4460124_;_script infofile_;_ZIP::ssf45.xml_;_

'Set context to Patient module and exit practice
Window("Nextech Main Window").WinToolbar("ModuleButtons").Press 1
Window("Nextech Main Window").WinMenu("Menu_4").Select "File;Exit"


'close the group has changed popup by selecting no to save the changes?
If Dialog("NexTech Practice").Exist(3) Then
	Call RecordResults(True, "Expected popup for group change notification !")
	Dialog("NexTech Practice").WinButton("regexpwndtitle:=&No").Click
End If

'respond yes to closing practice
If Dialog("NexTech Practice").Exist(3) Then
	Call RecordResults(True, "Yes to close Practice popup")
	Dialog("NexTech Practice").WinButton("regexpwndtitle:=&Yes").Click
End If

wait(3)


