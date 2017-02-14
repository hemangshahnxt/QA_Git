
'Window("Nextech Main Window").WinMenu("Menu").Select "File;Exit"


Window("regexpwndtitle:=Nextech \(UFT.....DB\)").Close

If Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").Exist (3) Then
	Window("Nextech Main Window").Dialog("regexpwndtitle:=Nextech Practice").Winbutton("regexpwndtitle:=&Yes").Click	
End If

