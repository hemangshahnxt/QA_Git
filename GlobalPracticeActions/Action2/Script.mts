Dim sDBarg
Dim P1
Dim IsPresent

Dim StartUpMethod


SUMethod = Environment.Value("StartUpMethod")
P1=parameter("DBName")


sDBarg="r:" + P1
wait (3)

Set dr = CreateObject("Mercury.DeviceReplay")
If SUMethod = "Crayola" Then
	'oShell.strShortcutPath.Click
	'oShell.strShortcutPath.Click
	
    dr.KeyDown 29 'PRESS THE CONTROL KEY DOWN & HOLD "29" IS ASCII CODE FOR "CONTROL"
    dr.KeyDown 42 'PRESS THE SHIFT KEY DOWN & HOLD "42" IS ASCII CODE FOR "SHIFT"
    wait 1
    systemUtil.Run "C:\PracStation\Practice.exe",sDbarg
    wait 2'WAIT BEFORE RELEASING THE KEYS
    dr.KeyUp 29 'RELEAS THE CONTROL KEY "29" IS ASCII CODE FOR "CONTROL"
    dr.KeyUp 42 'PRESS THE SHIFT KEY "42" IS ASCII CODE FOR "SHIFT"
    wait 1 ' MAKE SURE THE KEYS ARE RELEASED
	

	Dialog("regexpwndtitle:=Log In").Dialog("regexpwndtitle:=Input").Activate
	Dialog("regexpwndtitle:=Log In").Dialog("regexpwndtitle:=Input").WinEdit("regexpwndclass:=Edit").SetSecure "56e3410873559047dd7bbeb5edfc41948aff03407add"
	Dialog("regexpwndtitle:=Log In").Dialog("regexpwndtitle:=Input").WinButton("regexpwndtitle:=&OK").Click
	
 @@ hightlight id_;_65976_;_script infofile_;_ZIP::ssf5.xml_;_
	If Dialog("regexpwndtitle:=Log In").Dialog("regexpwndtitle:=NexTech Practice").Exist (30) Then
		Dialog("regexpwndtitle:=Log In").Dialog("regexpwndtitle:=NexTech Practice").WinButton("regexpwndtitle:=Bypass backup").Click
	End If
Else
	SystemUtil.Run "C:\PracStation\Practice.exe", sDBarg
End If

wait (2)	

IsPresent = Dialog("Practice - Log In").Check (CheckPoint("Practice - Log In")) @@ hightlight id_;_1510652_;_script infofile_;_ZIP::ssf3.xml_;_
Call RecordResults(IsPresent, "Login opened ")

If Dialog("Practice - Log In").Dialog("text:=Nextech Practice").Static("regexpwndtitle:=A backup is currently.*").Exist(2) Then
	Dialog("Practice - Log In").Dialog("text:=Nextech Practice").WinButton("regexpwndtitle:=OK").Click
End If

Dialog("regexpwndtitle:= Log In", "text:=Practice - Log In").WinButton("regexpwndtitle:=&Log In").Click

If Dialog("regexpwndtitle:= Log In", "text:=Practice - Log In").Dialog("regexpwndtitle:=Server Busy").Exist Then
	While Dialog("regexpwndtitle:= Log In", "text:=Practice - Log In").Dialog("regexpwndtitle:=Server Busy").Exist
		Wait (.5)
		Dialog("regexpwndtitle:= Log In", "text:=Practice - Log In").Dialog("regexpwndtitle:=Server Busy").WinButton("regexpwndtitle:=&Retry").Click
	Wend
	Reporter.ReportNote "The usual retry login popup on wald appeared!"

'wait(30)
'dr.PressKey 205  ' presses the right-arrow key
'dr.PressKey 28   ' presses the enter key
'Reporter.ReportNote "Total time taken to execute this test minus the time insert for wait statement"

End If
Set dr = Nothing


