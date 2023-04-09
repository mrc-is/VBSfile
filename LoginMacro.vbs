On Error Resume Next

Do While usrINPT = ""
usrINPT = InputBox("1:user        ID" & vbcrlf & "2: game PWD" & vbcrlf & "3: emel  PWD","To quit type Q","Type 1,2 or 3")
If usrINPT = 1 Then useNo ="wakjowo@hotmail.com"
If usrINPT = 2 Then useNo ="tampines"
If usrINPT = 3 Then useNo ="Tamp1ne5"
If usrINPT = Q Then useNo ="Tamp1ne5"
Loop

Set sh = WScript.CreateObject("WScript.Shell")
WScript.Sleep 2000
sh.SendKeys useNo
sh.SendKeys "{ENTER}"
WScript.Quit



