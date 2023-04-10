On Error Resume Next

Code = "223610"

Do While usrINPT = ""
usrINPT = InputBox("1: user        ID" & vbcrlf & "2: game PWD" & vbcrlf & "3: emel  PWD","To quit type Q + 2FA code type C","Type 1,2 or 3")
If usrINPT = 1 Then useNo ="wakjowo@hotmail.com"
If usrINPT = 2 Then useNo ="tampines"
If usrINPT = 3 Then useNo ="Tamp1ne5"
If LCase(usrINPT) = "c" Then useNo = Code
If usrINPT = Q Then WScript.Quit
Loop

Set sh = WScript.CreateObject("WScript.Shell")
WScript.Sleep 2000
sh.SendKeys useNo
sh.SendKeys "{ENTER}"
WScript.Quit



