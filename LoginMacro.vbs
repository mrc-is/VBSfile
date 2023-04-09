On Error Resume Next

Do While usrINPT = ""
usrINPT = InputBox("use1: wakjowo@hotmail.com" & vbcrlf & "use2: tampines" & vbcrlf & "use3: 601194","Choose Text to Copy","Type 1,2 or 3")
If usrINPT = 1 Then useNo ="wakjowo@hotmail.com"
If usrINPT = 2 Then useNo ="tampines"
If usrINPT = 3 Then useNo ="Tamp1ne5"
Loop

Set sh = WScript.CreateObject("WScript.Shell")
WScript.Sleep 2000
sh.SendKeys useNo
sh.SendKeys "{ENTER}"
WScript.Quit


