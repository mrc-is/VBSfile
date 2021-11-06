dTimer=InputBox("Enter timer interval in minutes","Set Timer") 'minutes

do until IsNumeric(dTimer)=True
  dTimer=InputBox("Invalid Entry" & vbnewline & vbnewline & _ 
         "Enter timer interval in minutes","Set Timer") 'minutes
loop

if dTimer()"" then 'change () to brackets before run program
do
  WScript.Sleep dTimer*60*1000 'convert from minutes to milliseconds
  t=MsgBox("Take a Walk." & vbnewline & vbnewline & "Restart Timer?", _
    vbYesNo, "It's been " & dTimer &" minute(s)")
  if t=6 then 'if yes
       'continue loop
  else 'exit loop
       exit do
  end if
loop
end if