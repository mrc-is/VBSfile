Dim speaks, speech
hour_now=hour(time)
if hour_now > 12 or hour_now = 12 Then
hour12 = hour_now - 12
ampm = "PM"
Else
hour12 = hour_now
hourfinal = hour_now
ampm = "AM"
End If
If hour12 = 10 Then
hourfinal = "Ten"
Elseif hour12 = 11 Then
hourfinal = "Eleven"
Elseif hour12 = 12 Then
hourfinal = "Twelve"
Elseif hour12 = 0 Then
hourfinal = "Twelve"
Elseif hour12 > 0 and hour12 < 10 Then
hourfinal = hour12
End If
speaks = "It is " & hourfinal & " o clock " & ampm
Set speech=CreateObject("sapi.spvoice")
speech.Speak speaks