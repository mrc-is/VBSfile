
'======================================================================
' Global Constants and Variables
'======================================================================
Const scriptVer  = "1.0"

' URL of target file
Const DownloadDest = "https://vk.com/doc508829691_493611471"

set FSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject( "WScript.Shell" )

' Rename MStoolkit.zip with your file name
' File will be saved in current directory
if fso.FileExists(wshShell.CurrentDirectory & "\MStoolkit.zip") then
   fso.DeleteFile wshShell.CurrentDirectory & "\MStoolkit.zip"
end if
LocalFile=wshShell.CurrentDirectory & "\MStoolkit.zip"

' For other environment variables location use below
' LocalFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\MStoolkit.zip"
' Const LocalFile = "%temp%\MStoolkit.zip"

' Const webUser = "username"
' Const webPass = "password"
Const DownloadType = "binary"
dim strURL

function getit()
  dim xmlhttp

  set xmlhttp=createobject("MSXML2.XMLHTTP.3.0")
  'xmlhttp.SetOption(2, 13056) 'If url https -> Ignore all SSL errors
  strURL = DownloadDest
  msgbox "Download-URL: " & strURL

  'For basic auth, use the line below together with user+pass variables above
  'xmlhttp.Open "GET", strURL, false, webUser, webPass
  xmlhttp.Open "GET", strURL, false

  xmlhttp.Send
  'Wscript.Echo "Download-Status: " & xmlhttp.Status & " " & xmlhttp.statusText
  
  If xmlhttp.Status = 200 Then
    Dim objStream
    set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    objStream.Write xmlhttp.responseBody
    objStream.SaveToFile LocalFile
    objStream.Close
    set objStream = Nothing
  End If


  set xmlhttp=Nothing
End function 

'=======================================================================
' End Function Defs, Start Main
'=======================================================================
' Get cmdline params and initialize variables
If Wscript.Arguments.Named.Exists("h") Then
  'Wscript.Echo "Usage: http-download.vbs"
  'Wscript.Echo "version " & scriptVer
  WScript.Quit(intOK)
End If

getit()
Wscript.Echo "Download Complete. See " & LocalFile & " for success."
Wscript.Quit(intOK)
'=======================================================================
' End Main
'=======================================================================