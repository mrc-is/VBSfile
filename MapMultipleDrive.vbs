Option Explicit
Dim objNetwork, strDrive, objShell, objUNC, objFSO
Dim strRemotePath1, strDriveLetter1, strNewName1
Dim strRemotePath2, strDriveLetter2, strNewName2
Dim strRemotePath3, strDriveLetter3, strNewName3
' Create more Dim if neccessary

' NEEDED FOR IF THEN
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("Wscript.Network") 

' DECLARING DRIVES AND UNC PATHS
strDriveLetter1 = "S:" 
strRemotePath1 = "\\mediacorp.grp\fileserver\media_research\info-sys" 
strNewName1 = "Sdrive"
' Change the Letter, Path, and Name

strDriveLetter2 = "T:" 
strRemotePath2 = "\\mediacorp.grp\fileserver\media_research\info-sys"
strNewName2 = "Tdrive"
' Change the Letter, Path, and Name

strDriveLetter3 = "U:" 
strRemotePath3 = "\\mediacorp.grp\fileserver\media_research\info-sys"
strNewName3 = "Udrive"
' Change the Letter, Path, and Name


' CHECKING TO SEE IF DRIVE IS MAPPED, IF IT IS IT IGNORES IT AND MOVES ON, IF IT IS NOT IT MAPS IT.

' Section to map the S network drive
If (objFSO.DriveExists("S:") = True) Then
    objNetwork.RemoveNetworkDrive "S:", True, True
End If
    Set objNetwork = CreateObject("WScript.Network") 
    objNetwork.MapNetworkDrive strDriveLetter1, strRemotePath1 

' Section which actually (re)names the S Mapped Drive
Set objShell = CreateObject("Shell.Application")
objShell.NameSpace(strDriveLetter1).Self.Name = strNewName1

' Section to map the T network drive
If (objFSO.DriveExists("T:") = True) Then
    objNetwork.RemoveNetworkDrive "T:", True, True
End If
    Set objNetwork = CreateObject("WScript.Network") 
    objNetwork.MapNetworkDrive strDriveLetter2, strRemotePath2 

' Section which actually (re)names the T Mapped Drive
Set objShell = CreateObject("Shell.Application")
objShell.NameSpace(strDriveLetter2).Self.Name = strNewName2

' Section to map the U network drive
If (objFSO.DriveExists("U:") = True) Then
    objNetwork.RemoveNetworkDrive "U:", True, True
End If
    Set objNetwork = CreateObject("WScript.Network") 
    objNetwork.MapNetworkDrive strDriveLetter3, strRemotePath3 

' Section which actually (re)names the U Mapped Drive
Set objShell = CreateObject("Shell.Application")
objShell.NameSpace(strDriveLetter3).Self.Name = strNewName3


WScript.Quit

' End of VBScript.