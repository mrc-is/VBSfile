Option Explicit
Dim objNetwork, strDrive, objShell, objUNC, objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("Wscript.Network") 
Dim strRemotePath, strDriveLetter, strNewName
strNewName = "MRC share"
strDriveLetter = "S:" 

If (objFSO.DriveExists("S:") = True) Then
objNetwork.RemoveNetworkDrive "S:", True, True
Else
	If (objFSO.DriveExists("S:") = False) Then
	Set objNetwork = CreateObject("WScript.Network") 
	objNetwork.MapNetworkDrive "S:" , "\\mediacorp.grp\fileserver\media_research"
	Set objShell = CreateObject("Shell.Application")
	objShell.NameSpace(strDriveLetter).Self.Name = strNewName
	Else
	End If
End If

WScript.Quit
    objNetwork.RemoveNetworkDrive "T:", True, True