'On Error Resume Next

' Current foldername
Dim filesys
Dim fso, folder, file, folderName, dict, wshShell
Set filesys = CreateObject("Scripting.FileSystemObject")
varPathCurrent = filesys.GetParentFolderName(WScript.ScriptFullName)
varPathParent = filesys.GetParentFolderName(varPathCurrent)
varPathGrandParent = filesys.GetParentFolderName(varPathParent)
varNameFolderCurrent = mid(varPathCurrent, len(varPathParent) + 2 , len(varPathCurrent) - len(varPathParent) - 1)
varNameFolderParent = mid(varPathCurrent, len(varPathGrandParent) + 2 , len(varPathCurrent) - len(varPathGrandParent) - len(varNameFolderCurrent) - 2)
SuggestName = varNameFolderCurrent &"_"& varNameFolderParent

'Path for folder name
Set wshShell = CreateObject( "WScript.Shell" )
folderName = wshShell.CurrentDirectory
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(folderName)

'Count how many original
counter = 0
extension = "pdf"
For Each File In Folder.Files  
	if LCase((FSO.GetExtensionName(File))) = LCase(extension) then  
		counter = counter + 1
	end if  
Next
	if counter >= 2 then
		Wscript.Echo "We can only copy SINGLE file.Please remove other PDF"
  		WScript.Quit
	end if


'Input number of copies
number = Int(InputBox("number of copies?"))
count = 1

'Input user agree naming convention
intAnswer = _
    Msgbox("Do you agree with filename below?" & vbCrLf & SuggestName , _
        vbYesNo, "Confirm filename")
If intAnswer = vbNo Then
	' Retain current filename
    For Each File In Folder.Files 
	If Right(File.name, 4) = ".pdf"  Then
	CurrName = FSO.GetFileName(file.name)
	SuggestName = fso.getbasename(file.name)
	End If
    Next
Else
End If

For Each File In Folder.Files
	If Right(File.name, 4) = ".pdf"  Then
      Do Until count > number
		If Len(count) = 2 Then
    			ncount = string(2 - Len(inputStr), "0") & count
		ElseIf Len(count) = 3 Then
    			ncount = string(1 - Len(inputStr), "0") & count
		ElseIf Len(count) = 4 Then
    			ncount = string(0 - Len(inputStr), "0") & count
		Else
			ncount = string(3 - Len(inputStr), "0") & count
		End If
      		FSO.CopyFile file.Name, SuggestName &"_"& ncount & ".pdf"
      count = count + 1
      Loop
	End If

Next

Set objShell = Nothing
Set FSO = Nothing