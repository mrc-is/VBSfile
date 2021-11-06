
Set fso = CreateObject("Scripting.FileSystemObject")
For Each f In fso.GetFolder(FSO.GetAbsolutePathName(".")).Files
  extension = fso.GetExtensionName(f.Name)
  If LCase(extension) = "bat" Then
    basename = fso.GetBaseName(f.Name)
    'f.Name = Mid(basename, 4) & Left(basename, 3) & "." & extension
    f.Name = basename & extension & ".TXT" 
  End If
Next