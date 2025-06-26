Set WshShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

currentFolder = fso.GetAbsolutePathName(".")

zipFile = currentFolder & "\_internal.zip"
destFolder = currentFolder

' Ensure zip file exists
If Not fso.FileExists(zipFile) Then
    MsgBox "Zip file doesn't exist: " & zipFile, vbCritical, "Error"
    WScript.Quit
End If

' Create destination folder if needed
If Not fso.FolderExists(destFolder) Then
    fso.CreateFolder(destFolder)
End If

Set objZip = objShell.NameSpace(zipFile)
Set objDest = objShell.NameSpace(destFolder)

Dim response
response = MsgBox("Heyyy, let little owe me install Niko for you?", vbYesNo + vbQuestion, "Install Niko?")

If response = vbYes Then
    MsgBox "Hey calm down, we're just doing the install for you", vbOKOnly + vbInformation, "Installing Niko (Amelia's version :D)"
    MsgBox "Psssst, so doing python -> exe is really hard, just be patient", vbOKOnly + vbInformation, "Installing Niko (Amelia's version :D)"
    If Not objZip Is Nothing Then
        objDest.CopyHere objZip.Items, 4
        ' Optional wait before launching EXE
        WScript.Sleep 3000
        WshShell.Run "main.exe", 0, False
        MsgBox "Phew, installation done, when you want to run the app again just double click 'main.exe' ok?", vbOKOnly + vbInformation, "Installing Niko (Amelia's version :D)"
    Else
        MsgBox "Zip file invalid or couldn't be opened.", vbCritical, "Error"
    End If
Else
    MsgBox "Baiee", vbOKOnly + vbExclamation, "Alright see you then"
End If
