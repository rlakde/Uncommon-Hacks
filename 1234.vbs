'Get the Count of Items in Inbox
Dim waittime : waittime = 5 * 60 * 1000
Set oShell = WScript.CreateObject ("WScript.Shell")
Set app = CreateObject("Outlook.Application")
Set nameSpace = app.GetNamespace("MAPI")
do
Set MyFolders = nameSpace.GetDefaultFolder(6)

'Read unread items in Inbox
Set cols = MyFolders.Items
MsgBox "Hii"
For each mail In cols
If mail.subject="Please switch off your PC before leaving" Then
	MsgBox mail.subject
	MsgBox mail.sendername
	MsgBox mail.body
	mail.delete
	oShell.run "powershell.exe -noexit WMIC /node:192.5.85.178 process call create 'powershell.exe /c shutdown -h'"
End If
Next
Wscript.Sleep(waittime)
loop