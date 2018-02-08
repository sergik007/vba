dim SavePath
dim Subject
dim FileExtension
dim k
dim ReceivedTime

SavePath = "C:\Temp\mails\"
mailBoxName ="siarheikalashynskifail2@gmail.com"

'open outlook.exe
Set WshShell = WScript.CreateObject ("WScript.Shell")
Set colProcessList = GetObject("Winmgmts:").ExecQuery ("Select * from Win32_Process")

For Each objProcess in colProcessList
	If objProcess.name = "outlook.exe" then
		vFound = True
	End if
Next

If vFound = True then
	WshShell.Run (".exe location")
End If


Set objOutlook = CreateObject("Outlook.Application")
For Each oAccount In objOutlook.Session.Accounts
  If oaccount = mailBoxName then
		Set store = oaccount.DeliveryStore
		Set folder = store.GetDefaultFolder(6) 'here it selects the inbox folder of account.
		'Set objNamespace = objOutlook.GetNamespace("MAPI")
		'Set folder = objNamespace.GetDefaultFolder(6) 'Inbox
		'Set colItems = folder.Items
		Set colItems = folder.Items
		Set colFilteredItems = colItems.Restrict("[Unread]=true")
		'Set colFilteredItems = colFilteredItems.Restrict("[Subject] = " & Subject)

		For k = colFilteredItems.Count to 1 step -1
			set messageObj  = colFilteredItems.Item(k)
			intCount = messageObj.Attachments.Count
				If intCount > 0 Then
					For i = 1 To intCount
						if right(Ucase(messageObj.Attachments.Item(i)),3) = "CSV" then
							fileName = Year(messageObj.ReceivedTime) & "_" & Month(messageObj.ReceivedTime) & "_" & Day(messageObj.ReceivedTime) & "_" & Hour(messageObj.ReceivedTime) & "_" & Minute(messageObj.ReceivedTime) & "_" & Second(messageObj.ReceivedTime) & ".csv"
							messageObj.Attachments.Item(i).SaveAsFile SavePath & fileName
						End If
					Next
					messageObj.Unread = False
				End If
		next
  end if
next