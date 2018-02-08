dim SavePath
dim Subject
dim FileExtension
dim k

SavePath = "C:\Temp\mails\"

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
  If oaccount ="siarheikalashynskifail2@gmail.com" then
  
		Set objNamespace = objOutlook.GetNamespace("MAPI")
		Set objFolder = objNamespace.GetDefaultFolder(6) 'Inbox

		Set colItems = objFolder.Items
		Set colFilteredItems = colItems.Restrict("[Unread]=true")
		'Set colFilteredItems = colFilteredItems.Restrict("[Subject] = " & Subject)

		For k = colFilteredItems.Count to 1 step -1
		  set objMessage  = colFilteredItems.Item(k)
		  intCount = objMessage.Attachments.Count
			If intCount > 0 Then
				For i = 1 To intCount
					if right(Ucase(objMessage.Attachments.Item(i)),3) = "CSV" then
						curDate = Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_" & Hour(Time) & "_" & Minute(Time) & "_" & Second(Time) & "_" & k & ".csv"
						objMessage.Attachments.Item(i).SaveAsFile SavePath & curDate
					End If
				Next
				objMessage.Unread = False
			End If
		next
  end if
next
