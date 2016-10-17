Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)
objMail.Display   'To display message
objMail.To = "recipient@example.com"
objMail.Recipients.Add ("recipient1@example.com")
objMail.cc = "copyrecipient@example.com"
objMail.Subject = "Test Mail Subject"
objMail.Body = "Hi All,"  & vbCrLf & "this is a test mail"
'objMail.Body=InputBox("Please enter the body of the message")
'objMail.Attachments.Add("C:\Attachment\abc.jpg")   'Make sure attachment exists at given path. Then uncomment this line.
'objMail.Send   'I intentionally commented this line
'objOutlook.Quit
Set objMail = Nothing
Set objOutlook = Nothing