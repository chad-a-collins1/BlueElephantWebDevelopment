<%
Dim oMail


Set oMail = server.CreateObject("MEMail.Message")


oMail.MailFrom = "dannr09@yahoo.com"

oMail.MailFromDisplayName = "Daniel Rigney"

oMail.ContentType = "text/html;"

oMail.MailTo = "dannr09@sbcglobal.net"

oMail.Subject = "Welcome to our service"

oMail.MessageBody = "<html><body><h1>Hello there,<BR>Welcome to our new service.</h1></body></html>"

oMail.SendMessage
%>
