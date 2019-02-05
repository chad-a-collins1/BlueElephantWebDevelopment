<%

Dim objCDOMail     
Set objCDOMail = Server.CreateObject("CDONTS.NewMail") 

' From
'***********************************************
objCDOMail.From = "Dan Rigney <dannr09@yahoo.com>"

' To
'*********************************************** 
objCDOMail.To = "chad_a_collins1@yahoo.com"  '"dannr09@sbcglobal.net" 

' Copy 
'***********************************************
'objCDOMail.Cc = "myFriend1@thereDomain.com;myFriend2@anotherDomain.com" 

' Blind Copy
'***********************************************
'objCDOMail.Bcc = "myFriend1@thereDomain.com;myFriend2@anotherDomain.com" 

' Subject
'*********************************************** 
objCDOMail.Subject = "Test CDONTS" 
 
' Set the e-mail body format (0=HTML 1=Text)
'***********************************************
objCDOMail.BodyFormat = 0
 
'Set the mail format (0=MIME 1=Text)
'***********************************************
objCDOMail.MailFormat = 0

' Body of the email 
'*********************************************** 
objCDOMail.Body = "<html><body><h1>HTML IS COOL</h1><br><img src=""http://www.bayareaconsulting.biz/picts/about.jpg""><br>Test CDONTS<br>Line 2" & "</body></html>" 
 
'Importance of the e-mail (0=Low, 1=Normal, 2=High)
'***********************************************
objCDOMail.Importance = 1  

' Send the email
'*********************************************** 
objCDOMail.Send 

Set objCDOMail = Nothing 
 

%>
