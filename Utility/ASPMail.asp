<%
Const GLOB_EMAILHOST = "mail.bayareaconsulting.biz"
Const GLOB_DEFAULT_EMAILNAME = "BayAreaConsulting"
Const GLOB_DEFAULT_EMAIL = "Support@bayareaconsulting.biz"
Const GLOB_ALLOW_ASPQMAIL = "yes"
Const ALLOW_ASPQMAIL = "yes"

'Function sends an email
'**************************************************************************
Function fn_SendEmail(ByVal strTo, ByVal strFrom, ByVal strSubject, ByVal strBody)
   
    
     Dim Mailer
     Dim blnMailSent
     
     Set Mailer = Server.CreateObject("CDONTS.NewMail")
     Mailer.BodyFormat = 0
     Mailer.MailFormat = 0
     Mailer.To = strTo
     Mailer.From   = strFrom
     Mailer.Subject    = strSubject
     Mailer.Body   = strBody 
     Mailer.Importance = 1
     Mailer.Send        
     
     Set Mailer = Nothing
     
End Function  'Send email

%>