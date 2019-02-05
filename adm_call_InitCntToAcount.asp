<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!--#include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp"-->
<!--#include file="Utility/DBUtil.asp"-->
<!--#include file="Utility/Incrypt.asp"-->
<!--#include file="Utility/Random.asp"-->
<!--#include file="Utility/ASPMail.asp"-->
<!--#include file="BusinessLayer/BL_adm_setAccount.asp"-->
<!--#include file="BusinessLayer/BL_admContact.asp"-->
<!--#include file="BusinessLayer/BL_adm_getAccount.asp"-->
<!--#include file="BusinessLayer/BL_adm_setBillingInfo.asp"-->
<!--#include file="BusinessLayer/BL_adm_setProject.asp"-->


<%
On Error Resume Next

Dim arryTmp
Dim intConsultantId, lngCntctID
Dim lngRC
Dim lngDC
Dim tmpAID
Dim dbconTmp
Dim rs
Dim strSQL
	   

lngCntctID = Request.QueryString("cID")
intConsultantId = Session("ConsultantId")

If Request.Form("handleContact") = 1 Then
   Dim strTo
   Dim strFrom
   Dim strSubject
   Dim strBody
   Dim strUID
   Dim strPASSWD
   Dim strEmail


   lngRC = fn_BL_adm_InitContactToAccount(lngCntctID, intConsultantId, arryTmp)

If Err.Number <> 0 Then
  Response.Write "after BL"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if

   If lngRC <> 0 Then
      Response.Write "Problem Creating Account"
      Response.End
   End If

   lngDC = fn_DL_admDeleteInitContact(lngCntctID)
   If lngDC <> 0 Then
      Response.Write "Problem Deleting Initial Contact"
      Response.End
   End If  
   
 If Err.Number <> 0 Then
  Response.Write "after del"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if        

   tmpAID = fn_getAccountID


    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
	   
    strSQL = "SELECT * FROM Account, BillingInfo WHERE BillingInfo.account_id = Account.account_id AND Account.account_id = " & tmpAID
    Set rs = CreateObject("ADODB.RecordSet")
    rs.Open strSQL, dbconTmp, 3, 3
    strUID = rs.Fields("username")
    strPASSWD = rs.Fields("account_password")
    strEmail = rs.Fields("email")

If Err.Number <> 0 Then
  Response.Write "after select"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if


'Auto email response parameters
   strTo = strEmail
   strFrom = "Blue Elephant<support@blueelephantwebdev.com>"
   strSubject = "New Account Information"
   strBody = "<p>Welcome to Blue Elephant Web Development! We value your business and we look forward to working with you.</p>" 
   strBody = strBody & "<p>The first step is for you to log into your new account and to follow the instructions that are provided for you after you have logged in. To log in to your account, go to <a href=" & "'" & "http://www.blueelephantwebdev.com" & "'" & "> blueelephantwebdev.com</a> and click the CLIENT LOGIN menu option. Please use the login information below to access you account.</p>"
   strBody = strBody & "<p><font color=red><b>NOTE: Please save this login information as you will need it during the course of your project development.</b></font></p>"
   strBody = strBody & "<br><b>USERID: " & strUID & "</b>"
   strBody = strBody & "<br><b>PASSWORD: " & strPASSWD & "</b>"
   strBody = strBody & "<br><br>Thank you for you interest in Blue Elephant Web Development." 

   Call fn_SendEmail(strTo, strFrom, strSubject, strBody)

If Err.Number <> 0 Then
  Response.Write "after send mail"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if

   rs.Close
   Set rs = Nothing
   dbconTmp.Close
   Set dbconTmp = Nothing

   Response.Write "Account Created Successfully"
       
ElseIf Request.Form("handleContact") = 2 Then

   lngRC = fn_DL_admDeleteInitContact(lngCntctID)
   If lngRC <> 0 Then
      Response.Write "Problem Deleting Initial Contact"
      Response.End
   End If
   
   Response.Write "Initial Contact Deleted Successfully"

End If 


If Err.Number <> 0 Then
  Response.Write "Err=" & Err.Description
End if


%>
















