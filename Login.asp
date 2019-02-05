<%@
LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp" -->
<!--#include file="Utility/DBUtil.asp"-->
<!--#include file="BusinessLayer/BL_Login.asp" -->
<!-- #include file="Utility/ASPMail.asp" -->


<%
Const ACTION_FORGOTPASSWD = "lasj8gd6328bd"
Const ACTION_FETCHPASSWORD = "kqA1hs7shSAd08"
%>


<%
'Sub Main
'*********************************************
Sub Main()

On Error Resume Next
  
   
   'Call sub_SetUserAgent


   Dim strQry
   Dim strAct
   Dim strLP
   Dim intRC
   Dim intAID
   
   strAct = Request(QSVAR & "1")

   Select Case strAct
   
      Case ACTION_VALIDATE:          
            Call sub_ValidateLogin
            
      Case "", ACTION_MAIN: 
            'Call sub_CheckForLogin
            Call DisplayHeader(TITLE_LOGIN, PIC_LOGIN, LOGO_GRAPHIC)
            Call DisplayLoginMain ("")
            Call DisplayFooter
      
      Case ACTION_LOGOUT:
            Call sub_ValidateLogout      

      Case ACTION_FORGOTPASSWD:
            Call DisplayHeader(TITLE_LOGIN, PIC_LOGIN, LOGO_GRAPHIC)
            Call DisplayForgotLogin
            Call DisplayFooter	    
            
      Case ACTION_FETCHPASSWORD:
            Call DisplayHeader(TITLE_LOGIN, PIC_LOGIN, LOGO_GRAPHIC)
            Call DisplayValidateEmailAddress
            Call DisplayFooter	                
            
      Case Else:
           Call sub_HandleLogicError(ERR_INVALID_ACTION)
           
   End Select
   
   
  If Err.Number <> 0 Then
     Response.Write "Error = " & Err.Number & ", " & Err.Description
     Response.End 
  End If 
   
  'sub_ErrorCatch
   

End Sub


' This Sub checks the last page cookie and if it is a valid referring 
' page for the current page then the current page becomes the new last page   
'*******************************************************************************   
Sub sub_CheckForLogin(strPage) 
   
   Dim blnLoggedIn
   blnLoggedIn = Session("sess_blnLoggedIn")
   
   Select Case blnLoggedIn
     
      Case True:
         ' was logged in and came back to Login page probably using back button
         Dim arryVar(1)
         arryVar(0) = ACTION_LOGOUT
         arryVar(1) = fn_GetRandomAlphaNumeric(11,16)
         Response.Redirect fn_CreateURL(True, PAGE_LOGIN, arryVar)
              
      Case Else:
          ' Do Nothing
          
   End Select

End Sub
 
 

'Sub Validate Login
'*********************************************
Sub sub_ValidateLogin()

   Dim strUN   
   Dim strPwd    
   Dim lngAcctId    
   Dim strAcctStatus
   Dim strAcctType
   Dim strFN
   Dim strLN
   Dim lngRC
   Dim strResponse

  strUN = Request("txtUsername")
  strPwd = Request("txtPassword")
 
  lngRC = fn_BL_ValidateLogin(strUN, strPwd, lngAcctId, strAcctStatus, strAcctType, strFN, strLN)        
        
  Select Case lngRC
    Case 0:  
       
       lngRC = fn_BL_DBLogIn(lngAcctId, Session.SessionID)
       If lngRC <> 0 Then
          Call sub_HandleLogicError(ERR_DBINSERT_USERLOGIN)
       End If
       
       Session("sess_blnLoggedIn") = True
       Session("sess_strUN") = strUN
       Session("sess_lngAcctId") = lngAcctId
       Session("sess_strAcctStatus") = strAcctStatus
       Session("sess_strAcctType") = strAcctType
       Session("sess_strFirstName") = strFN
       Session("sess_strLastName") = strLN
       
       Dim arryVar(1)
       arryVar(0) = ACTION_MAIN
       arryVar(1) = fn_GetRandomAlphaNumeric(11,16)
  
       Response.Redirect fn_CreateURL(True, PAGE_PORTAL, arryVar)
       
    Case ERR_INVALID_UID, ERR_INVALID_PWD:
          Call DisplayHeader(TITLE_LOGIN, PIC_LOGIN, LOGO_GRAPHIC)
          Call DisplayLoginMain ("Invalid Login! Please make sure your username and password are spelled correctly!")
          Call DisplayFooter
          Exit Sub

    Case Else:
         Call sub_HandleLogicError(ERR_INVALID_VAP_RC)

   End Select


End Sub    ' sub_ValidateLogin




'Sub DisplayLoginMain
'*********************************************
Sub DisplayLoginMain(strError)
   Dim arryVar(1)
   Dim arryVarForgot(1)
   arryVar(0) = ACTION_VALIDATE
   arryVar(1) = fn_GetRandomAlphaNumeric(11,16)
   arryVarForgot(0) = ACTION_FORGOTPASSWD
   arryVarForgot(1) = fn_GetRandomAlphaNumeric(11,16)
%>
<table width="600">
  <tr>
    <td align="left">
   <form action="<% = fn_CreateURL(True,PAGE_LOGIN,arryVar) %>" method="post" name="theForm">
   <span class="smalltitle">Client Login</span>
     <table name="main" width="90%"  bgcolor="#FFFFFF">
       <tr>
         <td>
       	<% 
       	  If strError <> "" Then
       	     %>
       	      <font color="#FF0000"><b><I><% = strError %></I></b></font>
       	     <%
       	  End If
       	%>
       	 <br>
         <table cellpadding=2>	
       	  <tr>
       	    <td align="right"><b><font color="#666666">Username:</font></b></td><td align="left"><input type="text" name="txtUsername"></td>
       	  </tr>
       	  <tr>
       	    <td align="right"><b><font color="#666666">Password:</font></b></td><td align="left"><input type="password" name="txtPassword"></td>
       	  </tr>
       	  <tr>
             <td align="center">
                <input type="submit" value="Login"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
             </td></form>
             <td align="center">
		<form action="<% = fn_CreateURL(True,PAGE_LOGIN,arryVarForgot) %>" method="post" name="theForm2">
                <input type="submit" value="Help!" onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
             </td>
           </tr>
       	  <tr>
       	    <td colspan=4>&nbsp;</td>
       	  </tr>

       	</table>
    
     
         </td>
       </tr>
     </table>
   </form>
  </td>
 </tr>
</table>   
<% 
End Sub           'DisplayLoginMain


'Sub Validate LogOut
'*********************************************
Sub sub_ValidateLogout()   

   Dim lngAcctId
   Dim lngRC

   lngAcctId = Session("sess_lngAcctId")
   
   lngRC = fn_BL_DBLogOut(lngAcctId, Session.SessionID)
   If lngRC <> 0 Then
      Call sub_HandleLogicError(ERR_DBINSERT_USERLOGOUT)
   End If
   
   Session.Contents.RemoveAll
   
   Dim arryVar(1)
   arryVar(0) = ACTION_MAIN
   arryVar(1) = fn_GetRandomAlphaNumeric(11,16)

   Response.Redirect fn_CreateURL(True, PAGE_LOGIN, arryVar)


End Sub    ' sub_ValidateLogout



'*********************************************
Sub DisplayForgotLogin()
   Dim arryVar(1)
   arryVar(0) = ACTION_FETCHPASSWORD
   arryVar(1) = fn_GetRandomAlphaNumeric(11,16)
%>



<table width="600">
  <tr>
    <td align="left">
   <form action="<% = fn_CreateURL(True,PAGE_LOGIN,arryVar) %>" method="post" name="theForm">
   <span class="smalltitle">Password Assistance</span>
   <br><br><br>
     <table name="main" width="90%"  bgcolor="#FFFFFF">
       <tr>
         <td>
       	 <br>
         <table cellpadding=2>	
       	  <tr>
       	    <td align="left"><b><font color="#666666">Please enter your email address:</font></b></td>
	  </tr>
	  <tr>  
	    <td align="left"><input type="text" name="txtUsername" size="35"></td>
       	  </tr>
       	  <tr>
             <td align="center">
                <input type="submit" value="Submit"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
             </td></form>

           </tr>
       	  <tr>
       	    <td colspan=4>&nbsp;</td>
       	  </tr>

       	</table>
         </td>
       </tr>
     </table>
   </form>
  </td>
 </tr>
</table>   

<%
End Sub
%>

<%
'****************************************************
Sub DisplayValidateEmailAddress


   Dim strPwd    
   Dim strEmail
   Dim strUN
   Dim lngRC
   Dim strResponse

  strEmail = Request("txtUsername")
 
  lngRC = fn_BL_ValidateEmail(strEmail, strPwd, strUN)        
        
        
  Select Case lngRC

    Case 0:
          
          Dim strTo
          Dim strFrom
          Dim strSubject
          Dim strBody
          
			strTo = strEmail
   			strFrom = "Bay Area Consulting <support@bayareaconsulting.biz>"
   			strSubject = "Login Information"
   			strBody = "<p>Your username and password for your <a href=" & "'" & "http://www.bayareaconsulting.biz" & "'" & "> Bay Area Consulting</a> account are listed below. </p>" 
   			strBody = strBody & "<br>USERID<font color=blue>:<b> " & strUN & "</font></b>"
   			strBody = strBody & "<br>PASSWORD<font color=blue>:<b> " & strPwd & "</font></b>"
   			strBody = strBody & "<p>Please let us know if you have any addition problems accessing your account.</p>"

   			Call fn_SendEmail(strTo, strFrom, strSubject, strBody)      
          Response.Write "<br><br><br><font size=3 color=red><b>Your Username and Password have been emailed to you!</b></font>"

          Exit Sub
          
    Case 1:
          Call DisplayHeader(TITLE_LOGIN, PIC_LOGIN, LOGO_GRAPHIC)
          Call DisplayLoginMain ("Invalid Login! Please make sure your username and password are spelled correctly!")
          Call DisplayFooter
          Exit Sub

   End Select

End Sub
%>




<%
' Call the Main Sub Routine
'*******************************
Call Main


%>

























