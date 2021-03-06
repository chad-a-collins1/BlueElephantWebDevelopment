<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp" -->
<!--#include file="Utility/DBUtil.asp"-->
<!--#include file="BusinessLayer/BL_InitialContact.asp" -->
<!-- #include file="Utility/ASPMail.asp" -->
<!--#include file="DataLayer/DL_adm_getAccount.asp" -->

<%
'Sub Main
'*********************************************
Sub Main()

'On Error Resume Next

   'Call sub_SetUserAgent

   Dim strAct
   Dim intRC
   
   strAct = Request(QSVAR & "1")

   Select Case strAct
   
      Case ACTION_VALIDATE:
            Call sub_ValidateContact
            
      Case "", ACTION_MAIN: 
            Call DisplayHeader(TITLE_CONTACT, PIC_CONTACT, LOGO_GRAPHIC)
            Call DisplayContactMain ("")
            Call DisplayFooter
      
      Case ACTION_SUCCESS: 
            Call DisplayHeader(TITLE_CONTACT, PIC_CONTACT, LOGO_GRAPHIC)
            Call DisplayContactSuccess      
            Call DisplayFooter
            
      Case Else:
           Response.Redirect PAGE_ERROR & "?" & QSVAR & "1=" & ERR_INVALID_ACTION
           
   End Select
   
   'If Err.Number <> 0 Then
   '   Print "Err = " & Err.Number
   '   PrintStop "ErrDEsc = " & Err.Description  
   'End If  
   
  'Call sub_ErrorCatch
   

End Sub


'Sub Validate Contact Info
'*********************************************
Sub sub_ValidateContact()

   Dim strCompany  'As String
   Dim strName     'As String
   Dim strEmail      'As String
   Dim strPhone      'As String
   Dim strFax      'As String
   Dim strAddress      'As String
   Dim strCity      'As String
   Dim strState      'As String
   Dim strZip      'As String
   Dim strShort      'As String                     
   Dim intLogin   'As integer
   Dim intAID     'As Integer
   Dim intRC
   Dim strResponse
   Dim strTo
   Dim strFrom
   Dim strSubject
   Dim strBody
   Dim strUID
   Dim strPASSWD

   strCompany = Request.Form("txtCompany") 
   strName = Request.Form("txtName") 
   strEmail = Request.Form("txtEmail") 
   strPhone = Request.Form("txtPhone") 
   strFax = Request.Form("txtFax") 
   strAddress = Request.Form("txtAddress") 
   strCity = Request.Form("txtCity") 
   strState = Request.Form("txtState") 
   strZip = Request.Form("txtZip") 
   strShort = Request.Form("txtShort")



  intRC = fn_BL_InitialContactInsert(strCompany, strName, strEmail, strPhone, strFax, strAddress, strCity, strState, strZip, strShort)       
        
  Select Case intRC
    Case 0:  

       Response.Redirect PAGE_INITIALCONTACT & "?" & QSVAR & "1=" & ACTION_SUCCESS 

    Case Else:
       Response.Redirect "Error.asp?" & QSVAR & "1=" & intRC

   End Select
   
   strBody = "Company: " & strCompany
   strBody = strBody & Chr(13) & Chr(10) & "Name: " & strName
   strBody = strBody & Chr(13) & Chr(10) & "Email: " & strEmail
   strBody = strBody & Chr(13) & Chr(10) & "Phone: " & strPhone
   strBody = strBody & Chr(13) & Chr(10) & "Fax: " & strFax
   strBody = strBody & Chr(13) & Chr(10) & "Address: " & strAddress
   strBody = strBody & Chr(13) & Chr(10) & "City: " & strCity
   strBody = strBody & Chr(13) & Chr(10) & "State: " & strState
   strBody = strBody & Chr(13) & Chr(10) & "Zip: " & strZip
   strBody = strBody & Chr(13) & Chr(10) & "Short: " & strShort
   
   strFrom = "Blue Elephant<support@blueelephantwebdev.com>"
   strSubject = "Initial Contact"
   strTo = "chad.collins@blueelephantwebdev.com"
   intRC = fn_SendEmail(strTo, strFrom, strSubject, strBody)


End Sub    ' sub_ValidateContact




'Sub DisplayContactMain
'*********************************************
Sub DisplayContactMain(strError)
   Dim arryVar(0)
   arryVar(0) = ACTION_VALIDATE
%>
   <form action="<% = fn_CreateURL(True,PAGE_INITIALCONTACT,arryVar) %>" method="post" name="theForm">
   <span class="smalltitle">Contact Us toll free at <font color=red><b>(866) 863 - 4721</b></font> or submit the following form: </span><br><BR>
     <table name="main" width="90%" >
       <tr>
         <td>
             <font size="2"><font color=red><b>*</b></font> Please provide a brief description of your project or problem in the space below:</font>
             <textarea name="txtShort" cols=75 rows=3></textarea>
       	
       	<br>
       	<font size="2">Please provide the appropriate information below as it applies and<br> proceed by clicking "Send".&nbsp;&nbsp;<I>This information will remain confidential.</I></font>
       	<table name="company_main" width="100%" cellpadding=2>	

       	  <tr>
       	    <td colspan=1 align="right"><b><font>Business Name:</font></b></td><td colspan=3><input type="text" name="txtCompany" size=20></td>
       	  </tr>	
       
       	  <tr>
       	    <td align="right"><font color=red><b>*</b></font> <b><font>Contact Name:</font></b></td><td><input type="text" name="txtName"></td>
       	    <td align="right"><font color=red><b>*</b></font> <b><font>Contact Email:</font></b></td><td><input type="text" name="txtEmail"></td>
       	  </tr>
       	  <tr>
       	    <td align="right"><b><font>Contact Phone:</font></b></td><td><input type="text" name="txtPhone"></td>
       	    <td align="right"><b><font>Contact Fax:</font></b></td><td><input type="text" name="txtFax"></td>
       	  </tr>
       	  <tr>
       	    <td colspan=4>&nbsp;</td>
       	  </tr>
       
       	  <tr>
       	    <td align="right"><b><font>Business Street:</font></b></td><td><input type="text" name="txtAddress"></td>
       	    <td align="right"><b><font>City:</font></b></td><td><input type="text" name="txtCity"></td>
       	  </tr>	
       	  <tr>
       	    <td align="right"><b><font>State/Province:</font></b></td><td><input type="text" name="txtState"></td>
       	    <td align="right"><b><font>ZIP/Postal Code:</font></b></td><td><input type="text" name="txtZIP"></td>
       	  </tr>	
       	</table name="company_main">
     
         <br>
         <table width="100%">
           <tr align="center">
           <td colspan=2 align="center"><input type="button" value="Cancel" onClick="fctBack()" onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                 <input type="submit" value="Send"  onClick="return CheckEmpty2(document.theForm.txtShort.value, 'Please enter a brief description of your problem.') && CheckEmpty2(document.theForm.txtName.value, 'Please enter a Contact Name.') && CheckEmpty2(document.theForm.txtEmail.value, 'Please enter an email.') && isEmail2(document.theForm.txtEmail.value);" onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
                 <br><font size=-1 color=red><b>* indicates a required field</b></font>
            </td>
             <!-- <td align="right"><input type="button" value=" Cancel " onClick="fctBack()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="   Send   " ></td> -->
           </tr>
         </table>
         </td>
       </tr>
     </table name="main">
   </form>
<% 
End Sub           'DisplayContactMain



'Sub DisplayContactSuccess
'*********************************************
Sub DisplayContactSuccess()
%>
<table width="600">
  <tr>
    <td align="left">
      <span class="smalltitle">Contact Us</span><BR><br>
      <font>Your information was sent successfully !<br>Blue Elephant Web Development will contact you within 1 to 2 business days.</font> 
    </td>
  </tr>
</table>          
<% 
End Sub           'DisplayContactSuccess


' Call the Main Sub Routine
'*******************************
Call Main
%>

<script language="JavaScript">
function fctBack() {
	window.navigate("index.htm")
}
</script>
