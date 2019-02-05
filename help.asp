<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="False"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!-- #include file="Utility/Util.asp" -->
<%

Const ACTION_TABLEOFCONTENTS = "ajkd_4837feh"


'Sub Main
'*********************************************
Sub Main()

'On Error Resume Next
  
Dim i  
Dim strAct
  
   strAct = Request(QSVAR & "1")

   Select Case strAct
            
      Case ACTION_MAIN: 
        Call DisplayHelpTableOfContents()
   
               
      Case ACTION_TABLEOFCONTENTS
		 Call DisplayHelpTableOfContents()
         
'      Case Else:
'        Call sub_HandleLogicError(ERR_INVALID_ACTION)
           
   End Select
   
   

End Sub   'Main
 
 

'Sub DisplayPortalMain
'*********************************************
Sub DisplayHelpTableOfContents()

Dim arryVar(1)
   

arryVar(0) = ACTION_VALIDATECHOOSE
arryVar(1) = fn_GetRandomAlphaNumeric(11,16)
%>
<table width="600">
  <tr>
    <td align="left">
   <form action="<% = fn_CreateURL(True,PAGE_PORTAL,arryVar) %>" method="post" name="theForm">

   <table> 
     <tr>
       <%      %>
     </tr>   

    <tr>
       <td colspan=2>
         <br><br>
         <table width=200>
          <tr>
            <td align=center>
 <% response.write "add contents tree here" %>

            <br><br>
            <input type="submit" value="Go"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      
           </td>
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
End Sub           'DisplayPortalMain





' Call the Main Sub Routine
'*******************************
Call Main


%>











































































































































