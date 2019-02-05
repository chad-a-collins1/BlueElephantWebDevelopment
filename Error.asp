<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp"-->
<!--#include file="Utility/DBUtil.asp"-->
<%
'Sub Main
'*********************************************
Sub Main()

'On Error Resume Next

   'Call sub_SetUserAgent

   Dim strErr
   
   strErr = Request(QSVAR & "1")

   Call DisplayHeader(TITLE_ERROR, PIC_ERROR, LOGO_GRAPHIC)
   Call DisplayErrorMain (strErr)
   Call DisplayFooter

End Sub


'Sub DisplayErrorMain
'*********************************************
Sub DisplayErrorMain(strErr)
%>
<table width="600">
  <tr>
    <td align="left">
 <%
    If strErr = 1111 Then
        Response.Write "Session Timeout"
    Else
 %>   
      <font size=3 color="#000000"><b>Error</b></font><BR><br>
      <font><% = strErr %></font> 
<%
   End If
%>    
    </td>
  </tr>
</table>  
<% 
End Sub           'DisplayErrorMain


' Call Main Subroutine
'**********************************************************
Call Main

%>
