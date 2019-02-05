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
Sub Main


Dim strAction

strAction = Request("x1")

Select Case strAction

Case "":
     
     Call DisplayMain(Request("x2"))

Case "VAL"
     Call sub_ValidateLogin

Case Else
     Response.Redirect Error.asp

End Select


End Sub



Sub DisplayMain (strMsg)
%>
<html>
<head>
   <title>Admin Login</title>
</head>
<body>
<center>
<% 
  Select Case strMsg
     Case "1":
          Response.Write "Invalid UserName Or Password!!!"
     Case Else:
  End Select
%>
<br><br>
<form action="AdminLogin.asp?x1=VAL" method="POST">

UserName: <input type=text name=txtUser maxlength="30" size=30>
<br>Password: <input type=password name=txtPwd maxlength="15" size=30>
<br><input type=submit name=cmdSubmit value="Login">

</form>
</center>
</body>
</html>
<%
End Sub


Sub sub_ValidateLogin

Dim strUser, strPwd

strUser = Request("txtUser")
strPwd = Request("txtPwd")

Select Case strUser

Case "Daniel.Rigney":
      If strPwd = "1977" Then
          Session("blnAdminLoggedIn") = True
          Session("ConsultantId") = 2
      Else
          Response.Redirect "AdminLogin.asp?x2=1" 
      End If

Case "Chad.Collins"
      If strPwd = "2205" Then
          Session("blnAdminLoggedIn") = True
          Session("ConsultantId") = 1
      Else
          Response.Redirect "AdminLogin.asp?x2=1" 
      End If
      
Case Else:
     Response.Redirect "AdminLogin.asp?x2=1" 

End Select

Response.Redirect "AdminMain.asp"

End Sub

Call Main
%>