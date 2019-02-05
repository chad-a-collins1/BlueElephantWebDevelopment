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
<!--#include file="Utility/Incrypt.asp"-->
<%
  If Session("blnAdminLoggedIn") <> True Then
     Response.Redirect "Error.asp?x1=1111"
  End If

Dim aID, strMsg, strFrom
Dim lngRC, strUserName, strNewPass1, strNewPass2

strMsg=""
strFrom = Request("txtFrom")

aID = Request.QueryString("aID")

If strFrom = "changepass" Then
  Session("sess_strMsg") = ""
    aID = aID
    strUserName = Request.Form("txtUserName")
    strNewPass1 = Request.Form("txtNewPass1")
    strNewPass2 = Request.Form("txtNewPass2")
    
    'Response.Write "username = " & strUserName
    'Response.End
    
    If strUserName <> "" Then
    
        If strNewPass1 = strNewPass2 Then
    
           lngRC = fn_BL_EditUserNameAndPasword(aID, strUserName, strNewPass1)
           If lngRC = 0 Then
              
              Session("sess_strMsg") = "UserName/Password Changed Successfully!<br>UserName Now = """ & strUserName & """"    
           Else
              Session("sess_strMsg") =  "Problem Editing UserName and Password!" & lngRC
           End If
        
        Else
           Session("sess_strMsg") =  "Passwords do not match"
        End If 
    
    Else
       Session("sess_strMsg") =  "Please Enter a username"
       
    End If
     
    Response.Redirect "admProject_ChangePassword.asp?txtFrom=verify&aID=" & aID
    
ElseIf strFrom = "verify" Then
    strMsg = Session("sess_strMsg") 
    lngRC = fn_BL_GetUserName(aID, strUserName)
    If lngRC <> 0 Then
       strMsg = "problem getting username"
    End If   
Else
    Session("sess_strMsg") = ""
    lngRC = fn_BL_GetUserName(aID, strUserName)
    If lngRC <> 0 Then
       strMsg = "problem getting username"
    End If   
End If

%>
<html>
<head>

</head>

<body bgcolor="#ededed">

<%
Response.Write "<br><br><center>" & strMsg & "</center>"
%>

<form method="post" action="admProject_ChangePassword.asp?aID=<%= aID %>" name="theForm">
<br><br><br><br><br><br>
<input type="hidden" name="txtFrom" value="changepass">
<center>
<table>

	<tr>
	<td><b>UserName:</b></td>
	<td><input type="text" name="txtUserName" value="<% = strUserName %>"></td>
	</tr>

	<tr>
	<td><b>New Password:</b></td>
	<td><input type="text" name="txtNewPass1"><br><font size="-1">Leave blank to NOT change</font></td>
	</tr>
	
	<tr>
	<td><b>Confirm:</b></td>
	<td><input type="text" name="txtNewPass2"></td>
	</tr>

	<tr>
	<td colspan=2><center><input type="submit" value="Update UserName/Password"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>
<%


' Edit UserName and password 
'*******************************************************************************
Function fn_BL_EditUserNameAndPasword(aID, strUserName, strNewPass1)

 On Error Resume Next

    Dim dbconTmp
    Dim strSQL   
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE Account"
    strSQL = strSQL & " " & Chr(13) & "set username = '" & strUserName & "'"
    If strNewPass1 <> "" Then
       strSQL = strSQL & " " & Chr(13) & ", RowId = '" & fn_IncryptString(strNewPass1) & "'"
    End If
    strSQL = strSQL & " " & Chr(13) & "where account_id = " & aID 

    dbconTmp.Execute (strSQL)

     
    If Err.Number <> 0 Then
       fn_BL_EditUserNameAndPasword = Err.Number
       dbconTmp.RollbackTrans
       dbconTmp.Close 
       Set dbconTmp = Nothing
       Exit Function 
    End If
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
     
   fn_BL_EditUserNameAndPasword = 0

End Function 'fn_BL_EditUserNameAndPasword 



' Get User Name for an Account
'*******************************************************************************
Function fn_BL_GetUserName(aID, strUserName)

On Error Resume Next

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL   
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = "SELECT username from Account where account_id = " & aID

    Set rsTmp = dbconTmp.Execute (strSQL)
    
    strUserName = rsTmp.Fields("username")
     
      
    If Err.Number <> 0 Then
       fn_BL_GetUserName = Err.Number
       dbconTmp.Close
       Set rsTmp = Nothing
       Set dbconTmp = Nothing
       Exit Function 
    End If
    
    dbconTmp.Close
    Set rsTmp = Nothing
    Set dbconTmp = Nothing   
     
    fn_BL_GetUserName = 0

End Function 'fn_BL_GetUserName

%>