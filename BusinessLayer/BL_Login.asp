<!--#include file="../DataLayer/DL_Login.asp"-->
<!--#include file="../Utility/Incrypt.asp"-->
<!--#include file="../Utility/Random.asp"-->
<%


' Validate the password, If the password is valid then return the Acct DB id
'***************************************************************************************************
Function fn_BL_ValidateLogin(ByVal strInputUN, ByVal strInputPwd, lngAcctId, strAcctStatus, strAcctType, strFN, strLN)

   fn_BL_ValidateLogin = -1
   
   Dim lngRC
   Dim strPwd
   
   strPwd = fn_IncryptString(strInputPwd)

   lngRC = fn_DL_ValidateLogin(strInputUN, strPwd, lngAcctId, strAcctStatus, strAcctType, strFN, strLN)
       
   If lngRC <> 0 Then
      fn_BL_ValidateLogin = lngRC
      Exit Function
   End If
   
   fn_BL_ValidateLogin = 0

End Function


' Create a Log Out record in the DB
'***************************************************************************************************
Function fn_BL_DBLogOut(ByVal intAID, ByVal strSID)

   fn_BL_DBLogOut = 0
End Function  

' Create a LogIn record in the DB
'***************************************************************************************************
Function fn_BL_DBLogIn(ByVal intAID, ByVal strSID)

   fn_BL_DBLogIn = 0
End Function  




' Validate Email address for password retrieval
'***************************************************************************************************
Function fn_BL_ValidateEmail(ByVal strInputEmail, strPwd, strUN)  

   Dim lngRC

   fn_BL_ValidateEmail = -1
   
   lngRC = fn_DL_ValidateEmail(strInputEmail, strPwd, strUN)
       
   If lngRC <> 0 Then
      fn_BL_ValidateEmail = 1
      Exit Function
   End If
   
   fn_BL_ValidateEmail = 0


End Function

%>





















