<%

' Validate the password, If the password is valid then return the Acct DB id
'***************************************************************************************************
Function fn_DL_ValidateLogin(ByVal strUN, ByVal strPwd, lngAcctId, strAcctStatus, strAcctType, strFN, strLN)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "a.account_id,"
    strSQL = strSQL & " " & Chr(13) & "a.RowID,"
    strSQL = strSQL & " " & Chr(13) & "a.accountstatus_code," 
    strSQL = strSQL & " " & Chr(13) & "a.accounttype_code,"
    strSQL = strSQL & " " & Chr(13) & "ci.first_name," 
    strSQL = strSQL & " " & Chr(13) & "ci.last_name"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "Account a," 
    strSQL = strSQL & " " & Chr(13) & "ContactInfo ci" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "a.username = '" & strUN & "'"
    strSQL = strSQL & " " & Chr(13) & "and ci.account_id = a.account_id"
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    
    Set rsTmp = dbconTmp.Execute(strSQL)
    
    With rsTmp
       If Not .EOF and Not.BOF Then
          strDBPwd = .Fields("RowID")
          lngAcctId = .Fields("account_id")
          strAcctStatus = .Fields("accountstatus_code")
          strAcctType = .Fields("accounttype_code")
          strFN = .Fields("first_name")
          strLN = .Fields("last_name")   
       End If
       .Close
    End With  'rsTmp
    
    Set rsTmp = Nothing
    dbconTmp.Close
    Set dbconTmp = Nothing
     
    If strDBPwd = strPwd Then 
       fn_DL_ValidateLogin = 0
    Else
       If strDBPwd = "" Then
          fn_DL_ValidateLogin = ERR_INVALID_UID
          Exit Function
       Else
          fn_DL_ValidateLogin = ERR_INVALID_PWD
          Exit Function
       End If
    End If
      
    'If Err.Number <> 0 Then
    '   fn_DL_ValidateLogin = Err.Number
    '   Exit Function
    'End If

End Function




%>





















