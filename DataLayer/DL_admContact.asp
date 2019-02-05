<%

' Insert Contact Information
'***************************************************************************************************
Function fn_DL_ContactInsert(strAccountID, strFirst, strLast, strPhone1, strPhone2, strEmail, strAddress, strCity, strPostalCode, strStateDescrip)

'On Error Resume Next

	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 



    strSQL = "INSERT INTO ContactInfo"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "account_id,"
    strSQL = strSQL & " " & Chr(13) & "first_name,"
    strSQL = strSQL & " " & Chr(13) & "last_name,"
    strSQL = strSQL & " " & Chr(13) & "phone1,"
    strSQL = strSQL & " " & Chr(13) & "phone2,"
    strSQL = strSQL & " " & Chr(13) & "email,"
    strSQL = strSQL & " " & Chr(13) & "address,"
    strSQL = strSQL & " " & Chr(13) & "city,"
    strSQL = strSQL & " " & Chr(13) & "postal_code,"
    strSQL = strSQL & " " & Chr(13) & "state_descrip"
    strSQL = strSQL & " " & Chr(13) & ")"
    strSQL = strSQL & " " & Chr(13) & "VALUES"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "'" & strAccountID & "',"
If  strFirst = "" Then
    strSQL = strSQL & " " & Chr(13) & "null,"
Else
    strSQL = strSQL & " " & Chr(13) & "'" & strFirst & "',"
End If
If  strLast = "" Then
    strSQL = strSQL & " " & Chr(13) & "null,"
Else
    strSQL = strSQL & " " & Chr(13) & "'" & strLast & "',"
End If
    strSQL = strSQL & " " & Chr(13) & "'" & strPhone1 & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strPhone2 & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strEmail & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strAddress & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strCity & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strPostalCode & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strStateDescrip & "'"
    strSQL = strSQL & " " & Chr(13) & ")"
    
    'Response.Write "SQL = " & strSQL
    'Response.End
    

    dbconTmp.Execute(strSQL)
    dbconTmp.Close
    Set dbconTmp = Nothing
      
    If Err.Number <> 0 Then
       fn_DL_ContactInsert = 0
       Exit Function
    End If
    
    fn_DL_ContactInsert = 1

End Function




%>





















