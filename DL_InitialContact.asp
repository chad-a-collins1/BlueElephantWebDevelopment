<%

' Insert Contact Information
'***************************************************************************************************
Function fn_DL_InitialContactInsert(strCompany, strName, strEmail, strPhone, strFax, strAddress, strCity, strState, strZip, strShort)

'On Error Resume Next

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "INSERT INTO InitialContact"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "company,"
    strSQL = strSQL & " " & Chr(13) & "contactName,"
    strSQL = strSQL & " " & Chr(13) & "email,"
    strSQL = strSQL & " " & Chr(13) & "phone,"
    strSQL = strSQL & " " & Chr(13) & "fax,"
    strSQL = strSQL & " " & Chr(13) & "address,"
    strSQL = strSQL & " " & Chr(13) & "city,"
    strSQL = strSQL & " " & Chr(13) & "state,"
    strSQL = strSQL & " " & Chr(13) & "zip,"
    strSQL = strSQL & " " & Chr(13) & "summary"
    strSQL = strSQL & " " & Chr(13) & ")"
    strSQL = strSQL & " " & Chr(13) & "VALUES"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "'" & strCompany & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strName & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strEmail & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strPhone & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strFax & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strAddress & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strCity & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strState & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strZip & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strShort & "'"
    strSQL = strSQL & " " & Chr(13) & ")"
    
    'Response.Write "SQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
      
    'If Err.Number <> 0 Then
    '   fn_DL_InitialContactInsert = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_InitialContactInsert = 0

End Function




%>





















