<%

Function fn_DL_adm_InitContactToAccount(lngCntctID, arryTmp)

	Dim dbconTmp
	Dim rs
	Dim strSQL
	   
	Set dbconTmp = Server.CreateObject("ADODB.Connection")
	dbconTmp.ConnectionString = fn_GetConnectionString
	dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM InitialContact WHERE initialcontact_id = " & lngCntctID	
	rs.Open strSQL, dbconTmp, 3, 3


	With rs
	  If Not .BOF Or .EOF Then
	      arryTmp = .GetRows
	  End If
	End WIth

	
	rs.Close
	Set rs = Nothing
	dbconTmp.Close
	Set dbconTmp = Nothing

End Function



'
'*********************************************************************************
Function fn_DL_initialAccountInsert(strAcctNum, strUID, strFakePwd, strRealPwd, lngAcctId)

    Dim rsTmp
    
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
    rsTmp.Open "Account", fn_GetConnectionString, adOpenDynamic, adLockOptimistic, adCmdTable
   
    rsTmp.AddNew    ' begin an insert  						
      
    rsTmp.Fields("account_num") = strAcctNum	
    rsTmp.Fields("accountstatus_code") = "ACT" 
    rsTmp.Fields("accounttype_code") = "HOURLY" 
    rsTmp.Fields("username") = strUID 
    rsTmp.Fields("account_password") = strFakePwd 
    rsTmp.Fields("RowId") = strRealPwd 
     
    rsTmp.Update     'commit the insert 
      
    lngAcctId = rsTmp.Fields("account_id") 
	
    rsTmp.Close
    Set rsTmp = Nothing   
    
    fn_DL_initialAccountInsert = 0  

End Function



Function fn_DL_admUpdateInitContact(lngInitCntcID, intConId)

	    Dim dbconTmp
	    Dim strSQL
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

    	    strSQL = "Update InitialContact set open_yn = 1, consultant_id = " & intConId & " WHERE initialcontact_id = " & lngInitCntcID

	    dbconTmp.Execute(strSQL)
	    dbconTmp.Close
	    Set dbconTmp = Nothing

End Function


Function fn_DL_admDeleteInitContact(lngInitCntcID)

	    Dim dbconTmp
	    Dim strSQL
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

    	    strSQL = "DELETE FROM InitialContact WHERE initialcontact_id = " & lngInitCntcID

	    dbconTmp.Execute(strSQL)
	    dbconTmp.Close
	    Set dbconTmp = Nothing

End Function


















%>