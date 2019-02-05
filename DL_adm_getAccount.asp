<%
Function fn_DL_getAccountID()


	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	    Set rs = Server.CreateObject("ADODB.RecordSet")
	    strSQL = "SELECT * FROM Account ORDER BY account_id ASC"
	    rs.Open strSQL, dbconTmp, 3, 3

 If Not rs.EOF And Not rs.BOF Then
 
	rs.MoveLast
	tmpAID = rs.Fields("account_id")

	fn_DL_getAccountID = tmpAID

	rs.Close
	Set rs = Nothing

   	dbconTmp.Close
   	Set dbconTmp = Nothing

 Else
 
 	rs.Close
	Set rs = Nothing

   	dbconTmp.Close
   	Set dbconTmp = Nothing
 
 End If


End Function
%>