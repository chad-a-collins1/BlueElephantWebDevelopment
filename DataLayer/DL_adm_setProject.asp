<%
Function fn_DL_admProjectInsert(AccountID, ProjectTypeCode, StatusCode, intConId)

	Dim rs
	Dim strSQL

	Set rs = Server.CreateObject("ADODB.RecordSet")	
	rs.Open "Project", fn_GetConnectionString, adOpenDynamic, adLockOptimistic, adCmdTable

	rs.AddNew    
	rs.Fields("account_id") = AccountID
	rs.Fields("projecttype_code") = ProjectTypeCode
	rs.Fields("status_code") = StatusCode
	rs.Fields("project_balance") = 0
	rs.Fields("project_name") = "New Project"
	rs.Fields("consultant_id") = intConId
	rs.Update


'response.write "<BR>-->" & ProjectTypeCode
'response.write "<BR>-->" & StatusCode


	rs.Close
	Set rs = Nothing  
            
	fn_DL_admProjectInsert = 0


End Function
%>