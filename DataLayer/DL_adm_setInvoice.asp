<%
Function fn_DL_admInvoiceInsert(billinginfo_id, project_id, invoicetype_code, invoicestatus_code, descrip, comment, balance, balance_forward, total_amount, invoice_date, due_date, outstanding_amount, insert_datetime, begin_date, end_date)

'On Error Resume Next

	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM Invoice"		
	rs.Open strSQL, dbconTmp, 3, 3


rs.AddNew    
rs.Fields("billinginfo_id") = billinginfo_id
rs.Fields("project_id") = project_id
rs.Fields("invoicetype_code") = invoicetype_code
rs.Fields("invoicestatus_code") = invoicestatus_code 
rs.Fields("descrip") = descrip 
rs.Fields("comment") = comment 
rs.Fields("balance") = balance 
rs.Fields("balance_forward") = balance_forward 
rs.Fields("total_amount") = total_amount
rs.Fields("invoice_date") = now() 'invoice_date 
If due_date <> "" Then
   rs.Fields("due_date") = CDate(due_date) 
End If
rs.Fields("outstanding_amount") = outstanding_amount 
'rs.Fields("insert_datetime") = now() 
rs.Fields("invoicestatus_code") = invoicestatus_code
If begin_date <> "" Then
   rs.Fields("begin_date") = CDate(begin_date) 
End If
If end_date <> "" Then
   rs.Fields("end_date") = CDate(end_date)
End If   
rs.Update



	    rs.Close
	    Set rs = Nothing  

	    dbconTmp.Close
	    Set dbconTmp = Nothing
       
     
	fn_DL_admInvoiceInsert = 1

End Function



'*************************************************************************************************************************
Function fn_DL_admInsertLineItem(lngInvoiceID, intQty, strComment, strRateCode, strLineItemTypeCode, blbDescrip)




End Function




'*************************************************************************************************************************
Function fn_adm_UpdateProjectAccruedHours(lngProjID, intQty)



End Function
















%>