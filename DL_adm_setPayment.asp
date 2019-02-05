<%

function fn_DL_adm_setPayment(invoice_id, paymenttype_code, descrip, comment, amount, paymentdate)

'On Error Resume Next

	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM Payment"		
	rs.Open strSQL, dbconTmp, 3, 3


rs.AddNew    
rs.Fields("invoice_id") = invoice_id
rs.Fields("paymenttype_code") = paymenttype_code
rs.Fields("descrip") = descrip
rs.Fields("comment") = comment 
rs.Fields("amount") = amount
rs.Fields("payment_datetime") = paymentdate 
rs.Update



	    rs.Close
	    Set rs = Nothing  

	    dbconTmp.Close
	    Set dbconTmp = Nothing
       
     
	fn_DL_admInvoiceInsert = 1

end function





function fn_DL_adm_setInvoicePayment(invoice_id, payment_id, amount)


end function





fn_DL_adm_getPaymentID()


end function





%>