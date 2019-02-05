<%



'**********************************************************************************************************
function fn_DL_getInvoices(lngProjID)


	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	    Dim intInvoiceCount
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	    Set rs = Server.CreateObject("ADODB.RecordSet")
	    strSQL = "SELECT project_id, invoice_id FROM Invoice WHERE project_id = " & lngProjID	
	    rs.Open strSQL, dbconTmp, 3, 3

	    With rs
		If Not .EOF and Not .BOF Then
		   intInvoiceCount = .RecordCount
		Else
		   intInvoiceCount = 0
		End If
		.Close
	    End With
	    

	    Set rs = Nothing
	    dbconTmp.Close
	    Set dbconTmp = Nothing

	   fn_DL_getInvoices = intInvoiceCount

end function





'**********************************************************************************************************
function fn_DL_getOutstandingInvoices(lngProjID, arryTmp)


	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	    Set rs = Server.CreateObject("ADODB.RecordSet")
	    strSQL = "SELECT * FROM Invoice WHERE project_id = " & lngProjID & " and outstanding_amount <> 0 ORDER BY invoice_id ASC"	
	    
	    'response.Write "strSQL = " & strSQL
	    'response.End
	    
	    rs.Open strSQL, dbconTmp, 3, 3

	    With rs
		If Not .EOF and Not .BOF Then
		   arryTmp = .GetRows
		Else
'		   arryTmp(0,0) = "No Invoices"
		End If
		.Close
	    End With
	    

	    Set rs = Nothing
	    dbconTmp.Close
	    Set dbconTmp = Nothing

	   fn_DL_getOutstandingInvoices = 0

end function







'********************************************************************************************************
function fn_DL_getInvoice(lngProjID, arryTmp)
	    
	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	    Set rs = Server.CreateObject("ADODB.RecordSet")
	    strSQL = "SELECT * FROM Invoice WHERE project_id = " & lngProjID & " ORDER BY insert_datetime DESC"
	    rs.Open strSQL, dbconTmp, 3, 3
	    
	    'Grab the latest record, this will always be the current invoice 
	    'rs.MoveLast


	    With rs
		If Not .EOF and Not .BOF Then
		   arryTmp = .GetRows
		End If
		.Close
	    End With
	    

	    Set rs = Nothing
	    dbconTmp.Close
	    Set dbconTmp = Nothing

	    fn_DL_getInvoice = 0    
	    

end function

%>








