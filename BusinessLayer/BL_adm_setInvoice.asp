<!-- #include file="../DataLayer/DL_adm_setInvoice.asp"-->

<%
Function fn_BL_adm_setInvoice(billinginfo_id, project_id, invoicetype_code, invoicestatus_code, descrip, comment, balance, balance_forward, total_amount, invoice_date, due_date, outstanding_amount, insert_datetime, begin_date, end_date)

  Dim lngRC
  lngRC = fn_DL_admInvoiceInsert(billinginfo_id, project_id, invoicetype_code, invoicestatus_code, descrip, comment, balance, balance_forward, total_amount, invoice_date, due_date, outstanding_amount, insert_datetime, begin_date, end_date) 
  If lngRC <> 1 Then
     fn_BL_adm_setInvoice = lngRC
     Exit Function
  End If 

  fn_BL_adm_setInvoice = 1

End Function
%>