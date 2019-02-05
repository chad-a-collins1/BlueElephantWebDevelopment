<!-- #include file="../DataLayer/DL_adm_getInvoice.asp"-->
<%
'*********************************************************************************************
function fn_BL_getInvoices(lngProjID)
'On Error Resume Next


	Dim lngRC
	lngRC = fn_DL_getInvoices(lngProjID)

	If lngRC <> 0 Then
	   fn_BL_getInvoices = lngRC
	   Exit Function
	End If

end function



'*********************************************************************************************
function fn_BL_getOutstandingInvoices(lngProjID, arryTmp)
'On Error Resume Next


	Dim lngRC
	lngRC = fn_DL_getOutstandingInvoices(lngProjID, arryTmp)

	If lngRC <> 0 Then
	   fn_BL_getOutstandingInvoices = lngRC
	   Exit Function
	End If

end function


'*********************************************************************************************
function fn_BL_getInvoice(lngProjID, arryTmp)
'On Error Resume Next

	fn_BL_getInvoice = -1

	Dim lngRC
	lngRC = fn_DL_getInvoice(lngProjID, arryTmp)

	If lngRC <> 0 Then
	   fn_BL_getInvoice = lngRC
	   Exit Function
	End If

	fn_BL_getInvoice = 0

end function

%>