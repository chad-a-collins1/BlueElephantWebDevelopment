<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp"-->
<!--#include file="Utility/DBUtil.asp"-->
<%
  If Session("blnAdminLoggedIn") <> True Then
     Response.Redirect "Error.asp?x1=1111"
  End If

Dim pID
Dim lngRC, strMsg, strFrom
Dim due_date, invoice_date, invoice_id

strMsg=""
strFrom = Request("txtFrom")

pID = Request.QueryString("pID")

If strFrom = "printsendinvoice" Then

    invoice_id = Request.Form("lstInvoice")
    due_date = Request.Form("DUE_DATE")
    invoice_date = Request.Form("INVOICE_DATE")

    lngRC = fn_BL_PrintSendInvoice(pID, invoice_id, invoice_date, due_date)
    If lngRC = 0 Then
       Session("sess_strMsg") = "Invoice Print/Send Suckdickfully Bizzatch!"   
    Else
       Session("sess_strMsg") =  "Problem Print/Send Invoice!" & lngRC
    End If
   
    Response.Redirect "admProject_PrintSendInvoice.asp?txtFrom=verify&pID=" & pID
    
ElseIf strFrom = "verify" Then
    strMsg = Session("sess_strMsg") 
End If

%>
<html>
<head>

</head>

<body bgcolor="#ededed">

<%
Response.Write "<br><br><center>" & strMsg & "</center>"
%>

<form method="post" action="admProject_PrintSendInvoice.asp?pID=<%= pID %>" name="theForm">
<br><br><br><br><br><br>
<input type="hidden" name="txtFrom" value="printsendinvoice">
<center>
<table>

<%
    Dim arryInv
     lngRC = fn_BL_GetProjectBillWaitingInvoices(pId,arryInv)
    If lngRC <> 0 Then
         Response.Redirect "Error.asp?x1=" & lngRC
    End If
   
    If Not IsArray(arryInv) Then
       Response.Write "NO INVOICES To Print/Send for this PROJECT!"
       Response.End
    End  If
    
'invoice_id, descrip, total_amount, outstanding_amount, insert_datetime
%>	    
<tr><td colspan=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                          [inv_id][descrip][ins_date][$tot_amt][$amt_owed]
<br><b>Choose Invoice:</b><select name="lstInvoice">
<option value="none">Choose Invoice
<%
   Dim intI, intRecLst
   intRecLst = UBound(arryInv, 2)
   For intI = 0 To intRecLst
      Response.Write "<option value=" & arryInv(0,intI) & ">[ " & arryInv(0,intI) & " ][ " & arryInv(1,intI) & " ][ " & Left(arryInv(4,intI),InStr(arryInv(4,intI) - 1," ")) & " ][ $" & arryInv(2,intI) & " ][ $" & arryInv(3,intI) & " ]"  	
   Next
 	 	
%>
</select>
</td></tr>

	<tr>
	<td><b>INVOICE DATE:</b></td>
	<td><input type="text" value="<% = Date() %>" name="INVOICE_DATE"></td>
	</tr>

	<tr>
	<td><b>DUE DATE:</b></td>
	<td><input type="text" name="DUE_DATE"></td>
	</tr>

	<tr>
	<td colspan=2><center><input type="submit" value="Send Out Invoice"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>
<%

' Get Bill Waiting Invoices
'**********************************************************************************************************
Function fn_BL_GetProjectBillWaitingInvoices(lngProjID, arryInv)


	    Dim dbconTmp
	    Dim rsTmp
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    Set rsTmp = Server.CreateObject("ADODB.RecordSet")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	    strSQL = ""
	    strSQL = strSQL & "SELECT invoice_id, descrip, total_amount, outstanding_amount, insert_datetime"
	    strSQL = strSQL & " FROM Invoice"
	    strSQL = strSQL & " WHERE project_id = " & lngProjID
	    strSQL = strSQL & " and invoicestatus_code = 'BW' order by insert_datetime desc"

            Set rsTmp = dbconTmp.Execute(strSQL)
 
            If Not rsTmp.EOF and Not rsTmp.BOF Then
		arryInv = rsTmp.GetRows
	    End If
	    
	    rsTmp.Close
	    Set rsTmp = Nothing
	    dbconTmp.Close
	    Set dbconTmp = Nothing

	   fn_BL_GetProjectBillWaitingInvoices = 0

End Function



'   UpdateInvoiceToBillOut
'*********************************************************************************
Function fn_BL_PrintSendInvoice(project_id, invoice_id, invoice_date, due_date)

'On Error Resume Next

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL, inv_amount
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

 ' Response.Write "strSQL = " & strSQL & "<br>"
    strSQL = "select total_amount from Invoice where invoice_id = " & invoice_id
    Set rsTmp = dbconTmp.Execute (strSQL)
    
    inv_amount = rsTmp.Fields("total_amount")
    
    rsTmp.Close
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE Invoice"
    strSQL = strSQL & " " & Chr(13) & "set invoicestatus_code = 'BO',"
    strSQL = strSQL & " " & Chr(13) & "invoice_date = '" & invoice_date & "',"
    strSQL = strSQL & " " & Chr(13) & "due_date = '" & due_date & "'"
    strSQL = strSQL & " " & Chr(13) & "where invoice_id = " & invoice_id 

    'Response.Write "strSQL = " & strSQL & "<br>"
    dbconTmp.Execute (strSQL)
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE Project"
    strSQL = strSQL & " " & Chr(13) & "set project_balance = project_balance + " & inv_amount
    strSQL = strSQL & " " & Chr(13) & "where project_id = " & project_id

    'Response.Write "strSQL = " & strSQL & "<br>"
    'Response.End

    dbconTmp.Execute (strSQL)
        
    dbconTmp.Close
    Set dbconTmp = Nothing
       
     
   fn_BL_PrintSendInvoice = 0

End Function
%>