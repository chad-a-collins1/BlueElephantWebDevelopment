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
Dim lngRC, billinginfo_id, project_id, invoicetype_code, invoicestatus_code, descrip, comment, balance, strMsg, strFrom
Dim balance_forward, total_amount, due_date, outstanding_amount, insert_datetime, begin_date, end_date

strMsg=""
strFrom = Request("txtFrom")

pID = Request.QueryString("pID")

If strFrom = "createinvoice" Then

    project_id = pID
    invoicetype_code = Request.Form("TYPE")
    invoicestatus_code = Request.Form("STATUS")
    descrip = Request.Form("DESCRIP")
    comment = Request.Form("COMMENT")
   ' due_date = Request.Form("DUE_DATE")
   ' begin_date = Request.Form("BEGIN_DATE")
   ' end_date= Request.Form("END_DATE")

    lngRC = fn_BL_GenerateInvoice(project_id, invoicetype_code, invoicestatus_code, descrip, comment, due_date, begin_date, end_date)
    If lngRC = 0 Then
       Session("sess_strMsg") = "Invoice Generated Suckdickfully Bizzatch!"   
    ElseIf lngRC = "111" Then
       Session("sess_strMsg") = "No Line Items or Overdue Outstanding balances exist at this time."
    Else
       Session("sess_strMsg") =  "Problem Creating Invoice!" & lngRC
    End If
   
    Response.Redirect "admProject_CreateInvoice.asp?txtFrom=verify&pID=" & pID
    
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

<form method="post" action="admProject_CreateInvoice.asp?pID=<%= pID %>" name="theForm">
<br><br><br><br><br><br>
<input type="hidden" name="txtFrom" value="createinvoice">
<center>
<table>
	<tr>
	<td><B>INVOICE TYPE:</b></td>
	<td><SELECT name="TYPE">
	    <!-- <OPTION Value="BUF">Bill Up Front</OPTION> -->
	     <OPTION Value="MO" selected>Mail Out</OPTION>
	   </SELECT></td>
	</tr>

	<tr>
	<td><b>INVOICE STATUS:</b></td>
	<td><SELECT name="STATUS">
	            <!--<OPTION Value="BC">bill cancelled</OPTION>
		    <OPTION Value="BD">bill dispute</OPTION>
		    <OPTION Value="BO">bill out</OPTION> 
		    <OPTION Value="BP">bill paid</OPTION> -->
		    <OPTION Value="BW" selected>bill waiting</OPTION>
	    </SELECT></td>
	</tr>


	<tr>
	<td colspan=2><b>DESCRIPTION:</b><br><TEXTAREA cols=30 rows=5 name="DESCRIP"></TEXTAREA></td>
	</tr>


	<tr>
	<td colspan=2><b>COMMENT:</b><br><TEXTAREA cols=30 rows=5 name="COMMENT"></TEXTAREA></td>
	</tr>

<!--
	<tr>
	<td><b>DUE DATE:</b></td>
	<td><input type="text" name="DUE_DATE"></td>
	</tr>


	<tr>
	<td><b>START DATE:</b></td>
	<td><input type="text" name="BEGIN_DATE"></td>
	</tr>
	
	<tr>
	<td><b>END DATE:</b></td>
	<td><input type="text" name="END_DATE"></td>
	</tr>
-->

	<tr>
	<td colspan=2><center><input type="submit" value="Generate Invoice"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>
<%

Function fn_BL_GenerateInvoice(project_id, invoicetype_code, invoicestatus_code, descrip, comment, due_date, begin_date, end_date)

On Error Resume Next

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strStatus
    Dim curBalance, curBalanceForward, curTotalAmount, i, arryItems, blnNoBal, blnNoBalForward, lngInvoiceId    

    strStatus = "PEND"
    blnNoBal = False
    blnNoBalForward = False
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "select lineitem_id, amount"
    strSQL = strSQL & " " & Chr(13) & "from LineItem"
    strSQL = strSQL & " " & Chr(13) & "where project_id = " & project_id 
    strSQL = strSQL & " " & Chr(13) & "and lineitemstatus_code = 'PEND'"

    Set rsTmp = dbconTmp.Execute (strSQL)
    curTotalAmount = 0
    curBalance = 0
    curBalanceForward = 0
    If Not rsTmp.EOF Then
       arryItems = rsTmp.GetRows
       rsTmp.Close
       Set rsTmp = Nothing
       For i = 0 to Ubound(arryItems, 2)
          curBalance = curBalance + CCur(arryItems(1,i))
       Next
    Else
       ' No line Items at this time
       blnNoBal = True
       rsTmp.Close
    End If
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "select sum(outstanding_amount) as bal_forward"
    strSQL = strSQL & " " & Chr(13) & "from Invoice"
    strSQL = strSQL & " " & Chr(13) & "where project_id = " & project_id 
    strSQL = strSQL & " " & Chr(13) & "and invoicestatus_code = 'BO'"
    strSQL = strSQL & " " & Chr(13) & "and outstanding_amount > 0"
    strSQL = strSQL & " " & Chr(13) & "and due_date < Date()"

    Set rsTmp = dbconTmp.Execute (strSQL)
    
    If Not rsTmp.EOF Then
       If Not IsNull(rsTmp.Fields("bal_forward")) and rsTmp.Fields("bal_forward") <> "" Then
          curBalanceForward = rsTmp.Fields("bal_forward")
       Else
           blnNoBalForward = True  
       End If
       rsTmp.Close
    Else
       ' No Line Items
       rsTmp.Close
       blnNoBalForward = True
    End If
    
    'Response.Write "blnNoBal = " & blnNoBal & "<br>"
    'Response.Write "blnNoBalForward = " & blnNoBalForward & "<br>"
    If blnNoBalForward And blnNoBal Then
       fn_BL_GenerateInvoice = 111
       Exit Function
    End If
    
    'Response.Write "curBalanceForward = " & curBalanceForward & "<br>"
    'Response.Write "curBalance = " & curBalance & "<br>"
    'Response.End
    curTotalAmount = curBalance + curBalanceForward
    
    dbconTmp.BeginTrans
    
    strSQL = "select * from Invoice where project_id = " & project_id

    rsTmp.Open strSQL, dbconTmp, adOpenDynamic, adLockOptimistic, adCmdText 
    				
    rsTmp.AddNew    ' begin an insert  						
  
    rsTmp.Fields("project_id") = project_id	
    rsTmp.Fields("invoicetype_code") = invoicetype_code 
    rsTmp.Fields("invoicestatus_code") = invoicestatus_code 
    rsTmp.Fields("descrip") = descrip 
    rsTmp.Fields("comment") = comment 
    rsTmp.Fields("balance") = curBalance 
    rsTmp.Fields("balance_forward") = curBalanceForward 
    rsTmp.Fields("total_amount") = curTotalAmount
    rsTmp.Fields("outstanding_amount") = curTotalAmount
    If due_date <> "" Then
       rsTmp.Fields("due_date") = due_date
    End If 
    If begin_date <> "" Then
       rsTmp.Fields("begin_date") = begin_date
    End If
    If end_date <> "" Then
       rsTmp.Fields("end_date") = end_date
    End If
      
    rsTmp.Update     'commit the insert
    lngInvoiceId = rsTmp.Fields("invoice_id") 
      
    rsTmp.Close

    If Not blnNoBal Then
       For i = 0 to Ubound(arryItems, 2)
          strSQL = "insert into InvoiceLineItem (invoice_id, lineitem_id) values(" & lngInvoiceId & ", " & arryItems(0,i) & ")"
          dbconTmp.Execute (strSQL)
          strSQL = "update LineItem set lineitemstatus_code = 'INVOICED' where lineitem_id = " & arryItems(0,i)
          dbconTmp.Execute (strSQL)
       Next
       Erase arryItems
    End If
     
    If Err.Number <> 0 Then
       fn_BL_GenerateInvoice = Err.Number
       dbconTmp.RollbackTrans
       dbconTmp.Close
       Set rsTmp = Nothing
       Set dbconTmp = Nothing
       Exit Function 
    Else   
       dbconTmp.CommitTrans
    End If
    
    dbconTmp.Close
    Set rsTmp = Nothing
    Set dbconTmp = Nothing
       
     
   fn_BL_GenerateInvoice = 0

End Function
%>