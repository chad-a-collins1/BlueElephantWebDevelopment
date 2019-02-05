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
<html>
<head>

</head>

<body bgcolor="#ededed">

<%
  If Session("blnAdminLoggedIn") <> True Then
     Response.Redirect "Error.asp?x1=1111"
  End If
  
Dim lngProjID
Dim arryTmp
Dim lngRC
Dim lngInvcNmbr
Dim intRecLst
Dim intI
Dim strFrom, strMsg
Dim strPayType, strName, strIdNum1, strIdNum2, strExpir, curAmount, strDescrip, strComment, dtPayDate

lngProjID = Request.QueryString("pID")
    
strFrom = Request("txtFrom")

If strFrom = "pay" Then

   strPayType = Request.Form("lstPayType")
   strName = Request.Form("txtName")
   strIdNUm1 = Request.Form("txtIDNum1")
   strIdNum2 = Request.Form("txtIDNUm2")
   strExpir = Request.Form("txtExpir")
   curAmount = Request.Form("txtAmount")
   dtPayDate = Request.Form("txtPayDate")
   strDescrip = Request.Form("txtDescrip")
   strComment = Request.Form("txtComment")

 
    lngRC = fn_BL_AddPayment(lngProjID, strPayType, strName, strIdNum1, strIdNum2, strExpir, curAmount, strDescrip, strComment, dtPayDate )
    If lngRC <> 0 Then
       Call sub_HandleLogicError(lngRC)
    End If
    
    Session("sess_strMsg") = "Payment Made Successfully!!"
    Response.Redirect "admProject_EnterPayment.asp?txtFrom=verify&pID=" & lngProjID

ElseIf strFrom = "verify" Then
    strMsg = Session("sess_strMsg") 
End If

%>

<form method="post" action="admProject_EnterPayment.asp?pID=<%= lngProjID %>" name="theForm">
<input type="hidden" name="txtFrom" value="pay">
<br>
<center>
<table>

	<tr><td><% = strMsg %></td></tr>

	<tr><td>&nbsp;</td></tr>
	
	<tr>
	<td><B>Payment Type: </b>
	    <select name="lstPayType">
	         <option value="CCAMEX">credit card - American Express
	         <option value="CCDISC">credit card - Discover Card
	         <option value="CCMC">credit card - MasterCard
	         <option value="CCVISA">credit card - Visa
	         <option value="CHECK" selected>check or money order
	         <option value="PAYPAL">Pay Pal
	    </select>      
	</td>
	<td></td>
	</tr>

	<tr>
	<td><B>Name on Card / Name on Check: </b><input type="text" name="txtName" value="" size="30" maxlength="60"></td>
	<td></td>
	</tr>
	<tr>
	<td><B>Credit Card Num / Check Acct Num: </b><input type="text" name="txtIDNum1" value="" size="30" maxlength="50"></td>
	<td></td>
	</tr>	
	<tr>
	<td><B>CC Security Digits / Check Num: </b><input type="text" name="txtIDNum2" value="" size="30" maxlength="50"></td>
	<td></td>
	</tr>
	<tr>
	<td><B>Expir Date (MM/YYYY):</b><input type="text" name="txtExpir" value="" size="15" maxlength="7"></td>
	<td></td>
	</tr>	
	
	<tr>
	<td><B>Payment Amount:</b><input type="text" name="txtAmount" size="15" maxlength="22"></td>
	<td></td>
	</tr>

	<tr>
	<td><B>Descrip:</b><input type="text" name="txtDescrip" value="" size="50" maxlength="50"></td>
	<td></td>
	</tr>
	<tr>
	<td><B>Comment:</b><input type="text" name="txtComment" value="" size="100" maxlength="255"></td>
	<td></td>
	</tr>
		
	<tr>
	<td><B>Payment Date:</b><input type="text" name="txtPayDate" value="<%= Date() %>" size="15" maxlength="10"></td>
	<td></td>
	</tr>
	<tr><td>&nbsp;</td></tr>

	<tr>
	<td colspan=2><center><input type="submit" value="Enter"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>

<%
Function fn_BL_AddPayment(lngProjID, strPayType, strName, strIdNum1, strIdNum2, strExpir, curAmount, strDescrip, strComment, dtPayDate )

   On Error Resume Next

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL, lngPayId, i, arryInvoice, curCreditAmount, curRemainingAmount, curAppliedAmount 
    Dim curOldOutAmount, curNewOutAmount, intUB
    
    curAmount = CCur(curAmount)
    
    ' Open DB Connection
    '*******************************************************
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    dbconTmp.BeginTrans   'Begin a transaction
    
    ' Insert Payment
    '*********************************************************************
    strSQL = "select * from Payment where project_id = " & lngProjID
    rsTmp.Open strSQL, dbconTmp, adOpenDynamic, adLockOptimistic, adCmdText 
    rsTmp.AddNew    ' begin an insert  						
    rsTmp.Fields("project_id") = lngProjID	
    rsTmp.Fields("paymenttype_code") = strPayType 
    rsTmp.Fields("paymentstatus_code") = "PEND" 
    If strName <> "" Then
       rsTmp.Fields("idname") = strName 
    End If
    If strIdNum1 <> "" Then
       rsTmp.Fields("idnum1") = strIdNum1 
    End If
    If strIdNum2 <> "" Then
       rsTmp.Fields("idnum2") = strIdNum2 
    End If
    If strExpir <> "" Then
       rsTmp.Fields("idexpiredate") = strExpir 
    End If
    If strDescrip <> "" Then
       rsTmp.Fields("descrip") = strDescrip 
    End If
    If strComment <> "" Then
       rsTmp.Fields("comment") = strComment
    End If
    rsTmp.Fields("amount") = curAmount
    rsTmp.Fields("payment_datetime") = dtPayDate
    rsTmp.Update     'commit the insert
    lngPayId = rsTmp.Fields("payment_id") 
    rsTmp.Close

    ' Get any invoices that have been sent out and have outstanding amounts
    '******************************************************************************
    strSQL = "select invoice_id, outstanding_amount"
    strSQL = strSQL & " from Invoice"
    strSQL = strSQL & " where project_id = " & lngProjID 
    strSQL = strSQL & " and outstanding_amount > 0"
    strSQL = strSQL & " and invoicestatus_code = 'BO'"
    strSQL = strSQL & " ORDER BY invoice_date"
    
    'Response.Write "strSQL = " & strSQL & "<BR>"
    Set rsTmp = dbconTmp.Execute(strSQL)
    
    If Not rsTmp.EOF Then  'If there are any invoices then apply payments
       
       'Response.Write "If<br>"

             
       arryInvoice = rsTmp.GetRows
       rsTmp.Close
       curRemainingAmount = curAmount
       intUB = Ubound(arryInvoice, 2)
       i = 0
       While i <= intUB and curRemainingAmount > 0
       
          curOldOutAmount = CCur(arryInvoice(1,i))
          If curRemainingAmount > curOldOutAmount Then
             'Response.Write "If1<br>"
             'Response.Write "curRemainingAmount = " & curRemainingAmount & "<br>"
             'Response.Write "curOldOutAmount = " & curOldOutAmount & "<br><br>"
             curNewOutAmount = CCur(0)
             curAppliedAmount = curOldOutAmount
             curRemainingAmount = curRemainingAmount - curOldOutAmount
          Else
             'Response.Write "If2<br>"
             'Response.Write "curRemainingAmount = " & curRemainingAmount & "<br>"
             'Response.Write "curOldOutAmount = " & curOldOutAmount & "<br><br>"
             curNewOutAmount = curOldOutAmount - curRemainingAmount
             curAppliedAmount = curRemainingAmount
             curRemainingAmount = CCur(0)
          End If
          'Response.Write "curNewOutAmount = " & curNewOutAmount & "<br>"

          
          strSQL = ""
          strSQL = strSQL & " " & Chr(13) & "UPDATE Invoice"
          strSQL = strSQL & " " & Chr(13) & "SET outstanding_amount = " & curNewOutAmount
          If curNewOutAmount = 0 Then
             strSQL = strSQL & " " & Chr(13) & ", invoicestatus_code = 'BP'"  
          End If
          strSQL = strSQL & " " & Chr(13) & "WHERE invoice_id = " & arryInvoice(0,i)
          dbconTmp.Execute (strSQL)
          
          strSQL = "insert into InvoicePayment (invoice_id, payment_id, applied_amount) values(" & arryInvoice(0,i) & ", " & lngPayId & ", " & curAppliedAmount & ")"
          dbconTmp.Execute (strSQL)
       
          i = i + 1
       Wend
       
       Erase arryInvoice
       
       If curRemainingAmount > 0 Then
          curCreditAmount = curRemainingAmount * -1
          strSQL = ""
          strSQL = strSQL & " " & Chr(13) & "Insert Into LineItem (project_id, lineitemtype_code, lineitemstatus_code, descrip, comment, amount, qty, sequence_num, rate_code)"
          strSQL = strSQL & " " & Chr(13) & "values ("
          strSQL = strSQL & " " & Chr(13) & "" & lngProjID & ","
          strSQL = strSQL & " " & Chr(13) & "'" & "PAYCRED" & "',"
          strSQL = strSQL & " " & Chr(13) & "'" & "PEND" & "',"
          strSQL = strSQL & " " & Chr(13) & "'" & "Payment Credit" & "',"
          strSQL = strSQL & " " & Chr(13) & "NULL,"
          strSQL = strSQL & " " & Chr(13) & "" & curCreditAmount & ","
          strSQL = strSQL & " " & Chr(13) & "" & 1 & ","
          strSQL = strSQL & " " & Chr(13) & "" & 1 & ","
          strSQL = strSQL & " " & Chr(13) & "NULL"
          strSQL = strSQL & " " & Chr(13) & ")"
          dbconTmp.Execute (strSQL)
       End If
       
       
       
       
    Else    'There are no invoices, so create a credit
       
       curCreditAmount = curAmount * -1
       strSQL = ""
       strSQL = strSQL & " " & Chr(13) & "Insert Into LineItem (project_id, lineitemtype_code, lineitemstatus_code, descrip, comment, amount, qty, sequence_num, rate_code)"
       strSQL = strSQL & " " & Chr(13) & "values ("
       strSQL = strSQL & " " & Chr(13) & "" & lngProjID & ","
       strSQL = strSQL & " " & Chr(13) & "'" & "PAYCRED" & "',"
       strSQL = strSQL & " " & Chr(13) & "'" & "PEND" & "',"
       strSQL = strSQL & " " & Chr(13) & "'" & "Payment Credit" & "',"
       strSQL = strSQL & " " & Chr(13) & "NULL,"
       strSQL = strSQL & " " & Chr(13) & "" & curCreditAmount & ","
       strSQL = strSQL & " " & Chr(13) & "" & 1 & ","
       strSQL = strSQL & " " & Chr(13) & "" & 1 & ","
       strSQL = strSQL & " " & Chr(13) & "NULL"
       strSQL = strSQL & " " & Chr(13) & ")"
       
       dbconTmp.Execute (strSQL)
       
    End If

    'Response.End

    strSQL = "update Payment set paymentstatus_code = 'APPLIED' where payment_id = " & lngPayId
    dbconTmp.Execute (strSQL)
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE Project"
    strSQL = strSQL & " " & Chr(13) & "set project_balance = project_balance - " & curAmount
    strSQL = strSQL & " " & Chr(13) & "where project_id = " & lngProjID

    dbconTmp.Execute (strSQL)
    

    If Err.Number <> 0 Then
       fn_BL_AddPayment = Err.Number
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
   
    fn_BL_AddPayment = 0

End Function
%>