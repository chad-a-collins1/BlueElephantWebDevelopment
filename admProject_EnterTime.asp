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

Dim pID, strDescrip, strLineItemType, strRate, strQty, curAmount, strDispSeq, strComment, strFrom, lngRC, strMsg
Dim arryRate

strMsg=""
strFrom = Request("txtFrom")
pID = Request.QueryString("pID")

If strFrom = "entertime" Then

    strDescrip = Request("txtDescrip")
    strLineItemType = Request("txtLineItemType")
    strRate = Request("txtRate")
    arryRate =SPlit(strRate,"|")
    strRate = arryRate(0)
    strQty = Request("txtQty")
    'Response.Write "strQty = " & strQty
    curAmount = CCur(arryRate(1)) * CInt(strQty)
    Erase arryRate
    strDispSeq = Request("txtDispSeq")
    strComment = Request("txtComment")
  
  
    lngRC = fn_BL_InsertTimeCharge(pId, strDescrip, strLineItemType, strRate, strQty, curAmount, strDispSeq, strComment)
    If lngRC <> 0 Then
         Response.Redirect "Error.asp?x1=" & lngRC
    End If
   
    Session("sess_strMsg") = "Time Charge Added Suckdickfully Bizzatch!"
    Response.Redirect "admProject_EnterTime.asp?txtFrom=verify&pID=" & pId
    
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
<br>
<form method="post" action="admProject_EnterTime.asp?pID=<%= pID %>" name="theForm">
<input type="hidden" name="txtFrom" value="entertime">
<center>
<table>
<%
  '  Dim arryInv
  '   lngRC = fn_BL_GetProjectInvoices(pId,arryInv)
  '  If lngRC <> 0 Then
  '       Response.Redirect "Error.asp?x1=" & lngRC
  '  End If
    
  ' If Not IsArray(arryInv) Then
  '    Response.Write "NO INVOICES YET FOR THIS PROJECT!"
  '    Response.End
  ' End  If
    
%>
<!--
	<tr>
	<td><B>INVOICE:</b></td>
	<td><SELECT name="txtInvoice">
	     <OPTION Value="none" selected>&lt; Select An Invoice &gt;</OPTION>
-->
<%
'   Dim i
'   For i = 0 to UBound(arryInv,2)
'      Response.Write "<OPTION Value=""" & arryInv(0,i) & """ >" & arryInv(5,i) & " - " & arryInv(10,i) & "</OPTION>" & vbCrLf
'   Next
%>	     
<!--	   </SELECT></td>
	</tr>
-->
	<tr>
	<td><b>Line Item Type:</b></td>
	<td><SELECT name="txtLineItemType">
	            <OPTION Value="SVCCHAR" selected>Service Charge</OPTION>
		    <OPTION Value="ONETIMECHAR">One Time Charge	</OPTION>
		    <OPTION Value="OTHERCHAR">Other Charge</OPTION>
		    <OPTION Value="PRODCHAR">Product Charge</OPTION>
		  <OPTION Value="TAXCHAR">Tax Charge</OPTION>
		  <OPTION Value="GENCRED">General Credit</OPTION>
	    </SELECT></td>
	</tr>
	
	
	

	<tr>
	<td colspan=2><b>DESCRIPTION:</b><br><input type=text size=30 maxlength=50 name="txtDescrip">
	</tr>

	<tr>
	<td><b>Rate:</b></td>
	<td><SELECT name="txtRate">
	            <OPTION Value="GT80HR001|25.00" selected>80 plus hours - $25.00</OPTION>
      		    <OPTION Value="GTE45LT80HR001|34.00">45 to 80 hours -$34.00</OPTION>
		    <OPTION Value="GTE20LT45HR001|38.00">20 to 45 hours -$38.00</OPTION>
		    <OPTION Value="LT20HR001|40.00">less than 20 hours - $40.00</OPTION>
	    </SELECT></td>
	</tr>		

	<tr>
	<td><b>Quantity</b></td>
	<td><input type="text" name="txtQty" value="1"></td>
	</tr>

<!--
	<tr>
	<td><b>Amount: $</b></td>
	<td><input type="text" name="txtAmount"></td>
	</tr>
-->


	<tr>
	<td><b>Display Sequence:</b></td>
	<td><input type="text" name="txtDispSeq" value="0"></td>
	</tr>


	<tr>
	<td colspan=2><b>COMMENT:</b><br><TEXTAREA cols=30 rows=5 name="txtComment"></TEXTAREA></td>
	</tr>




	<tr>
	<td colspan=2><center><input name="cmdSubmit" type="submit" value="Enter"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>


<%

Function fn_BL_InsertTimeCharge (pId, strDescrip, strLineItemType, strRate, strQty, curAmount, strDispSeq, strComment)

    Dim dbconTmp
    Dim rs
    Dim strSQL
    Dim strStatus
    
    strStatus = "PEND"
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "Insert Into LineItem (project_id, lineitemtype_code, lineitemstatus_code, descrip, comment, amount, qty, sequence_num, rate_code)"
    strSQL = strSQL & " " & Chr(13) & "values ("
    strSQL = strSQL & " " & Chr(13) & "" & pId & ","
    strSQL = strSQL & " " & Chr(13) & "'" & strLineItemType & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strStatus & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strDescrip & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strComment & "',"
    strSQL = strSQL & " " & Chr(13) & "" & curAmount & ","
    strSQL = strSQL & " " & Chr(13) & "" & strQty & ","
    strSQL = strSQL & " " & Chr(13) & "" & strDispSeq & ","
    strSQL = strSQL & " " & Chr(13) & "'" & strRate & "'"
    strSQL = strSQL & " " & Chr(13) & ")"

    'Response.Write "strSQL = " & strSQL
    'Response.End
    dbconTmp.Execute (strSQL)
    
    'strSQL = ""
    'strSQL = strSQL & " " & Chr(13) & "UPDATE Invoice"
    'strSQL = strSQL & " " & Chr(13) & "SET balance = balance + " & curAmount & ","
    'strSQL = strSQL & " " & Chr(13) & "total_amount = total_amount + " & curAmount & ","
    'strSQL = strSQL & " " & Chr(13) & "outstanding_amount = outstanding_amount + " & curAmount
    'strSQL = strSQL & " " & Chr(13) & "WHERE invoice_id = " & lngInvoiceId
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    'dbconTmp.Execute (strSQL)

    dbconTmp.Close
    Set dbconTmp = Nothing

   
    fn_BL_InsertTimeCharge = 0
   
End Function
%>