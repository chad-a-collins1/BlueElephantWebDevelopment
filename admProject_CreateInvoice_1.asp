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
Dim pID
pID = Request.QueryString("pID")
%>

<form method="post" action="admProject_CreateInvoice_2.asp?pID=<%= pID %>" name="theForm">
<br><br><br><br><br><br>
<center>
<table>
	<tr>
	<td><B>INVOICE TYPE:</b></td>
	<td><SELECT name="TYPE">
	     <OPTION Value="BUF">Bill Up Front</OPTION>
	     <OPTION Value="MO" selected>Mail Out</OPTION>
	   </SELECT></td>
	</tr>

	<tr>
	<td><b>INVOICE STATUS:</b></td>
	<td><SELECT name="STATUS">
	            <OPTION Value="BC">bill cancelled</OPTION>
		    <OPTION Value="BD">bill dispute</OPTION>
		    <OPTION Value="BO">bill out</OPTION>
		    <OPTION Value="BP">bill paid</OPTION>
		    <OPTION Value="BW" selected>bill waiting</OPTION>
	    </SELECT></td>
	</tr>


	<tr>
	<td colspan=2><b>DESCRIPTION:</b><br><TEXTAREA cols=30 rows=5 name="DESCRIP"></TEXTAREA></td>
	</tr>

	<tr>
	<td colspan=2><b>COMMENT:</b><br><TEXTAREA cols=30 rows=5 name="COMMENT"></TEXTAREA></td>
	</tr>

	<tr>
	<td><b>DUE DATE:</b></td>
	<td><input type="text" name="DUE_DATE"></td>
	</tr>

	<tr>
	<td><b>START DATE:</b></td>
	<td><input type="text" name="BEGIN_DATE"></td>
	</tr>





	<tr>
	<td><b>AMOUNT:</b></td>
	<td><input type="text" name="INVOICE_AMOUNT" value="0.00"></td>
	</tr>

	<tr>
	<td><b>BALANCE FORWARD:</b></td>
	<td><input type="text" name="BALANCE_FORWARD" value="0.00"></td>
	</tr>





	<tr>
	<td colspan=2><center><input type="submit" value="Enter"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>