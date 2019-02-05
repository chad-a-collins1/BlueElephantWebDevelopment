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


Dim conn, rs, strSQL, rs2, curAcctBal

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open fn_GetConnectionString

        Set rs2 = Server.CreateObject("adodb.recordset")
        strSQL = "select sum(project_balance) as account_balance from Project where account_id = " & Request.QueryString("aid")
        rs2.Open strSQL, conn, 3, 3
        
        curAcctBal = rs2.Fields("account_balance")
        
        rs2.Close
        Set rs2 = Nothing

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Account, BillingInfo WHERE Account.account_id = " & Request.QueryString("aid") & " AND BillingInfo.account_id = Account.account_id" 
	rs.Open strSQL, conn, 3, 3
%>


<HTML>

<HEAD>
<script language="javascript1.2">

function fctLoadControlPanel3(aID) {

window.open("adm_cntlPanel3.asp?aID=" + aID, Target="Right");

}

</script>
</HEAD>

<BODY onLoad="fctLoadControlPanel3(<%= Request.QueryString("aid") %>)">

<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "black";}
</STYLE>

<br>
<center>

<table width="99%"><tr><td>
<font size="4" class="bodycopy">
<table border=0 cellpadding=0>

<tr><td><b>ACTIVATION DATE:</b>&nbsp;&nbsp;&nbsp;<%= FormatDateTime(rs.Fields("insert_datetime"), vbShortDate) %></td></tr>
<tr><td><b>ACCOUNT STATUS:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("accountstatus_code") %></td></tr>
<tr><td><b>ACCOUNT TYPE:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("accounttype_code") %></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>

<tr><td>ACCOUNT NUMBER:&nbsp;&nbsp;&nbsp;<%= rs.Fields("account_num") %></td></tr>
<tr><td>USERNAME:&nbsp;&nbsp;&nbsp;<%= rs.Fields("username") %></td></tr>
<tr><td>PASSWORD:&nbsp;&nbsp;&nbsp;<%= rs.Fields("account_password") %></td></tr>


<tr><td><Input="hidden" name="cID" value="<%= Request.QueryString("aid") %>"><br><br></td></tr>

<tr colspan=4><td colspan=4></td><td></td><td bgcolor=red></td></tr>

<tr><td><b>COMPANY:</b>&nbsp;&nbsp;&nbsp;<font color="blue"><b><%= rs.Fields("business_name") %><b></font></td></tr>

<tr><td><b>CONTACT NAME:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("first_name") & Space(1) & rs.Fields("last_name") %></td></tr>

<tr><td><b>PHONE:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("phone1") %></td></tr>

<tr><td><b>FAX:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("phone2") %></td></tr>

<tr><td><b>EMAIL:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("email") %></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td><b>TOTAL BALANCE:&nbsp;&nbsp;&nbsp;$&nbsp;<% = curAcctBal %></b></td></tr>


</table>
</font>


</td></tr>

</center>


<%
rs.Close
set rs = Nothing
conn.Close
set conn = Nothing
%>


</BODY>
</HTML>
