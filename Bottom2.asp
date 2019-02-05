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


	Dim dbconTmp
	Dim rs
	Dim strSQL
	Dim i, row
	   
	Set dbconTmp = Server.CreateObject("ADODB.Connection")
	dbconTmp.ConnectionString = fn_GetConnectionString
	dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM InitialContact WHERE initialcontact_id = " & Request.QueryString("cid") 
	rs.Open strSQL, dbconTmp, 3, 3

%>


<HTML>

<HEAD>
<script language="javascript1.2">

function fctLoadadm_rightPanel_Contacts(cID) {

window.open("adm_rightPanel_Contacts.asp?cID=" + cID, Target="Right")

}

</script>
</HEAD>

<BODY onLoad="fctLoadadm_rightPanel_Contacts(<%= Request.QueryString("cid") %>)">

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
<tr><td><b>CONTACT DATE:</b>&nbsp;&nbsp;&nbsp;<%= FormatDateTime(rs.Fields("insert_datetime"), vbShortDate) %></td></tr>
<tr><td><Input="hidden" name="cID" value="<%= Request.QueryString("cid") %>"><br><br></td></tr>
<tr colspan=4><td colspan=4></td><td></td><td bgcolor=red></td></tr>
<tr><td><b>COMPANY:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("company") %></td></tr>
<tr><td><b>CONTACT NAME:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("contactname") %></td></tr>

<tr><td><b>STREET:</b>&nbsp;&nbsp;&nbsp;</td></tr>
<tr><td><b>CITY:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("city") %></td></tr>
<tr><td><b>STATE:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("state") %></td></tr>
<tr><td><b>ZIP:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("zip") %></td></tr>
<tr><td><b>PHONE:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("phone") %></td>
<tr><td><b>FAX:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("fax") %></td></tr>
<tr><td><b>EMAIL:</b>&nbsp;&nbsp;&nbsp;<%= rs.Fields("email") %></td></tr>
<tr><td><b>PROBLEM DESCRIPTION:</b></td></tr>
<tr><td colspan=4><%= rs.Fields("summary") %></td></tr>
</table>
</font>
<!--

<form name="theForm" method="post" action="./DataLayer/DLinitiateClientAccount.asp">
<table bgcolor=red>
<tr><td><select name="OpenYN"><option value="1">Close Account</OPtion><option value="2">Open Account</OPtion></select></td></tr>
<tr><td></td></tr>
<tr><td><select name="consultant"><option value="0">Not Assigned</OPtion><option value="1">Chad</OPtion><option value="2">Danny</OPtion></select></td></tr>
<tr><td><input type="hidden" name="initialcontact_id" value="Request.QueryString("cid")"></td></tr>
<tr><td><input type="submit" value="Update"></td></tr>
</table>
</form>

-->


</td></tr>

</center>


<%
rs.Close
set rs = Nothing
dbconTmp.Close
set dbconTmp = Nothing
%>


</BODY>
</HTML>




