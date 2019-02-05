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
<!--#include file="BusinessLayer/BL_admContact.asp"-->

<html>
<head>
</head>

<%

	Dim conn
	Dim strDBpath
	strAID = Request.QueryString("aID")
	strDBpath = server.MapPath("Collins.mdb")
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"


	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM Account, BillingInfo WHERE Account.account_id = " & Request.QueryString("aID") & " AND Account.account_id = BillingInfo.account_id"
	rs.Open strSQL, conn, 3, 3
%>

<body bgcolor="#ededed">
<br><br><br><br><br><br>
<center>
<input type="button" name="edit_Username" value="Change Password">
<br><br>
<input type="button" name="createProject" value="Create New Project">
<br><br>
<input type="button" name="editBillingInfo" value="Edit Billing Info">
<br><br>
<input type="button" name="editContactInfo" value="Edit Contact Info">
</center>


</body>
</html>















