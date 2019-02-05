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
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<html>
<head>
<title>Admin Contacts</title>

<script language="javascript1.2">

function clearCenterPanel() {

window.open("Bottom1.asp", Target="BottomMain" )

}

</script>

</head>
<body text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" bgcolor="silver" onLoad="clearCenterPanel()">
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 9pt;font-family: "Verdana";color: "black"; }
	th {font-size: 10pt;font-family: "Verdana"; color: "black";}
	a {font-size: 9pt;font-family: "Verdana"; color: "black";}
</STYLE>

<%

	Dim strProvider
	Dim DB
	Dim strSQL
	Dim conn
	Dim rs
	Dim i, row
	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open fn_GetConnectionString
	Set rs = Server.CreateObject("adodb.recordset")
	
	strSQL = "SELECT BillingInfo.billinginfo_id, BillingInfo.account_id, BillingInfo.billinginfostatus_code, BillingInfo.business_name, Project.project_id, Project.project_name, Project.status_code, Consultant.consultant_id, Consultant.fname, Consultant.lname"
        strSQL = strSQL & " FROM (Consultant RIGHT JOIN Project ON Consultant.consultant_id = Project.consultant_id) INNER JOIN BillingInfo ON Project.account_id = BillingInfo.account_id"
        strSQL = strSQL & " ORDER BY Project.project_name ASC"


'	strSQL = "SELECT Account.*, Project.*, BillingInfo.*, Consultant.fname, consultant.lname FROM Account, Project , BillingInfo , Consultant WHERE Account.account_id = BillingInfo.account_id AND Account.account_id = Project.account_id ORDER BY Project.project_name ASC"
		
%>
<CENTER>
<FONT FACE="VERDANA" COLOR="white">
<%

		rs.Open strSQL, conn, 3, 3
		rs.MoveFirst


		%>
		<CENTER>
		<table width="99%" border=1 bordercolor="GRAY" cellspacing=0 cellpadding=0 >
		<tr>
			<th bgcolor="lightblue" size=35><center>Project Name</center></th>
			<th bgcolor="lightblue" size=100><center>Business Name</center></th>
			<th bgcolor="lightblue" size=35><center>Assigned To</center></th>
		</tr>
		<%
		row = 1
		Do While Not rs.EOF 
			If row mod 2 = 0 AND NOT rs.Fields("status_code") = "DEVINPROG" Then
				%>
				<tr bgcolor="lightblue">
				<%

			Elseif rs.Fields("status_code") = "DEVINPROG" Then
				%>
				<tr bgcolor="yellow">
				<%
			Else
				%>
				<tr bgcolor="lightgrey">
				<%
			End If
				%>


				<td size=35><center><B><%= rs.Fields("project_name") %></B></center></td>   
				<td size=100><center><a href="Bottom4.asp?pid=<%= rs.Fields("project_id") %>"  Target="BottomMain" ><B><FONT COLOR="blue"><%= rs.Fields("business_name") %></FONT></B></a></center></td>
  				<td size=35><B><% = rs.Fields("fname") & " " & rs.Fields("lname") %></B></td>			
	 		</tr>		
			<%
			rs.MoveNext 
			row = row + 1
		Loop

	rs.Close 
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
		
%>
</table>
</CENTER>
</FONT>
</body>
</html>

























