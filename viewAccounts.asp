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
	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
		 Dim i, row
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	    Set rs = Server.CreateObject("ADODB.RecordSet")
	    strSQL = "SELECT BillingInfo.account_id, ContactInfo.first_name, ContactInfo.last_name, BillingInfo.account_id, BillingInfo.business_name FROM ContactInfo, BillingInfo WHERE ContactInfo.account_id = BillingInfo.account_id ORDER BY ContactInfo.contactinfo_id DESC"

%>
<CENTER>
<FONT FACE="VERDANA" COLOR="white">
<%

	    rs.Open strSQL, dbconTmp, 3, 3


		%>
		<CENTER>
		<table width="99%" border=1 bordercolor="GRAY" cellspacing=0 cellpadding=0 >
		<tr>
			<th bgcolor="lightblue" size=35><center>Business Name</center></th>
			<th bgcolor="lightblue" size=100><center>Contact Name</center></th>
			<!-- <th bgcolor="lightblue" size=35><center>Assigned To</center></th> -->
		</tr>
		<%
		row = 1
		Do While Not rs.EOF 
			If row mod 2 = 0 Then
				%>
				<tr bgcolor="lightblue">
				<%
			Else
				%>
				<tr bgcolor="lightgrey">
				<%
			End If
			%>
				<td size=35><center><a href="Bottom3.asp?aid=<%= rs.Fields("account_id") %>"  Target="BottomMain" ><B><FONT COLOR="blue"><%= rs.Fields("business_name") %></FONT></B></a></center></td>   
				<td size=100><center><B><%= rs.Fields("first_name") & Space(1) %><%= rs.Fields("last_name") %></B></center></td>
  				<!-- <td size=35><B>N/A</B></td> -->			
	 		</tr>		
			<%
			rs.MoveNext 
			row = row + 1
		Loop

	rs.Close 
	Set rs = Nothing
	dbconTmp.Close
	Set dbconTmp = Nothing
		
%>
</table>
</CENTER>
</FONT>
</body>
</html>










