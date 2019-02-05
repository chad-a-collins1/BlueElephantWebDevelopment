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
	Dim UID				'<--------------------------- THIS VALUE IS NOT GETTING PASSED IN FROM THE FRAMES PAGE
	UID = Request.QueryString("UID")
%>
<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>
<script language="javascript1.2">

function clearCenterPanel() {

window.open("Bottom1.asp", Target="BottomMain" )

}

</script>


<title>Admin Contacts</title>
<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 9pt;font-family: "Verdana";color: "black"; }
	th {font-size: 10pt;font-family: "Verdana"; color: "black";}
	a {font-size: 9pt;font-family: "Verdana"; color: "black";}
</STYLE>

</head>
<body text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" bgcolor="silver" onLoad="clearCenterPanel()">

<%

	Dim dbconTmp
	Dim rs
	Dim strSQL
	Dim i, row
	   
	Set dbconTmp = Server.CreateObject("ADODB.Connection")
	dbconTmp.ConnectionString = fn_GetConnectionString
	dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT ic.initialcontact_id, ic.company, ic.contactname, ic.insert_datetime, c.fname, c.lname"
	strSQL = strSQL & " FROM Consultant c RIGHT OUTER JOIN InitialContact ic on c.consultant_id = ic.consultant_id ORDER BY ic.insert_datetime DESC"
	
	'Response.Write strSQL
	
	rs.Open strSQL, dbconTmp, 3, 3


		
%>
<CENTER>
<FONT FACE="VERDANA" COLOR="white">
<CENTER>
<table width="99%" border=1 bordercolor="GRAY" cellspacing=0 cellpadding=0 >
		<tr>
			<th bgcolor="lightblue" size=35><center>Business Name</center></th>
			<th bgcolor="lightblue" size=100><center>Contact Name</center></th>
			<th bgcolor="lightblue" size=35><center>Contact Date</center></th>
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
				<td size=35><center><B><%= rs.Fields("company") %></B></center></td>
				<td size=100><center><a href="Bottom2.asp?cid=<%= rs.Fields("initialcontact_id") %>"  Target="BottomMain" ><B><FONT COLOR="blue"><%= rs.Fields("contactname") %></FONT></B></a></center></td>
  				<td size=35><B><%= FormatDateTime(rs.Fields("insert_datetime"), vbShortDate) %></B></td>
  				<!-- <td size=35><B> = rs.Fields("fname") & " " & rs.Fields("lname") </B></td>	-->			
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
