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
<title>Requirements Console</title>


</head>
<body text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" bgcolor="#EDEDED">
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
	Dim dbconTmp
	Dim rsTmp, rs
	Dim i, row


    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "SELECT * FROM Requirement, RequirementStatus Where project_id = " & Request.QueryString("pID") & " AND Requirement.status_code = RequirementStatus.code"
    Set rsTmp = Server.CreateObject("adodb.recordset")
    rsTmp.Open strSQL, dbconTmp, 3, 3	


    Response.Write "<center><font size=5><b>Requirement Console:&nbsp;&nbsp; " & Request.QueryString("pName") & " </b></font></center><br><br>"
%>


<form action="admProject_UpdateRequirements.asp?pID=<%= Request.QueryString("pID") %>" method="post">
		<table width="99%" border=1 bordercolor="GRAY" cellspacing=2 cellpadding=5 bordercolor="black" >
		<tr>
			<th bgcolor="lightblue" size=35><center>Requirement Name</center></th>
			<th bgcolor="lightblue" size=250><center>Summary</center></th>
			<th bgcolor="lightblue" size=35><center>Client Status</center></th>
			<th bgcolor="lightblue" size=35><center>BAC Approved?</center></th>
			<th bgcolor="lightblue" size=35><center>BAC Signature</center></th>
		</tr>


<%
		i = 0
		Do While Not rsTmp.EOF 
	
			%>
			<tr>
				<td size=35><center><B><%= rsTmp.Fields("requirement_name") %></B></center></td>
				<td size=100 bgcolor="white"><%= rsTmp.Fields("summary") %><B><FONT COLOR="blue"></FONT></B></a></td>
  				<td size=35><B>
				<% 
					If rsTmp.Fields("status_code") = "APPROVED" Then
						Response.Write "<font color=green>" & rsTmp.Fields("status_code") & "</font>"
					Else
						Response.write rsTmp.Fields("status_code")
					End If 
				
				%></B></td>
  				<td size=35><B>
				
			<% 

					If IsNull(rsTmp.Fields("chaddan_approved_date")) = True Then 
						i = i + 1
						Response.Write "<font color=red>Not Approved</font></B></td>"
						'Response.Write "<td size=35><input type=hidden name=lngReqID_" & i & ","
					Response.Write "<td size=35><input type=hidden name=lngReqID" & " "
						Response.Write "value=" & rsTmp.Fields("requirement_id") & ">"
						'Response.Write "<input type=text name=txtBAC_Signature_" & i & "></td></tr>"
					Response.Write "<input type=text name=txtBAC_Signature></td></tr>"

					Else
						Response.Write "<font color=green>Approved</font></B></td>"
						Response.Write "<td size=35></td></tr>"		
					End If

			rsTmp.MoveNext 
		Loop

%>
<tr><td colspan=5></td></tr>
<tr>
<td></td><td></td><td></td><td></td>
<td><input type="submit" value="Update"></td>

</tr>
		</table>

</form>

<%
rsTmp.Close
Set rsTmp = Nothing

dbconTmp.Close
Set dbconTmp = Nothing

%>









