<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!-- #include file="Utility/Util.asp" -->
<!-- #include file="Utility/DBUtil.asp" -->
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
	Dim rsTmp
	Dim i, row


    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
	strSQL = "SELECT * FROM ProjectQuestion, ProjectAnswer Where project_id = " & Request.QueryString("pID") & " AND ProjectQuestion.code = ProjectAnswer.question_code"
    Set rsTmp = Server.CreateObject("adodb.recordset")
    rsTmp.Open strSQL, dbconTmp, 3, 3	



	Response.Write "<center><font size=5><b>Initial Field Questions for " & Request.QueryString("pName") & " </b></font></center><br><br>"

	Do While Not rsTmp.EOF
		Response.Write "<b>" & rsTmp.Fields("code") & ":&nbsp;" & rsTmp.Fields("question") & "<b><br>"
		Response.Write "&nbsp;&nbsp;&nbsp;<font color=red>" & rsTmp.Fields("answer") & "</font><br><br>"


	rsTmp.MoveNext
	Loop

rsTmp.Close
Set rsTmp = Nothing

dbconTmp.Close
Set dbconTmp = Nothing

%>






