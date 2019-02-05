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
	Dim req_count, i, x
	Dim arryTmp(10)
	Dim txtSignature
	Dim Item
	Dim pID

   pID = Request.QueryString("pID")

	Set dbconTmp = Server.CreateObject("ADODB.Connection")
    	dbconTmp.ConnectionString = fn_GetConnectionString
    	dbconTmp.Open 
    

		For Item=1 to Request.Form("txtBAC_Signature").Count
		If Len(Request.Form("txtBAC_Signature")(Item)) = 0 Then
		Else		
		   
	
	       strSQL = "SELECT * FROM Requirement Where requirement_id = " & CInt(Request.Form("lngReqID")(Item)) 
    	    Set rsTmp = Server.CreateObject("adodb.recordset")
    	    rsTmp.Open strSQL, dbconTmp, 3, 3	

			If Not rsTmp.BOF And Not rsTmp.EOF Then

	    		rsTmp.Fields("chaddan_approved_date") = Now
	    		rsTmp.Update
	    		rsTmp.Close
	    		Set rsTmp = Nothing	

			End If

		End If
	   Next




	dbconTmp.Close
	Set dbconTmp = Nothing

	Response.Redirect("admProject_ViewRequirements.asp?pID=" & pID )

%>


										


















































































