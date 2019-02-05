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
<script language="javascript1.2">

function fctViewSurvey(pID, pName) {

	window.open("admProject_ViewSurvey.asp?pID=" + pID + "&pName=" + pName);

}


function fctViewRequirements(pID, pName) {

	window.open("admProject_ViewRequirements.asp?pID=" + pID + "&pName=" + pName);

}


function fctCreateInvoice(pID) {

	window.navigate("admProject_CreateInvoice_1.asp?pID=" + pID);

}



function fctEnterPayment(pID) {

	window.open("admProject_EnterPayment.asp?pID=" + pID);

}


function fctCreateMoU(pID) {

	window.open("admProject_CreateMoU.asp?pID=" + pID);

}


function fctEnterTime(pID) {

	window.open("admProject_EnterTime.asp?pID=" + pID);

}


function fctEditEstimates(pID) {

	window.open("admProject_EditEstimates.asp?pID=" + pID);

}

</script>
</head>

<%

	Dim dbconTmp
	Dim rs
	Dim strSQL
	Dim i, row, pID, pName
	   
	pID = Request.QueryString("pID")   
	   
	Set dbconTmp = Server.CreateObject("ADODB.Connection")
	dbconTmp.ConnectionString = fn_GetConnectionString
	dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM Project, ProjectStatus WHERE Project.project_id = " & pID & " AND ProjectStatus.code = Project.status_code"

	rs.Open strSQL, dbconTmp, 3, 3

%>



<body bgcolor="#ededed">
<br>
<center>
<%
 	response.write("<format color=red><b>STATUS:</b> " & rs.Fields("descrip") & "</format>") 
%>
</center>
<br>
<br>
<% 'Response.Write "pID = " & pID %>
<form method="post" action="admProject_ProcessManagement.asp?pID=<% = pID %>" name="processTrigger"> 
<center>
<table><tr><td align="right">
Proceed to Next Phase:&nbsp;<input type="radio" name="processControls" value="0">
</td></tr>
<tr><td align="right">
Backup to Previous Phase:&nbsp;<input type="radio" name="processControls" value="1">
</td></tr>
<tr><td align="right">
Reset Project:&nbsp;<input type="radio" name="processControls" value="2">
</td></tr>
<tr><td align="right">
Hold:&nbsp;<input type="radio" name="processControls" value="3" checked>
</td></tr>
<tr><td><center><input type="submit" value="Update" name="update"></center></td></tr>
</table>
</center>
</form>

<br>
<br>
<br>
<center>
<table cellspacing=0 cellpadding=1>
<tr><td>
<input type="button" value="     View Survey    " onClick="fctViewSurvey(<%= pID %>, <%= "'" &  pName & "'" %>)">
</td></tr>
<tr><td>
<input type="button" value="  View Requirements " onClick="fctViewRequirements(<%= pID %>, <%= "'" & pName & "'" %>)">
</td></tr>
<tr><td>
<input type="button" value="   Create Invoice   " onClick="fctCreateInvoice(<%= pID %>)">
</td></tr>
<tr><td>
<input type="button" value=" Payments Received " onClick="fctEnterPayment(<%= pID %>)">
</td></tr>
<tr><td>
<input type="button" value="     Create MoU     " onClick="fctCreateMoU(<%= pID %>)">
<tr><td>
<br>
<input type="button" value="  Input Time Charge " onClick="fctEnterTime(<%= pID %>)">
</td></tr>
<tr><td>
<input type="button" value="Edit Project Estimates" onClick="fctEditEstimates(<%= pID %>)">
</td></tr>
</table>
</center>
</body>
</html>




























