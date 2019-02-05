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

<script language="javascript1.2">

function fctChangePassword(aID) {

	window.navigate("admProject_ChangePassword.asp?aID=" + aID);

}


function fctCreateNewProject(aID) {

	window.open("admProject_CreateNewProject.asp?aID=" + aID);

}


function fctChangeBillInfo(aID) {

	window.open("admProject_EditBillInfo.asp?aID=" + aID);

}


function fctChangeContactInfo(aID) {

	window.open("admProject_EditContactInfo.asp?aID=" + aID);

}

</script>

<html>
<head>
</head>

<%
   Dim dbconTmp
   Dim rs
   Dim strSQL
   Dim aID
   
   aID =  Request.QueryString("aID")
   
   Set dbconTmp = Server.CreateObject("ADODB.Connection")
   dbconTmp.ConnectionString = fn_GetConnectionString
   dbconTmp.Open 

   Set rs = Server.CreateObject("adodb.recordset")
   strSQL = "SELECT * FROM Account, BillingInfo WHERE Account.account_id = " & aID & " AND Account.account_id = BillingInfo.account_id"
   rs.Open strSQL, dbconTmp, 3, 3

%>

<body bgcolor="#ededed">
<br><br><br><br><br><br>
<center>
<input type="button" name="edit_Username" value="Edit UserName/Password" onClick="fctChangePassword(<% = aID %>)">
<br><br>

<input type="button" name="createProject" value="Create New Project" onClick="fctCreateNewProject(<% = aID %>)">
<br><br>
<input type="button" name="editBillingInfo" value="Edit Billing Info" onClick="fctChangeBillInfo(<% = aID %>)">
<br><br>
<input type="button" name="editContactInfo" value="Edit Contact Info" onClick="fctChangeContactInfo(<% = aID %>)">
</center>


</body>
</html>

<%

   rs.Close
   Set rs = Nothing
   dbconTmp.CLose
   Set dbconTmp = Nothing
   
%>












