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
<!--#include file="BusinessLayer/BL_adm_getInvoice.asp" -->
<!--#include file="BusinessLayer/BL_adm_getProjectInformation.asp" -->
<html>
<head>

<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">


<script language="javascript1.2">

function fctViewSurvey(pID, pName) {

	window.open("admProject_ViewSurvey.asp?pID=" + pID + "&pName=" + pName);

}


function fctViewRequirements(pID, pName) {

	window.open("admProject_ViewRequirements.asp?pID=" + pID + "&pName=" + pName);

}


function fctCreateInvoice(pID) {

	window.navigate("admProject_CreateInvoice.asp?pID=" + pID);

}

function fctPrintSendInvoice(pID) {

	window.open("admProject_PrintSendInvoice.asp?pID=" + pID);

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
    Dim lngRC
    Dim lngProjectID, lngAccountID, lngBillingInfoID
    Dim strProjectName
    Dim blnMOUSigned
    Dim strMOUSignee
    Dim dtMOUSigned
    Dim curDownPayment
    Dim strProjectStatus
    Dim dtStart
    Dim dtTargetDate
    Dim dtCompletionDate, dtInsertDateTime
    Dim curProjectBalance
    Dim intEstHours
    Dim curEstCost
    Dim blnFirstTime
    Dim strStatusCode
    Dim blnFreeze
    Dim blnLockEdit        
    Dim strProjectStatusDescrip, strProjectStatusCode
    Dim strProjectTypeDescrip, strProjectTypeCode   
    Dim strBusinessName
    Dim arryTmp1    
    Dim arryTmp2
    Dim tmpInvoiceCOunt
    Dim blobMOUbody
  
    lngProjectID = Request.QueryString("pID")
    pID = Request.QueryString("pID")

    lngRC = fn_BL_GetProjectInfo(lngProjectID, arryTmp1, arryTmp2)
    If lngRC <> 0 Then
       Call sub_HandleLogicError(lngRC)
    End If
    
  
    
   lngAccountID = arryTmp1(1,0)
   strProjectName = arryTmp1(2,0)
   strProjectTypeCode = arryTmp1(3,0)
   curDownPayment = arryTmp1(4,0)
   strProjectStatusCode = arryTmp1(5,0)
   dtInsertDateTime = arryTmp1(6,0)
   curProjectBalance = arryTmp1(7,0)
   intEstHours = arryTmp1(8,0)    
   curEstCost = arryTmp1(9,0)       
   dtTargetDate = arryTmp1(10,0)  
   dtCompletionDate = arryTmp1(11,0)
   blnMOUSigned = arryTmp1(12,0)
   strMOUSignee = arryTmp1(13,0) 
   dtMOUSigned = arryTmp1(14,0)   
   blobMOUbody = arryTmp1(15,0)  
   blnFirstTime = arryTmp1(16,0)

 strProjectTypeDescrip = arryTmp1(20,0)    'ProjectType.descrip, 
 strProjectStatusDescrip = arryTmp1(21,0) 'ProjectStatus.descrip, 
   



	Dim dbconTmp
	Dim strSQL
	Dim i, row, pID, pName
	   
	   
	Set dbconTmp = Server.CreateObject("ADODB.Connection")
	dbconTmp.ConnectionString = fn_GetConnectionString
	dbconTmp.Open 

%>



<body bgcolor="#ededed">
<br>
<center>

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
 <tr>
   <td>
     <input type="button" value="     View Survey    " onClick="fctViewSurvey(<%= pID %>, <%= "'" &  pName & "'" %>)">
   </td>
 </tr>
 <tr>
   <td>
     <input type="button" value="  View Requirements " onClick="fctViewRequirements(<%= pID %>, <%= "'" & pName & "'" %>)">
   </td>
 </tr>

 <tr>
   <td>
     <input type="button" value="     Create MoU     " onClick="fctCreateMoU(<%= pID %>)">
   </td>
 </tr>  
 <tr>
  <tr>
   <td>
     <br>
     <input type="button" value="Edit Project Estimates" onClick="fctEditEstimates(<%= pID %>)">
   </td>
 </tr>
   <td>
    <br>
     <input type="button" value="  Input Time Charge " onClick="fctEnterTime(<%= pID %>)">
   </td>
 </tr>
  <tr>
   <td>
     <input type="button" value="   Create Invoice   " onClick="fctCreateInvoice(<%= pID %>)">
   </td>
 </tr>
 <tr>
   <td>
     <input type="button" value="Print/Send Invoices" onClick="fctPrintSendInvoice(<%= pID %>)">
   </td>
 </tr>
 <tr>
   <td>
     <input type="button" value=" Payments Received " onClick="fctEnterPayment(<%= pID %>)">
   </td>
 </tr>


</table>

</center>
</body>
</html>




























