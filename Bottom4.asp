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

<HTML>

<HEAD>
<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">

<script language="javascript1.2">

function fctLoadControlPanel4(pID, pName) {

window.open("adm_cntlPanel4.asp?pID=" + pID + "&pName=" + pName, Target="Right");

}

</script>

</HEAD>
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
    Dim blobMOUbody
    
    lngProjectID = Request.QueryString("pID")

    lngRC = fn_BL_GetProjectInfo(lngProjectID, arryTmp1, arryTmp2)
    If lngRC <> 0 Then
       Call sub_HandleLogicError(lngRC)
    End If
    
'Project.project_id, Project.account_id, Project.project_name, Project.projecttype_code, Project.down_payment, 
'Project.status_code, Project.insert_datetime, Project.project_balance, Project.project_estimate_hours, 
'Project.project_estimate_cost, Project.target_date, Project.completion_date, Project.mou_signed, Project.mou_signee, 
'Project.mou_signed_date, Project.mou_body, Project.first_time, Project.accrued_hours, Project.consultant_id, MOU.mou_body 
'ProjectType.descrip, ProjectStatus.descrip, AccountStatus.descrip, AccountType.descrip, 
'BillingInfo.balance, BillingInfo.balance_forward, BillingInfo.billinginfostatus_code,
'BillingInfo.business_name,BillingInfo.first_name, BillingInfo.last_name, BillingInfo.phone1, BillingInfo.phone2,
'BillingInfo.email, BillingInfo.address, BillingInfo.city, BillingInfo.state_code, BillingInfo.postal_code, 
'BillingInfo.state_descrip, BillingInfo.country_descrip, BillingInfoStatus.descrip"
   
    
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
 ' accrued_hours
 'consultant_id
 'MOU_Body
 strProjectTypeDescrip = arryTmp1(20,0)    'ProjectType.descrip, 
 strProjectStatusDescrip = arryTmp1(21,0) 'ProjectStatus.descrip, 
   
 'AccountStatus.descrip, 
 'AccountType.descrip, 
'   balance   
'   balance_forward
'   billinginfostatus_code
   strBusinessName = arryTmp1(27,0)
'   first_name
'   last_name
'   phone1
'   phone2
'   email
'   address
'   city
'   state_code
'   postal_code
'   state_descrip
'   country_dscrip
' BillingInfo StatusDescrip

 

%>




<BODY onLoad="fctLoadControlPanel4(<%= lngProjectID %>, <%= "'" & strProjectName & "'" %>)">

<STYLE TYPE="text/css">
	p {text-align:justify;font-size: 9pt;font-family: "Verdana"; }
	td {font-size: 8pt;font-family: "Verdana";color: "darkblue"; }
	th {font-size: 12pt;font-family: "Verdana"; color: "black";}
	a {font-size: 10pt;font-family: "Verdana"; color: "black";}
</STYLE>

<br>
<center>

<table width="99%"><tr><td>
<font size="4" class="bodycopy">
<table border=0 cellpadding=0>
<tr><td><b>PROJECT NAME:</b></td><td>&nbsp;</td><td><% = strProjectName %></td></tr>
<tr><td><b><format color="red"><b>PROJECT STATUS:</b></b></td><td>&nbsp;</td><td><% = strProjectStatusDescrip %></format></td></tr>
<tr><td></td><td></td></tr>
<tr><td></td><td></td></tr>
<tr><td><b>TARGET DATE:</b></td><td>&nbsp;</td><td><% = dtTargetDate %></td></tr>
<tr><td></td><td>&nbsp;</td><td></td></tr>
<tr><td></td><td>&nbsp;</td><td></td></tr>
<tr><td><b>BUSINESS NAME:</b></td><td>&nbsp;</td><td><% = strBusinessName %></td></tr>
<tr><td><Input="hidden" name="cID" value="<% = lngAccountID %>"><br><br></td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr><td><b>ESTIMATED COST:</b></td><td>&nbsp;</td><td>$<% = curEstCost %></td></tr>
<tr><td><b>PROJECT BALANCE:</b></td><td>&nbsp;</td><td>$<% = curProjectBalance %></td>
<tr><td><b>DOWN PAYMENT:</b></td><td>&nbsp;</td><td>$<% = curDownPayment %></td>
<tr><td><b>ESTIMATED HOURS:</b></td><td>&nbsp;</td><td><% = intEstHours %>&nbsp;Hours</td></tr>
<tr><td><b>HOURLY RATE:</b></td><td>&nbsp;</td><td>
<%

If intEstHours <= 20 Then
	Response.Write("$40 / Hour")
ElseIf intEstHours > 20 AND intEstHours <= 45 Then
	Response.Write("$38 / Hour")
ElseIf intEstHours > 45 AND intEstHours <= 80 Then
	Response.Write("$34 / Hour")
ElseIf intEstHours > 80 Then
	Response.Write("$28 / Hour")	
End If


Dim rsPay, dbconTmp, curTotalPay, dtLastPayDate, curLastPayAmt, strSQL

Set dbconTmp = Server.CreateObject("ADODB.Connection")
Set rsPay = Server.CreateObject("ADODB.Recordset")

dbconTmp.ConnectionString = fn_GetConnectionString
dbconTmp.Open

strSQL = "select sum(amount) as total_pay, max(payment_datetime) as last_pay  from Payment where project_id = " & lngProjectID
Set rsPay = dbconTmp.Execute (strSQL)
curTotalPay = rsPay.Fields("total_pay")
dtLastPayDate = rsPay.Fields("last_pay")
rsPay.Close

If IsDate(dtLastPayDate) Then
   strSQL = "select amount as last_pay_amt from Payment where project_id = " & lngProjectID & " and payment_datetime = #" & dtLastPayDate & "# order by insert_datetime desc"
   Set rsPay = dbconTmp.Execute (strSQL)
   curLastPayAmt = rsPay.Fields("last_pay_amt")
   rsPay.Close
End If


Dim intInvoiceCount, dtLastInvDate, curLastInvAmt, strLastInvStat

strSQL = "select count(*) as total_count, max(invoice_date) as last_inv from Invoice where project_id = " & lngProjectID
Set rsPay = dbconTmp.Execute (strSQL)
intInvoiceCount = rsPay.Fields("total_count")
dtLastInvDate = rsPay.Fields("last_inv")
rsPay.Close

If IsDate(dtLastInvDate) Then
   strSQL = "select total_amount as last_inv_amt, ist.descrip from Invoice i, InvoiceStatus ist where project_id = " & lngProjectID & " and invoice_date = #" & dtLastInvDate & "# and i.invoicestatus_code = ist.invoicestatus_code order by i.insert_datetime desc"
   Set rsPay = dbconTmp.Execute (strSQL)
   curLastInvAmt = rsPay.Fields("last_inv_amt")
   strLastInvStat = rsPay.Fields("descrip")
   rsPay.Close
End If

Set rsPay = Nothing

%></td></tr>
<tr><td></td><td>&nbsp;</td><td></td></tr>
<tr><td></td><td>&nbsp;</td><td></td></tr>
<tr><td><b><font color="red">TOTAL PAYMENTS:</font></b></td><td>&nbsp;</td><td>$<% = curTotalPay %></td>
<tr><td nowrap><b><font color="green">LAST PAYMENT AMOUNT:</font></b></td><td>&nbsp;</td><td>$<% = curLastPayAmt %></td></tr>
<tr><td><b><font color="green">LAST PAYMENT DATE:</font></b></td><td>&nbsp;</td><td><% = dtLastPayDate %></td></tr>
<tr><td></td><td>&nbsp;</td><td></td></tr>
<tr><td></td><td>&nbsp;</td><td></td></tr>
<tr><td><b><font color="green">INVOICE COUNT:</font></b></td><td>&nbsp;</td><td><%= intInvoiceCount %></td></tr>
<tr><td><b><font color="green">LAST INVOICE AMOUNT:</font></b></td><td>&nbsp;</td><td>$<%= curLastInvAmt %></td></tr>
<tr><td><b><font color="green">LAST INVOICE DATE:</font></b></td><td>&nbsp;</td><td><% = dtLastInvDate %></td></tr>
<tr><td><b><font color="green">LAST INVOICE STATUS:</font></b></td><td>&nbsp;</td><td><% = strLastInvStat %></td></tr>
</table>
</font>
</td></tr>
</center>
</BODY>
</HTML>



















































