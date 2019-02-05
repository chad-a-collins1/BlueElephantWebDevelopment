<% @LANGUAGE = "VBScript" %>
<%
  Option Explicit
%>
<% Response.Buffer = True %>
<!-- #include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp" -->
<!--#include file="Utility/DBUtil.asp"-->
<!-- #include file="BusinessLayer/BL_adm_setInvoice.asp"-->
<%

	Dim conn, rs, strSQL
	Dim strDBpath
	Dim pID

	'strDBpath = server.MapPath("db/Collins.mdb")
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open fn_GetConnectionString


	pID = Request.QueryString("pID") 

	Set rs = Server.CreateObject("adodb.recordset")
	strSQL = "SELECT * FROM BillingInfo, Project WHERE Project.project_id = " & pID & " AND BillingInfo.account_id = Project.account_id" 
	rs.Open strSQL, conn, 3, 3


Dim billinginfo_id
Dim project_id
Dim invoicetype_code
Dim invoicestatus_code
Dim descrip
Dim comment
Dim balance
Dim balance_forward
Dim total_amount
Dim invoice_date
Dim due_date
Dim outstanding_amount
Dim insert_datetime 
Dim begin_date 
Dim end_date

billinginfo_id = rs.Fields("billinginfo_id") 
project_id = pID
invoicetype_code = Request.Form("TYPE")
invoicestatus_code = Request.Form("STATUS")
descrip = Request.Form("DESCRIP")
comment = Request.Form("COMMENT")
balance = CCur(rs.Fields("project_balance"))
balance_forward = CCur(0)
total_amount = CCur(rs.Fields("project_balance")) - CCur(rs.Fields("down_payment"))
invoice_date = now()
due_date = Request.Form("DUE_DATE")
outstanding_amount = total_amount 
insert_datetime = now()
begin_date = Request.Form("BEGIN_DATE")
end_date= Request.Form("END_DATE")

Dim lngRC


   lngRC = fn_BL_adm_setInvoice(billinginfo_id, project_id, invoicetype_code, invoicestatus_code, descrip, comment, balance, balance_forward, total_amount, invoice_date, due_date, outstanding_amount, insert_datetime, begin_date, end_date)
   If lngRC <> 1 Then
      Response.Write "Problem Creating Invoice!"
      Response.End
   End If

rs.Close
set rs = Nothing
conn.Close
Set conn = Nothing


      Response.Write "Invoice Successfully Created!"


%>