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
<!--#include file="./BusinessLayer/BL_adm_setPayment.asp" -->
<%
Dim pID
Dim dtPymntRecvd
Dim lngPymntInvcNmb
Dim curpymntAmntRcvd


pID = Request.QueryString("pID")



Request.Form("paymentReceivedDate")
Request.Form("paymentInvoiceNumber")
Request.Form("paymentAmountReceived")



function fn_BL_adm_setInvoicePayment(invoice_id, payment_id, amount)
function fn_BL_adm_setPayment(invoice_id, paymenttype_code, descrip, comment, amount, paymentdate)



%>



