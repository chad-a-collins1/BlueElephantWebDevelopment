<%@
LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp" -->
<!--#include file="Utility/DBUtil.asp"-->
<%

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "DELETE from InitialContact where consultant_id is not null"
    
    'Response.Write "SQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing


%>