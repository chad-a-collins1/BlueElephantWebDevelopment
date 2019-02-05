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
<%
Dim cID
cID = Request.QueryString("cID")
%>


<html>
<head></head>
<body>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<center>


<form method="post" action="adm_call_InitCntToAcount.asp?cID=<%= cID %>" name="theForm"> 

<table>
<tr><td align="left">
<input type="radio" name="handleContact" value="1">Create Account
</td></tr>
<br><br>
<tr><td align="left">
<input type="radio" name="handleContact" value="2">Delete this Contact
</td></tr>
<tr><td><input type="hidden" name="CID" value="cID"></td></tr>
</table>
<br><br>

<input type="submit" name="submit" value="Execute">

</form>






</center>
</body>
</html>











