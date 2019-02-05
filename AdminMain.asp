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
  If Session("blnAdminLoggedIn") <> True Then
     Response.Redirect "Error.asp"
  End If
%>
<HTML>
<HEAD>
<TITLE>Admin Portal</TITLE>
</HEAD>
<!-- frames -->
<FRAMESET  ROWS="10%,*" FRAMEBORDER="0" BORDER=0 FRAMESPACING="0" BORDER="0">
	    <FRAME NAME="Top" SRC="Top.asp" MARGINWIDTH="0" MARGINHEIGHT="0" SCROLLING="no" FRAMEBORDER="no" NORESIZE>
	    <FRAMESET COLS="40%,*,30%" FRAMEBORDER="0" BORDER=0 FRAMESPACING="0" BORDER="0">
			<FRAME NAME="Mn" SRC="viewContacts.asp" MARGINWIDTH="0" MARGINHEIGHT="0" SCROLLING="yes" FRAMEBORDER="no" NORESIZE>
			<FRAME NAME="BottomMain" SRC="Bottom1.asp" MARGINWIDTH="0" MARGINHEIGHT="0" SCROLLING="no" FRAMEBORDER="yes" NORESIZE>
			<FRAME NAME="Right" SRC="Left.htm" MARGINWIDTH="0" MARGINHEIGHT="0" SCROLLING="no" FRAMEBORDER="yes" NORESIZE>
		</FRAMESET>	
</FRAMESET>
</HTML>





