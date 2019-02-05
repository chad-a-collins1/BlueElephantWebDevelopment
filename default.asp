<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<%
' this is code you use to put in your homepage
' it will redirect any subDOMAIN to a subDIRECTORY on your website
'

dim fullhost
dim subdomain
dim domain1
dim domain2
dim firstDot
dim secondDot
dim vbcrlf
fullhost=request.servervariables("HTTP_HOST")
firstdot=instrrev(fullhost,".")
domain1=mid(fullhost,firstdot+1)
seconddot=instrRev(fullhost,".",firstdot-1)
if seconddot>0 then
	subdomain=mid(fullhost,1,seconddot-1)
	else
	subdomain=""
end if
if subdomain="" or ucase(subdomain)="WWW" then
	'response.write("index.asp")
	response.redirect("index.htm")
	else
	'response.write("/" & subdomain & "/")
	response.redirect("/" & subdomain & "/")
end if
'InstrRev(string1, string2[, start[, compare]])
'Mid(string, start[, length])
%>

