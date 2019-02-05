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
	Dim dbconTmp
	Dim rs
	Dim strSQL
	Dim cnrtlValue
	Dim strStatusForward, strStatusBack, strStatus
	Dim currentStatus, pID
	Dim strProjectStatus
	Dim strProject
	Dim rsProjectStatus
	Dim rsProject
	Dim tmpUpdatedStatusCode	


	pID = Request.QueryString("pID")   
	   
	Set dbconTmp = Server.CreateObject("ADODB.Connection")
	dbconTmp.ConnectionString = fn_GetConnectionString
	dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM Project, ProjectStatus WHERE Project.project_id = " & pID & " AND ProjectStatus.code = Project.status_code"	
	rs.Open strSQL, dbconTmp, 3, 3

	currentStatus = rs.Fields("seq_no")
	
		If (currentStatus < 10) Then
		      strStatusForward = CInt(currentStatus) + 1
		Else
		      strStatusForward = 9		      
		End If
		
		
		If (currentStatus > -1) Then
      		   strStatusBack = CInt(currentStatus) - 1
      	Else
      		   strStatusBack = 0     		   
		End If

	
	cnrtlValue = CInt(Request.Form("processControls"))

	If cnrtlValue = 0 Then
		strStatus = strStatusForward
																
	ElseIf cnrtlValue = 1 Then
		strStatus = strStatusBack
																	
	ElseIf cnrtlValue = 2 Then
		strStatus = 1
																		
	ElseIf cnrtlValue = 3 Then	
		strStatus = currentStatus
																		
	End If

strProjectStatus = "SELECT * FROM ProjectStatus WHERE seq_no = " & CInt(strStatus)

Set rsProjectStatus = Server.CreateObject("ADODB.RecordSet")
rsProjectStatus.Open strProjectStatus, dbconTmp, 3, 3
tmpUpdatedStatusCode = rsProjectStatus.Fields("code")
rsProjectStatus.Close
Set rsProjectStatus = Nothing

strProject = "SELECT * FROM Project WHERE project_id = " & pID
Set rsProject = Server.CreateObject("ADODB.RecordSet")
rsProject.Open strProject, dbconTmp, 3, 3
rsProject.Fields("status_code") = tmpUpdatedStatusCode
rsProject.Update


	rsProject.Close
	Set rsProject = Nothing
	rs.Close
	Set rs = Nothing
	dbconTmp.Close
	Set dbconTmp = Nothing
	
'response.write "currentstatus = " & currentStatus & "<br><br>" 	
'response.write strSQL & "<br><br>"
'response.write strProjectStatus

Response.Redirect("adm_cntlPanel4.asp?pID="& pID)



%>














