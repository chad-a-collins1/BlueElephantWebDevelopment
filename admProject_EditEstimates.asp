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
     Response.Redirect "Error.asp?x1=1111"
  End If
  Dim arryInfo, project_name, projecttype_code, down_payment, status_code, project_balance, project_estimate_hours, project_estimate_cost, target_date, completion_date, mou_signed, mou_signee, mou_signed_date, first_time, accrued_hours, consultant_id
  Dim pID, lngInvoiceId, strDescrip, strLineItemType, strRate, strQty, curAmount, strDispSeq, strComment, strFrom, lngRC, strMsg
  Dim arryRate

strMsg=""
strFrom = Request("from")
pID = Request.QueryString("pID")

If strFrom = "updateproj" Then
   
   project_name = Request("txtProjName")
   projecttype_code = Request("lstProjType")
   down_payment = Request("txtDownPay")
   project_balance = Request("txtProjBal")
   project_estimate_hours = Request("txtProjEstHours")
   project_estimate_cost = Request("txtProjEstCost")
   target_date = Request("txtTargetDate")
   completion_date = Request("txtCompletionDate")
   mou_signed = Request("chkMOUSigned")
   mou_signee = Request("txtMOUSignee")
   mou_signed_date = Request("txtMOUSignedDate")
   first_time = Request("chkFirstTime")
   accrued_hours = Request("txtAccruedHours")
   consultant_id = Request("txtConsultantId")
  
  
   lngRC = fn_BL_EditProject(pId, project_name, projecttype_code, down_payment, project_balance, project_estimate_hours, project_estimate_cost, target_date, completion_date, mou_signed, mou_signee, mou_signed_date, first_time, accrued_hours, consultant_id)
   If lngRC <> 0 Then
      Response.Redirect "Error.asp?x1=" & lngRC
   End If
   
    Session("sess_strMsg") = "Project Updated Suckdickfully Bizzatch!"
    Response.Redirect "admProject_EditEstimates.asp?from=verify&pID=" & pId
    
ElseIf strFrom = "verify" Then
    strMsg = Session("sess_strMsg") 
End If

%>
<html>
<head>

</head>

<body bgcolor="#ededed">

<%
Response.Write "<br><br><center>" & strMsg & "</center>"
%>
<br>
<form method="post" action="admProject_EditEstimates.asp?from=updateproj&pID=<%= pID %>" name="theForm">
<center>
<table>
<%

    lngRC = fn_BL_GetProjectInfo (pId, arryInfo)
    If lngRC <> 0 Then
         Response.Redirect "Error.asp?x1=" & lngRC
    End If

    project_name = arryInfo(0,0)
    projecttype_code = arryInfo(1,0)
    down_payment = arryInfo(2,0)
    status_code = arryInfo(3,0)
    project_balance = arryInfo(4,0)
    project_estimate_hours = arryInfo(5,0)
    project_estimate_cost = arryInfo(6,0)
    target_date = arryInfo(7,0)
    completion_date = arryInfo(8,0)
    mou_signed = arryInfo(9,0)
    mou_signee = arryInfo(10,0)
    mou_signed_date = arryInfo(11,0)
    first_time = arryInfo(12,0)
    accrued_hours = arryInfo(13,0)
    consultant_id = arryInfo(14,0)
    
    
    'Response.Write "<br>project_name = " & project_name
    'Response.Write "<br>projecttype_code = " & projecttype_code
    'Response.Write "<br>down_payment = " & down_payment
    'Response.Write "<br>status_code = " & status_code
    'Response.Write "<br>project_balance = " & project_balance
    'Response.Write "<br>project_estimate_hours = " & project_estimate_hours
    'Response.Write "<br>project_estimate_cost = " & project_estimate_cost
    'Response.Write "<br>target_date = " & target_date
    'Response.Write "<br>completion_date = " & completion_date
    'Response.Write "<br>mou_signed = " & mou_signed
    'Response.Write "<br>mou_signee = " & mou_signee
    'Response.Write "<br>mou_signed_date = " & mou_signed_date
    'Response.Write "<br>first_time = " & first_time
    'Response.Write "<br>accrued_hours = " & accrued_hours
    'Response.Write "<br>consultant_id = " & consultant_id

        
%>


	<tr>
	<td>
	<B>Project Name: </b><input type=text size=30 maxlength=50 name="txtProjName" value="<% = project_name %>"></td>
	<td></td>
	</tr>

	<tr>
	<td><b>Project Type: </b>
	   <SELECT name="lstProjType">
           <%   
              Dim arryProjTypeCodes, arryProjTypeDesc, i
              arryProjTypeCodes = Split("DATABASE|OTHER|WEBAPP|WIN32APP","|")
              arryProjTypeDesc = Split("Database Development|Other|Web Application|Win32 App","|")
              For i = 0 To Ubound(arryProjTypeCodes) 
                 If projecttype_code = arryProjTypeCodes(i) Then
	            Response.Write "<OPTION Value=""" & arryProjTypeCodes(i) & """ selected>" & arryProjTypeDesc(i) & "</OPTION>"
	         Else
	            Response.Write "<OPTION Value=""" & arryProjTypeCodes(i) & """ >" & arryProjTypeDesc(i) & "</OPTION>"
	         End If
	      Next
	   %>
	    </SELECT>
	  </td>
	</tr>
	
	<tr>
	<td colspan=2><b>Down Payment: </b><input type=text size=15 maxlength=25 name="txtDownPay" value="<% = down_payment %>">
	</tr>
	<tr>
	<td colspan=2><b>Project Balance: </b><input type=text size=15 maxlength=25 name="txtProjBal" value="<% = project_balance %>">
	</tr>
	<tr>
	<td colspan=2><b>Project Estimated Hours: </b><input type=text size=15 maxlength=25 name="txtProjEstHours" value="<% = project_estimate_hours %>">
	</tr>
	<tr>
	<td colspan=2><b>Project Estimated Cost: </b><input type=text size=15 maxlength=25 name="txtProjEstCost" value="<% = project_estimate_cost %>">
	</tr>	
	<tr>
	<td colspan=2><b>Target Date: </b><input type=text size=15 maxlength=25 name="txtTargetDate" value="<% = target_date %>">
	</tr>		
	<tr>
	<td colspan=2><b>Completion Date: </b><input type=text size=15 maxlength=25 name="txtCompletionDate" value="<% = completion_date %>">
	</tr>
	<tr>
	<% 
	Dim strChecked
	 strChecked = ""
	   If mou_signed Then
	       strChecked = "checked"
	   End If
	%>
	<td colspan=2><b>MOU Signed: </b><input type="checkbox" size=15 maxlength=25 name="chkMOUSigned" value="1" <% = strChecked %>>
	</tr>
	<tr>
	<td colspan=2><b>MOU Signee: </b><input type=text size=15 maxlength=25 name="txtMOUSignee" value="<% = mou_signee %>">
	</tr>
	<tr>
	<td colspan=2><b>MOU Signed Date: </b><input type=text size=15 maxlength=25 name="txtMOUSignedDate" value="<% = mou_signed_date %>">
	</tr>
	<tr>
	<% 
	 strChecked = ""
	   If first_time Then
	       strChecked = "checked"
	   End If
	%>
	<td colspan=2><b>First Time: </b><input type="checkbox" size=15 maxlength=25 name="chkFirstTime" value="1" <% = strChecked %>>
	</tr>
	<tr>
	<td colspan=2><b>Accrued Hours: </b><input type=text size=15 maxlength=25 name="txtAccruedHours" value="<% = accrued_hours %>">
	</tr>
	<tr>
	<td colspan=2><b>Consultant Id: </b><input type=text size=15 maxlength=25 name="txtConsultantId" value="<% = consultant_id %>">
	</tr>						

	<tr>
	<td colspan=2><center><input name="cmdSubmit" type="submit" value="Update"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>


<%
' Get Project Info
'***************************************************************************************************
Function fn_BL_GetProjectInfo (pId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "p.project_name,"
    strSQL = strSQL & " " & Chr(13) & "p.projecttype_code,"
    strSQL = strSQL & " " & Chr(13) & "p.down_payment,"
    strSQL = strSQL & " " & Chr(13) & "p.status_code,"
    strSQL = strSQL & " " & Chr(13) & "p.project_balance,"
    strSQL = strSQL & " " & Chr(13) & "p.project_estimate_hours,"
    strSQL = strSQL & " " & Chr(13) & "p.project_estimate_cost,"
    strSQL = strSQL & " " & Chr(13) & "p.target_date,"
    strSQL = strSQL & " " & Chr(13) & "p.completion_date,"
    strSQL = strSQL & " " & Chr(13) & "p.mou_signed,"
    strSQL = strSQL & " " & Chr(13) & "p.mou_signee,"
    strSQL = strSQL & " " & Chr(13) & "p.mou_signed_date,"
    strSQL = strSQL & " " & Chr(13) & "p.first_time,"
    strSQL = strSQL & " " & Chr(13) & "p.accrued_hours,"
    strSQL = strSQL & " " & Chr(13) & "p.consultant_id"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "Project p" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "p.project_id = " & pId
        
    Set rsTmp = dbconTmp.Execute(strSQL)
    
    With rsTmp
       If Not .EOF and Not.BOF Then
          arryTmp = .GetRows
       End If
       .Close
    End With  'rsTmp
    
    Set rsTmp = Nothing
    dbconTmp.Close
    Set dbconTmp = Nothing
    
    fn_BL_GetProjectInfo = 0

End Function   'fn_BL_GetProjectInfo




Function fn_BL_EditProject (pId, project_name, projecttype_code, down_payment, project_balance, project_estimate_hours, project_estimate_cost, target_date, completion_date, mou_signed, mou_signee, mou_signed_date, first_time, accrued_hours, consultant_id)
   
    Dim dbconTmp
    Dim rs
    Dim strSQL
   
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE Project"   
    If project_name = "" Then
       strSQL = strSQL & " " & Chr(13) & "SET project_name = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "SET project_name = '" & project_name & "'," 
    End If          
    strSQL = strSQL & " " & Chr(13) & "projecttype_code = '" & projecttype_code & "',"
    If down_payment = "" Then
       strSQL = strSQL & " " & Chr(13) & "down_payment = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "down_payment = " & down_payment & "," 
    End If  
    strSQL = strSQL & " " & Chr(13) & "project_balance = " & project_balance & ","
    If project_estimate_hours = "" Then
       strSQL = strSQL & " " & Chr(13) & "project_estimate_hours = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "project_estimate_hours = " & project_estimate_hours & "," 
    End If   
    If project_estimate_cost = "" Then
       strSQL = strSQL & " " & Chr(13) & "project_estimate_cost = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "project_estimate_cost = " & project_estimate_cost & "," 
    End If   
    If target_date = "" Then
       strSQL = strSQL & " " & Chr(13) & "target_date = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "target_date = '" & target_date & "'," 
    End If  
    If completion_date = "" Then
       strSQL = strSQL & " " & Chr(13) & "completion_date = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "completion_date = '" & completion_date & "'," 
    End If                     

    If mou_signed = 1 Then
       mou_signed = -1
    Else
       mou_signed = 0   
    End If
    strSQL = strSQL & " " & Chr(13) & "mou_signed = " & mou_signed & "," 
    If mou_signee = "" Then
       strSQL = strSQL & " " & Chr(13) & "mou_signee = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "mou_signee = '" & mou_signee & "'," 
    End If 
    If mou_signed_date = "" Then
       strSQL = strSQL & " " & Chr(13) & "mou_signed_date = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "mou_signed_date = '" & mou_signed_date & "'," 
    End If  
    If first_time = 1 Then
       first_time = -1
    Else
       first_time = 0   
    End If        
    strSQL = strSQL & " " & Chr(13) & "first_time = " & first_time & "," 
    If accrued_hours = "" Then
       strSQL = strSQL & " " & Chr(13) & "accrued_hours = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "accrued_hours = " & accrued_hours & "," 
    End If  
    If consultant_id = "" Then
       strSQL = strSQL & " " & Chr(13) & "consultant_id = NULL," 
    Else
       strSQL = strSQL & " " & Chr(13) & "consultant_id = " & consultant_id & "" 
    End If          
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "project_id = " & pId    
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    dbconTmp.Execute (strSQL)

    dbconTmp.Close
    Set dbconTmp = Nothing

   
    fn_BL_EditProject = 0
   
End Function
%>