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
<!--#include file="Utility/Incrypt.asp"-->
<%
  If Session("blnAdminLoggedIn") <> True Then
     Response.Redirect "Error.asp?x1=1111"
  End If

Dim aID, strMsg, strFrom
Dim lngRC, arryBillInfo
Dim billinginfo_id, business_name, first_name, last_name, phone1, phone2, email, address, city, state_code, postal_code, state_descrip, country_descrip

strMsg=""
strFrom = Request("txtFrom")

aID = Request.QueryString("aID")

If strFrom = "editbillinfo" Then
    Session("sess_strMsg") = ""
    billinginfo_id = Request.Form("txtbillinginfo_id")
    business_name = Request.Form("txtbusiness_name")
    first_name = Request.Form("txtfirst_name")
    last_name = Request.Form("txtlast_name")
    phone1 = Request.Form("txtphone1")
    phone2 = Request.Form("txtphone2")
    email = Request.Form("txtemail")
    address = Request.Form("txtaddress")
    city = Request.Form("txtcity")
    state_code = Request.Form("txtstate_code")
    postal_code = Request.Form("txtpostal_code")
    state_descrip = Request.Form("txtstate_descrip")
    country_descrip = Request.Form("txtcountry_descrip")
    
    'Response.Write "username = " & strUserName
    'Response.End
   

       lngRC = fn_BL_EditBillInfo(billinginfo_id, business_name, first_name, last_name, phone1, phone2, email, address, city, state_code, postal_code, state_descrip, country_descrip)
       If lngRC = 0 Then
          
          Session("sess_strMsg") = "BillingInfo Changed Successfully! - " & Now    
       Else
          Session("sess_strMsg") =  "Problem Editing UserName and Password!" & lngRC
       End If
        
     
    Response.Redirect "admProject_EditBillInfo.asp?txtFrom=verify&aID=" & aID
    
ElseIf strFrom = "verify" Then
    strMsg = Session("sess_strMsg") 
    lngRC = fn_BL_GetBillingInfo(aID, arryBillInfo)
    If lngRC <> 0 Then
       strMsg = "problem getting billing info"
    End If   
Else
    Session("sess_strMsg") = ""
    lngRC = fn_BL_GetBillingInfo(aID, arryBillInfo)
    If lngRC <> 0 Then
       strMsg = "problem getting billing info"
    End If   
End If

%>
<html>
<head>

</head>

<body bgcolor="#ededed">

<%
Response.Write "<br><br><center>" & strMsg & "</center>"
%>

<form method="post" action="admProject_EditBillInfo.asp?aID=<%= aID %>" name="theForm">
<br><br><br><br><br><br>
<input type="hidden" name="txtFrom" value="editbillinfo">
<%
'SELECT billinginfo_id, business_name, first_name, last_name, phone1, phone2, email, address, city, state_code, postal_code, state_descrip, country_descrip

If IsArray(arryBillInfo) Then
    billinginfo_id = arryBillInfo(0,0)
    business_name = arryBillInfo(1,0)
    first_name = arryBillInfo(2,0)
    last_name = arryBillInfo(3,0)
    phone1 = arryBillInfo(4,0)
    phone2 = arryBillInfo(5,0)
    email = arryBillInfo(6,0)
    address = arryBillInfo(7,0)
    city = arryBillInfo(8,0)
    state_code = arryBillInfo(9,0)
    postal_code = arryBillInfo(10,0)
    state_descrip = arryBillInfo(11,0)
    country_descrip = arryBillInfo(12,0)
End If
%>

<input type="hidden" name="txtbillinginfo_id" value="<% = billinginfo_id %>">
<center>
<table>


	<tr><td><b>Business Name:</b></td><td><input type="text" name="txtbusiness_name" value="<% = business_name %>"></td></tr>
	<tr><td><b>First Name:</b></td><td><input type="text" name="txtfirst_name" value="<% = first_name %>"></td></tr>
	<tr><td><b>Last Name:</b></td><td><input type="text" name="txtlast_name" value="<% = last_name %>"></td></tr>
	<tr><td><b>Phone 1:</b></td><td><input type="text" name="txtphone1" value="<% = phone1 %>"></td></tr>
	<tr><td><b>Phone 2:</b></td><td><input type="text" name="txtphone2" value="<% = phone2 %>"></td></tr>
	<tr><td><b>Email:</b></td><td><input type="text" name="txtemail" value="<% = email %>"></td></tr>
	<tr><td><b>Address:</b></td><td><input type="text" name="txtaddress" value="<% = address %>"></td></tr>
	<tr><td><b>city:</b></td><td><input type="text" name="txtcity" value="<% = city %>"></td></tr>
	<tr><td><b>state code:</b></td><td><input type="text" name="txtstate_code" value="<% = state_code %>"></td></tr>
	<tr><td><b>postal code:</b></td><td><input type="text" name="txtpostal_code" value="<% = postal_code %>"></td></tr>
	<tr><td><b>state/provin descrip(optional):</b></td><td><input type="text" name="txtstate_descrip" value="<% = state_descrip %>"></td></tr>
	<tr><td><b>country descrip(optional):</b></td><td><input type="text" name="txtcountry_descrip" value="<% = country_descrip %>"></td></tr>

	<tr>
	<td colspan=2><center><input type="submit" value="Update BillingInfo"></center></td>
	</tr>

</table>
</center>
</form>

</body>
</html>
<%


' Edit BillingInfo 
'*******************************************************************************
Function fn_BL_EditBillInfo(billinginfo_id, business_name, first_name, last_name, phone1, phone2, email, address, city, state_code, postal_code, state_descrip, country_descrip)
           
 On Error Resume Next

    Dim dbconTmp
    Dim strSQL   
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE BillingInfo"
    
    If business_name = "" Then
       strSQL = strSQL & " " & Chr(13) & "set business_name = null"
    Else
       strSQL = strSQL & " " & Chr(13) & "set business_name = '" & business_name & "'"
    End If
    
    If first_name = "" Then
       strSQL = strSQL & " " & Chr(13) & ", first_name = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", first_name = '" & first_name & "'"
    End If
    
    If last_name = "" Then
       strSQL = strSQL & " " & Chr(13) & ", last_name = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", last_name = '" & last_name & "'"
    End If
    
    If phone1 = "" Then
       strSQL = strSQL & " " & Chr(13) & ", phone1 = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", phone1 = '" & phone1 & "'"
    End If
    
    If phone2 = "" Then
       strSQL = strSQL & " " & Chr(13) & ", phone2 = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", phone2 = '" & phone2 & "'"
    End If
    
    
    If email = "" Then
       strSQL = strSQL & " " & Chr(13) & ", email = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", email = '" & email & "'"
    End If
        
    If address = "" Then
       strSQL = strSQL & " " & Chr(13) & ", address = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", address = '" & address & "'"
    End If
    
    If city = "" Then
       strSQL = strSQL & " " & Chr(13) & ", city = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", city = '" & city & "'"
    End If        
        
    If state_code = "" Then
       strSQL = strSQL & " " & Chr(13) & ", state_code = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", state_code = '" & state_code & "'"
    End If
    
    If postal_code = "" Then
       strSQL = strSQL & " " & Chr(13) & ", postal_code = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", postal_code = '" & postal_code & "'"
    End If
    
    If state_descrip = "" Then
       strSQL = strSQL & " " & Chr(13) & ", state_descrip = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", state_descrip = '" & state_descrip & "'"
    End If
    
    If country_descrip = "" Then
       strSQL = strSQL & " " & Chr(13) & ", country_descrip = null"
    Else
       strSQL = strSQL & " " & Chr(13) & ", country_descrip = '" & country_descrip & "'"
    End If
    
    
    strSQL = strSQL & " " & Chr(13) & "where billinginfo_id = " & billinginfo_id 

    'Response.Write "strSQL = " & strSQL
    'Response.End

    dbconTmp.Execute (strSQL)

     
    If Err.Number <> 0 Then
       fn_BL_EditBillInfo = Err.Number
       dbconTmp.RollbackTrans
       dbconTmp.Close 
       Set dbconTmp = Nothing
       Exit Function 
    End If
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
     
   fn_BL_EditBillInfo = 0

End Function 'fn_BL_EditBillInfo 



' Get Billing Info for an Account
'*******************************************************************************
Function fn_BL_GetBillingInfo(aID, arryTmp)

On Error Resume Next

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL   
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    Set rsTmp = Server.CreateObject("ADODB.Recordset")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = "SELECT billinginfo_id, business_name, first_name, last_name, phone1, phone2, email, address, city, state_code, postal_code, state_descrip, country_descrip"
    strSQL = strSQL & " from BillingInfo where account_id = " & aID

    Set rsTmp = dbconTmp.Execute (strSQL)
    
    If Not rsTmp.EOF Then
    
       arryTmp = rsTmp.GetRows
   
    End If 
    
    rsTmp.Close
      
    If Err.Number <> 0 Then
       fn_BL_GetBillingInfo = Err.Number
       dbconTmp.Close
       Set rsTmp = Nothing
       Set dbconTmp = Nothing
       Exit Function 
    End If
    
    dbconTmp.Close
    Set rsTmp = Nothing
    Set dbconTmp = Nothing   
     
    fn_BL_GetBillingInfo = 0

End Function 'fn_BL_GetBillingInfo

%>