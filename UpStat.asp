<%@
LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!--#include file="./Utility/Util.asp" -->
<!--#include file="./BusinessLayer/BL_Portal.asp" -->
<%
  Const STATUS_PEND = "PEND"
  Const STATUS_NEW = "NEW"
  Const STATUS_REQGATH = "REQGATH"
  
  Dim lngRC, lngProjectId, strStatusCode
  lngProjectId = Request("x1")
  strStatusCode = Request("x2")

  lngRC = fn_BL_UpdateProjectStatus(lngProjectId, strStatusCode)
  If lngRC <> 0 Then
     Call sub_HandleLogicError(ERR_EDIT_USERDATA)
  End If
  
  Response.Write "Status Updated Succesfully to " & strStatusCode & "! Way To Go Mo Fo ! " & Now()

%>