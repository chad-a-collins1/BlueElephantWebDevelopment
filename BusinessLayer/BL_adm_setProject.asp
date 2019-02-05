<!-- #include file="../DataLayer/DL_adm_setProject.asp" -->
<%
Function fn_BL_admProjectInsert(AccountID, ProjectTypeCode, StatusCode, intConId)

   'response.write "<BR>" & ProjectTypeCode
   'response.write "<BR>" & StatusCode

   Dim lngRC

   lngRC = fn_DL_admProjectInsert(AccountID, ProjectTypeCode, StatusCode, intConId)
   If lngRC <> 0 Then
      fn_BL_admProjectInsert = lngRC
      Exit Function
   End If	

   fn_BL_admProjectInsert = 0

End Function

%>