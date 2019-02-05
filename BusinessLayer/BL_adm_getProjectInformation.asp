<!-- #include file="../DataLayer/DL_adm_getProjectInformation.asp"-->
<%

'***************************************************************************************************
Function fn_BL_GetProjectInfo(lngProjectId, arryTmp1, arryTmp2)
   
   Dim lngRC
   
   lngRC = fn_DL_GetProjectInfo(lngProjectId, arryTmp1, arryTmp2)
   If lngRC <> 0 Then
      fn_BL_GetProjectInfo = lngRC
      Exit Function
   End If

'   arryTmp1(6,0) = fn_GetDateAsYYYYMMDD(arryTmp1(6,0))  'Project Start Date
         
'   If Not IsNull(arryTmp1(3,0)) Then
'        arryTmp1(3,0) = fn_GetDateAsMMDDYYYY(arryTmp1(3,0))  ' date MOU was Signed
'   End If
'   If Not IsNull(arryTmp1(7,0) ) Then
'        arryTmp1(7,0) = fn_GetDateAsMMDDYYYY(arryTmp1(7,0) ) ' Project Completion Date
'   End If
'   If Not IsNull(arryTmp1(8,0) ) Then
'        arryTmp1(8,0) = fn_GetDateAsMMDDYYYY(arryTmp1(8,0) )  ' Project Target Date
'   End If   
'   If arryTmp1(1,0) = True Then
'      arryTmp1(1,0) = "Yes"
'   Else
'      arryTmp1(1,0) = "No"   
'   End If 
   
   
'   If IsArray(arryTmp2) Then 
'       Dim i
'       For i = 0 to Ubound(arryTmp2, 2) 
'          If IsNull(arryTmp2(2,i)) Then
'             arryTmp2(2,i) = 0
'          End If
'       Next
'   Else
'        Call sub_HandleLogicError(ERR_NO_REQ_DATA)
'   End If   
   
   
   fn_BL_GetProjectInfo = 0

End Function


%>