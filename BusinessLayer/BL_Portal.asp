<%
Const ERR_NO_REQ_DATA = "errnrd_3rr37efv"
Const REQSTAT_TENT = "TENTATIVE"

' Call the DB function to get Project Descriptions
'***************************************************************************************************
Function fn_BL_GetProjectDescriptions(lngAcctId, arryTmp)

   fn_BL_GetProjectDescriptions = -1
   
   Dim lngRC
   
   lngRC = fn_DL_GetProjectDescriptions(lngAcctId, arryTmp)
       
   If lngRC <> 0 Then
      fn_BL_GetProjectDescriptions = lngRC
      Exit Function
   End If
   
   fn_BL_GetProjectDescriptions = 0

End Function


' Call the DB function to get Project Info
'***************************************************************************************************
Function fn_BL_GetProjectInfo(lngProjectId, arryTmp1, arryTmp2)
   
   Dim lngRC
   
   lngRC = fn_DL_GetProjectInfo(lngProjectId, arryTmp1, arryTmp2)
   If lngRC <> 0 Then
      fn_BL_GetProjectInfo = lngRC
      Exit Function
   End If

   arryTmp1(6,0) = fn_GetDateAsYYYYMMDD(arryTmp1(6,0))  'Project Start Date
         
   If Not IsNull(arryTmp1(3,0)) Then
        arryTmp1(3,0) = fn_GetDateAsMMDDYYYY(arryTmp1(3,0))  ' date MOU was Signed
   End If
   If Not IsNull(arryTmp1(7,0) ) Then
        arryTmp1(7,0) = fn_GetDateAsMMDDYYYY(arryTmp1(7,0) ) ' Project Completion Date
   End If
   If Not IsNull(arryTmp1(8,0) ) Then
        arryTmp1(8,0) = fn_GetDateAsMMDDYYYY(arryTmp1(8,0) )  ' Project Target Date
   End If   
   If arryTmp1(1,0) = True Then
      arryTmp1(1,0) = "Yes"
   Else
      arryTmp1(1,0) = "No"   
   End If 
   
   
   If IsArray(arryTmp2) Then 
       Dim i
       For i = 0 to Ubound(arryTmp2, 2) 
          If IsNull(arryTmp2(2,i)) Then
             arryTmp2(2,i) = 0
          End If
       Next
   Else
        Call sub_HandleLogicError(ERR_NO_REQ_DATA)
   End If   
   
   
   fn_BL_GetProjectInfo = 0

End Function


' Call the DB function to Update the First Time Indicator
'***************************************************************************************************
Function fn_BL_UpdateFirstTime(lngProjectId)

   fn_BL_UpdateFirstTime = -1
   
   Dim lngRC
   
   lngRC = fn_DL_UpdateFirstTime(lngProjectId)
       
   If lngRC <> 0 Then
      fn_BL_UpdateFirstTime = lngRC
      Exit Function
   End If
   
   fn_BL_UpdateFirstTime = 0

End Function


' Call the DB function to Update the Project Status
'***************************************************************************************************
Function fn_BL_UpdateProjectStatus(lngProjectId, strStatusCode)
   
   Dim lngRC, arryTmp, strOldStatusCode
   
   ' Get Old Status
   '*******************************************************
   lngRC = fn_DL_GetProjectStatus(lngProjectId, arryTmp)
   If lngRC <> 0 Then
      fn_DL_GetProjectStatus = lngRC
      Exit Function
   End If
   
   strOldStatusCode = arryTmp(0,0)
   
   ' Update Status
   '*******************************************************   
   lngRC = fn_DL_UpdateProjectStatus(lngProjectId, strStatusCode)
   If lngRC <> 0 Then
      fn_DL_UpdateProjectStatus = lngRC
      Exit Function
   End If
   
   ' Insert Status History
   '*******************************************************   
   lngRC = fn_DL_InsertProjectStatusHistory(lngProjectId, strOldStatusCode, strStatusCode)
   If lngRC <> 0 Then
      fn_DL_InsertProjectStatusHistory = lngRC
      Exit Function
   End If
   
   fn_BL_UpdateProjectStatus = 0

End Function


' Call the DB function to get Project Status Blurb
'***************************************************************************************************
Function fn_BL_GetProjectStatusBlurb(strStatusCode, strBlurb)
   
   Dim lngRC
   Dim arryTmp
   
   lngRC = fn_DL_GetProjectStatusBlurb(strStatusCode, arryTmp)
   If lngRC <> 0 Then
      fn_BL_GetProjectStatusBlurb = lngRC
      Exit Function
   End If
   
   strBlurb = arryTmp(0,0)
   
   Erase arryTmp
   
   fn_BL_GetProjectStatusBlurb = 0

End Function


' Call the DB function to get Project Questions
'***************************************************************************************************
Function fn_BL_GetProjectQuestions(lngProjectId, arryTmp)
   
   Dim lngRC
   
   lngRC = fn_DL_GetProjectQuestions(lngProjectId, arryTmp)
       
   If lngRC <> 0 Then
      fn_BL_GetProjectQuestions = lngRC
      Exit Function
   End If
   
   fn_BL_GetProjectQuestions = 0

End Function


' Insert answers into DB
'***************************************************************************************************
Function fn_BL_InsertProjectAnswers(lngProjectId, arryQuestions, arryAns)

   Dim lngRC
   Dim i

   'lngRC = fn_DL_DeleteProjectQuestions(lngProjectId)
   'If lngRC <> 0 Then
   '   fn_BL_InsertProjectAnswers = lngRC
   '   Exit Function
   'End If
   
   For i = 0 to UBound(arryAns)
     If IsNull(arryQuestions(4,i)) or arryQuestions(4,i) = "" Then
        lngRC = fn_DL_InsertProjectAnswer(lngProjectId, arryQuestions(1,i), arryAns(i))    
     Else
        lngRC = fn_DL_UpdateProjectAnswer(lngProjectId, arryQuestions(1,i), arryAns(i))
     End If
     
     If lngRC <> 0 Then
        fn_BL_InsertProjectAnswers = lngRC
        Exit Function
     End If
   Next
   
   Erase arryAns
   Erase arryQuestions
   
   fn_BL_InsertProjectAnswers = 0

End Function


' Call the DB function to get Project Requirements Descrips
'***************************************************************************************************
Function fn_BL_GetRequirementDescriptions(lngProjectId, arryTmp)
   
   Dim lngRC
   lngRC = fn_DL_GetRequirementDescriptions (lngProjectId, arryTmp)    
   If lngRC <> 0 Then
      fn_BL_GetRequirementDescriptions = lngRC
      Exit Function
   End If
   
   fn_BL_GetRequirementDescriptions = 0

End Function


' Call the DB function to get Project Requirement Detail
'***************************************************************************************************
Function fn_BL_GetRequirementDetail(lngReqId, arryTmp)
   
   Dim lngRC
   lngRC = fn_DL_GetRequirementDetail (lngReqId, arryTmp)    
   If lngRC <> 0 Then
      fn_BL_GetRequirementDetail = lngRC
      Exit Function
   End If
   
   fn_BL_GetRequirementDetail = 0

End Function



' Call the DB function to Update Requirement
'***************************************************************************************************
Function fn_BL_EditRequirement(lngReqId, ByVal strName, ByVal strSummary)
   
   Dim lngRC
   Dim strLUDateTime
   
   strName = fn_EscapeSQL(strName) 
   strSummary = fn_EscapeSQL(strSummary) 
   strLUDateTime = fn_GetNowAsYYYYMMDD_HHMMSS(Now())
   
   lngRC = fn_DL_EditRequirement(lngReqId, strName, strSummary, strLUDateTime)
   If lngRC <> 0 Then
      fn_BL_EditRequirement = lngRC
      Exit Function
   End If
   
   fn_BL_EditRequirement = 0

End Function


' Call the DB function to Delete Requirement
'***************************************************************************************************
Function fn_BL_DeleteRequirement(lngReqId)
   
   Dim lngRC
   
   lngRC = fn_DL_DeleteRequirement(lngReqId)
   If lngRC <> 0 Then
      fn_BL_DeleteRequirement = lngRC
      Exit Function
   End If
   
   fn_BL_DeleteRequirement = 0

End Function


' Call the DB function to Insert a new Requirement
'***************************************************************************************************
Function fn_BL_AddNewRequirement(ByVal lngProjectId, ByVal strName, ByVal strSummary)
   
   Dim lngRC
   Dim strStatusCode
   
   strName = fn_EscapeSQL(strName) 
   strSummary = fn_EscapeSQL(strSummary)
   
   lngRC = fn_DL_AddNewRequirement(lngProjectId, strName, strSummary, REQSTAT_TENT)
   If lngRC <> 0 Then
      fn_BL_AddNewRequirement = lngRC
      Exit Function
   End If
   
   fn_BL_AddNewRequirement = 0

End Function


' Call the DB function to Insert a new Requirement File
'***************************************************************************************************
Function fn_BL_AddNewRequirementFile(ByVal lngReqId, ByVal strDescrip, ByVal strOldFileName, ByVal lngFileSize, ByVal strContentType, strNewFileName)
   
   Dim lngRC
   Dim strStatusCode
   
   lngRC = fn_DL_AddNewRequirementFile(lngReqId, strDescrip, strOldFileName, lngFileSize, strContentType, strNewFileName)
   If lngRC <> 0 Then
      fn_BL_AddNewRequirementFile = lngRC
      Exit Function
   End If
   
   fn_BL_AddNewRequirementFile = 0

End Function


' Call the DB function to get Project Requirements Files
'***************************************************************************************************
Function fn_BL_GetRequirementFiles(lngReqId, arryTmp)
 
   Dim lngRC
   lngRC = fn_DL_GetRequirementFiles (lngReqId, arryTmp)    
   If lngRC <> 0 Then
      fn_BL_GetRequirementFiles = lngRC
      Exit Function
   End If
   
   fn_BL_GetRequirementFiles = 0

End Function


' Call DB function to Delete Requirement File
'***************************************************************************************************
Function fn_BL_DeleteRequirementFile(lngReqId, lngRFId, strFileName)
   
   Dim lngRC
   
   lngRC = fn_DL_DeleteRequirementFile(lngReqId, lngRFId)
   If lngRC <> 0 Then
      fn_BL_DeleteRequirementFile = lngRC
      Exit Function
   End If
   
   lngRC = fn_BL_RemoveUploadedFile(strFileName)
   If lngRC <> 0 Then
      fn_BL_DeleteRequirementFile = lngRC
      Exit Function
   End If
   
   fn_BL_DeleteRequirementFile = 0

End Function

' Remove Uploaded File
'***************************************************************************************************
Function fn_BL_RemoveUploadedFile(strFileName)
   
   Dim fsoTmp
   Dim fileTmp
   
   Set fsoTmp = CreateObject("Scripting.FileSystemObject")
   'PrintStop Replace(Server.MapPath(PAGE_PORTAL),PAGE_PORTAL,"") & FOLDER_UPLOAD & "\" & strFileName
   Set fileTmp = fsoTmp.GetFile(Replace(Server.MapPath(PAGE_PORTAL),PAGE_PORTAL,"") & FOLDER_UPLOAD & "\" & strFileName)

   fileTmp.Delete

   Set fileTmp = Nothing
   Set fsoTmp = Nothing

   fn_BL_RemoveUploadedFile = 0

End Function


' Check to make sure the file belongs to the requirement
'***************************************************************************************************
Function fn_BL_CheckRequirementFile(lngReqId, lngRFId)
   
   Dim lngRC
   Dim arryTmp
   
   lngRC = fn_DL_GetRequirementFile(lngReqId, lngRFId, arryTmp)
   If lngRC <> 0 Then
      fn_BL_CheckRequirementFile = lngRC
      Exit Function
   End If
   
   If Not IsArray(arryTmp) Then
      fn_BL_CheckRequirementFile = ERR_UNAUTH_DOWNLOAD
      Exit Function
   End If
   
   fn_BL_CheckRequirementFile = 0

End Function


' Call the DB function to get Project Issue Types
'***************************************************************************************************
Function fn_BL_GetProjectIssueTypes(arryTmp)
   
   Dim lngRC
   
   lngRC = fn_DL_GetProjectIssueTypes(arryTmp)
       
   If lngRC <> 0 Then
      fn_BL_GetProjectIssueTypes = lngRC
      Exit Function
   End If
   
   fn_BL_GetProjectIssueTypes = 0

End Function


' Call the DB function to get Project Issue Descrips
'***************************************************************************************************
Function fn_BL_GetProjectIssueDescriptions(lngProjectId, strIssueType, arryTmp)
   
   Dim lngRC
   lngRC = fn_DL_GetProjectIssueDescriptions (lngProjectId, strIssueType, arryTmp)    
   If lngRC <> 0 Then
      fn_BL_GetProjectIssueDescriptions = lngRC
      Exit Function
   End If
   
   fn_BL_GetProjectIssueDescriptions = 0

End Function


' Call the DB function to Update Issue
'***************************************************************************************************
Function fn_BL_EditIssue(lngIssueId, strTitle, strDetails, blnResolved)
   
   Dim lngRC
   Dim dtResolved
   
   strTitle = fn_EscapeSQL(strTitle) 
   strDetails = fn_EscapeSQL(strDetails) 
   If blnResolved = 1 Then
      dtResolved = fn_GetNowAsYYYYMMDD_HHMMSS(Now())
   End If
   
   lngRC = fn_DL_EditIssue(lngIssueId, strTitle, strDetails, dtResolved)
   If lngRC <> 0 Then
      fn_BL_EditIssue = lngRC
      Exit Function
   End If
   
   fn_BL_EditIssue = 0

End Function


' Call the DB function to Delete Issue
'***************************************************************************************************
Function fn_BL_DeleteIssue(lngIssueId)
   
   Dim lngRC
   
   lngRC = fn_DL_DeleteIssue(lngIssueId)
   If lngRC <> 0 Then
      fn_BL_DeleteIssue = lngRC
      Exit Function
   End If
   
   fn_BL_DeleteIssue = 0

End Function


' Call the DB function to Insert a new Issue
'***************************************************************************************************
Function fn_BL_AddNewIssue(ByVal lngProjectId, ByVal strIssueType, ByVal strTitle, ByVal strDetails)
   
   Dim lngRC
   Dim strStatusCode
   
   strTitle = fn_EscapeSQL(strTitle) 
   strDetails = fn_EscapeSQL(strDetails)
   
   lngRC = fn_DL_AddNewIssue(lngProjectId, strIssueType, strTitle, strDetails)
   If lngRC <> 0 Then
      fn_BL_AddNewIssue = lngRC
      Exit Function
   End If
   
   fn_BL_AddNewIssue = 0

End Function


' Call the DB function to get Issue Detail
'***************************************************************************************************
Function fn_BL_GetIssueDetail(lngIssueId, arryTmp)
   
   Dim lngRC
   lngRC = fn_DL_GetIssueDetail (lngIssueId, arryTmp)    
   If lngRC <> 0 Then
      fn_BL_GetIssueDetail = lngRC
      Exit Function
   End If
   
   fn_BL_GetIssueDetail = 0

End Function

%>





















