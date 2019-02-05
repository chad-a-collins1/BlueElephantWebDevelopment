<%

' Get Project Descriptions
'***************************************************************************************************
Function fn_DL_GetProjectDescriptions(lngAcctId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "p.project_id,"
    strSQL = strSQL & " " & Chr(13) & "p.project_name"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "Project p" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "p.account_id = " & lngAcctId
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetProjectDescriptions = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetProjectDescriptions = 0

End Function   'fn_DL_GetProjectDescriptions


' Get Project Status Blurb
'***************************************************************************************************
Function fn_DL_GetProjectStatusBlurb(strStatusCode, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "ps.blurb"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "ProjectStatus ps" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "ps.code = '" & strStatusCode & "'"
        
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
     
    'If Err.Number <> 0 Then
    '   fn_DL_GetProjectStatusBlurb = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetProjectStatusBlurb = 0

End Function   'fn_DL_GetProjectStatusBlurb


' Get Project Info
'***************************************************************************************************
Function fn_DL_GetProjectInfo(lngProjectId, arryTmp1, arryTmp2)

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
    strSQL = strSQL & " " & Chr(13) & "p.mou_signed," 
    strSQL = strSQL & " " & Chr(13) & "p.mou_signee,"    
    strSQL = strSQL & " " & Chr(13) & "p.mou_signed_date,"          
    strSQL = strSQL & " " & Chr(13) & "p.down_payment," 
    strSQL = strSQL & " " & Chr(13) & "ps.descrip as status_descrip," 
    strSQL = strSQL & " " & Chr(13) & "p.insert_datetime," 
    strSQL = strSQL & " " & Chr(13) & "p.completion_date," 
    strSQL = strSQL & " " & Chr(13) & "p.target_date," 
    strSQL = strSQL & " " & Chr(13) & "p.project_balance," 
    strSQL = strSQL & " " & Chr(13) & "p.project_estimate_hours," 
    strSQL = strSQL & " " & Chr(13) & "p.project_estimate_cost,"
    strSQL = strSQL & " " & Chr(13) & "p.first_time,"
    strSQL = strSQL & " " & Chr(13) & "p.status_code,"
    strSQL = strSQL & " " & Chr(13) & "ps.freeze,"
    strSQL = strSQL & " " & Chr(13) & "ps.lockedit"                 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM" 
    strSQL = strSQL & " " & Chr(13) & "(Project p" 
    strSQL = strSQL & " " & Chr(13) & "INNER JOIN ProjectStatus ps ON p.status_code = ps.code)" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE" 
    strSQL = strSQL & " " & Chr(13) & "p.project_id = " & lngProjectId
    
    Set rsTmp = dbconTmp.Execute(strSQL)
   
    With rsTmp
       If Not .EOF and Not.BOF Then
          arryTmp1 = .GetRows
       End If
       .Close
    End With  'rsTmp
    
    Set rsTmp = Nothing
   
    Dim cmdTmp 
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "qry_GetRequirementCount"
       .CommandType = adCmdStoredProc

       .Parameters.Append .CreateParameter("@lngProjectId", adInteger, adParamInput, , lngProjectId)
       
       Set rsTmp = .Execute  
    End With 'cmdTmp
    
   Dim j 
   j = 0
   
   '****************************************************
   With rsTmp
      If Not .EOF Then
         If j = 0 Then
            'arrays are structured as the following -> arry(col , row)
            arryTmp2 = .GetRows()
         End If  
      End If
   End With   'rsTmp    
    
    Set rsTmp = Nothing
    Set cmdTmp = Nothing    
    dbconTmp.Close
    Set dbconTmp = Nothing

    
    'If Err.Number <> 0 Then
    '   fn_DL_GetProjectInfo = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetProjectInfo = 0      

End Function


' Delete Project Answers
'***************************************************************************************************
Function fn_DL_DeleteProjectAnswers(lngProjectId)

'On Error Resume Next

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "DELETE FROM ProjectAnswer WHERE project_id = " & lngProjectId
    
    'printstop "SQL = " & strSQL
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
      
    'If Err.Number <> 0 Then
    '   fn_DL_DeleteProjectAnswers = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_DeleteProjectAnswers = 0

End Function   'fn_DL_DeleteProjectAnswers



' Get Project Status
'***************************************************************************************************
Function fn_DL_GetProjectStatus(lngProjectId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "p.status_code,"
    strSQL = strSQL & " " & Chr(13) & "ps.descrip"    
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "Project p, ProjectStatus ps " 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "p.project_id = " & lngProjectId
    strSQL = strSQL & " " & Chr(13) & "and ps.code = p.status_code"     
        
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
     
    'If Err.Number <> 0 Then
    '   fn_DL_GetProjectStatus = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetProjectStatus = 0

End Function   'fn_DL_GetProjectStatus



' fn_DL_InsertProjectAnswer
'***************************************************************************************************
Function fn_DL_InsertProjectAnswer(lngProjectId, strQuestionCode, strAnswer)

'On Error Resume Next

    fn_DL_InsertProjectAnswer = -1

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "INSERT INTO ProjectAnswer (project_id, question_code, answer, last_update_datetime)"
    strSQL = strSQL & " VALUES (" & lngProjectId & ", '" & strQuestionCode & "', '" & strAnswer & "', '" & Now & "')"   'fn_GetNowAsYYYYMMDD_HHHMMSS
    
    'printstop "SQL = " & strSQL
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
      
    'If Err.Number <> 0 Then
    '   fn_DL_InsertProjectAnswer = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_InsertProjectAnswer = 0

End Function   'fn_DL_InsertProjectAnswer



' fn_DL_InsertProjectStatusHistory
'***************************************************************************************************
Function fn_DL_InsertProjectStatusHistory(lngProjectId, strOldStatusCode, strStatusCode)

'On Error Resume Next


    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "INSERT INTO ProjectStatusHistory (project_id, old_status_code, new_status_code)"
    strSQL = strSQL & " VALUES (" & lngProjectId & ", '" & strOldStatusCode & "', '" & strStatusCode & "')" 
    
    'printstop "SQL = " & strSQL
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
      
    'If Err.Number <> 0 Then
    '   fn_DL_InsertProjectStatusHistory = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_InsertProjectStatusHistory = 0

End Function   'fn_DL_InsertProjectStatusHistory



' fn_DL_InsertProjectAnswer
'***************************************************************************************************
Function fn_DL_UpdateProjectAnswer(lngProjectId, strQuestionCode, strAnswer)

'On Error Resume Next

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE ProjectAnswer"       
    strSQL = strSQL & " " & Chr(13) & "SET answer = '" & strAnswer & "'," 
    strSQL = strSQL & " " & Chr(13) & "last_update_datetime = '" & Now() & "'"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "project_id = " & lngProjectId    
    strSQL = strSQL & " " & Chr(13) & "and question_code = '" & strQuestionCode & "'"   
            
    'printstop "SQL = " & strSQL
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
      
    'If Err.Number <> 0 Then
    '   fn_DL_UpdateProjectAnswer = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_UpdateProjectAnswer = 0

End Function   'fn_DL_UpdateProjectAnswer



' Update First Time indicator
'***************************************************************************************************
Function fn_DL_UpdateFirstTime(lngProjectId)

'On Error Resume Next

    fn_DL_UpdateFirstTime = -1

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "UPDATE Project set first_time = 0 WHERE project_id = " & lngProjectId
    
    'Response.Write "SQL = " & strSQL
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
      
    'If Err.Number <> 0 Then
    '   fn_DL_UpdateFirstTime = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_UpdateFirstTime = 0

End Function   'fn_DL_UpdateFirstTime


' Update Project Status
'***************************************************************************************************
Function fn_DL_UpdateProjectStatus(lngProjectId, strStatusCode)

'On Error Resume Next

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "UPDATE Project set status_code = '" & strStatusCode & "' WHERE project_id = " & lngProjectId
    
    'Response.Write "SQL = " & strSQL
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing 
      
    'If Err.Number <> 0 Then
    '   fn_DL_UpdateProjectStatus = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_UpdateProjectStatus = 0

End Function   'fn_DL_UpdateProjectStatus


' Get Project Questions
'***************************************************************************************************
Function fn_DL_GetProjectQuestions(lngProjectId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    'strSQL = ""
    'strSQL = strSQL & " " & Chr(13) & "SELECT"
    'strSQL = strSQL & " " & Chr(13) & "pq.question,"     
    'strSQL = strSQL & " " & Chr(13) & "pq.code," 
    'strSQL = strSQL & " " & Chr(13) & "pq.display_seq," 
    'strSQL = strSQL & " " & Chr(13) & "pa.answer"
    'strSQL = strSQL & " " & Chr(13) & ""
    'strSQL = strSQL & " " & Chr(13) & "FROM" 
    'strSQL = strSQL & " " & Chr(13) & "ProjectQuestion pq"
    'strSQL = strSQL & " " & Chr(13) & "LEFT OUTER JOIN"
    'strSQL = strSQL & " " & Chr(13) & "("    
    'strSQL = strSQL & " " & Chr(13) & " select question_code, answer ProjectAnswer where project_id = " & lngProjectId
    'strSQL = strSQL & " " & Chr(13) & ") pa ON pq.code = pa.question_code"
    'strSQL = strSQL & " " & Chr(13) & ""
    'strSQL = strSQL & " " & Chr(13) & "WHERE" 
    'strSQL = strSQL & " " & Chr(13) & "pq.active = True"
    'strSQL = strSQL & " " & Chr(13) & "ORDER BY pq.display_seq"    
        
    Dim cmdTmp 
    Set cmdTmp = Server.CreateObject("ADODB.Command")
    
    With cmdTmp
       .ActiveConnection = fn_GetConnectionString
       .CommandText = "qry_GetProjectQuestions"
       .CommandType = adCmdStoredProc
    
       .Parameters.Append .CreateParameter("@lngProjectId", adInteger, adParamInput, , lngProjectId)
       
       Set rsTmp = .Execute  
    End With 'cmdTmp
    
   Dim j 
   j = 0
   
   '****************************************************
   With rsTmp
       
         If Not .EOF Then
             If j = 0 Then
                ' arrays are structured as the following -> arry(col , row)
                arryTmp = .GetRows()
             End If  
         End If     

   End With   'rsTmp    
    
    Set rsTmp = Nothing
    Set cmdTmp = Nothing 
    
    'If Err.Number <> 0 Then
    '   fn_DL_GetProjectQuestions = Err.Number
    '   Exit Function
    'End If    
    
    fn_DL_GetProjectQuestions = 0

End Function   'fn_DL_GetProjectQuestions


' Get Project Requirement Descriptions
'***************************************************************************************************
Function fn_DL_GetRequirementDescriptions(lngProjectId, arryTmp)

    fn_DL_GetRequirementDescriptions = -1

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "r.requirement_id,"
    strSQL = strSQL & " " & Chr(13) & "r.requirement_name," 
    strSQL = strSQL & " " & Chr(13) & "rs.descrip as status" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "Requirement r"
    strSQL = strSQL & " " & Chr(13) & "INNER JOIN RequirementStatus rs ON r.status_code = rs.code"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "r.project_id = " & lngProjectId
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "ORDER BY rs.display_seq, r.insert_datetime"
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetRequirementDescriptions = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetRequirementDescriptions = 0

End Function   'fn_DL_GetRequirementDescriptions



' Get Requirement Detail
'***************************************************************************************************
Function fn_DL_GetRequirementDetail(lngReqId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "r.requirement_name," 
    strSQL = strSQL & " " & Chr(13) & "r.summary,"
    strSQL = strSQL & " " & Chr(13) & "r.last_update_datetime," 
    strSQL = strSQL & " " & Chr(13) & "r.approved_date,"
    strSQL = strSQL & " " & Chr(13) & "r.chaddan_approved_date,"
    strSQL = strSQL & " " & Chr(13) & "rs.descrip as status," 
    strSQL = strSQL & " " & Chr(13) & "r.completetion_date"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "Requirement r"
    strSQL = strSQL & " " & Chr(13) & "INNER JOIN RequirementStatus rs ON r.status_code = rs.code"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "r.requirement_id = " & lngReqId
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetRequirementDetail = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetRequirementDetail = 0

End Function   'fn_DL_GetRequirementDetail


' Update Requirement
'***************************************************************************************************
Function fn_DL_EditRequirement(ByVal lngReqId, ByVal strName, ByVal strSummary, ByVal strLUDateTime)

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE Requirement" & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "SET requirement_name = '" & strName & "'," & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "summary = '" & strSummary & "'," & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "last_update_datetime = '" & strLUDateTime & "'" & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "WHERE requirement_id = " & lngReqId
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
      
    'If Err.Number <> 0 Then
    '   fn_DL_EditRequirement = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_EditRequirement = 0

End Function


' Delete Requirement
'***************************************************************************************************
Function fn_DL_DeleteRequirement(ByVal lngReqId)

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "DELETE FROM Requirement" & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "WHERE requirement_id = " & lngReqId
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
      
    'If Err.Number <> 0 Then
    '   fn_DL_DeleteRequirement = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_DeleteRequirement = 0

End Function



' Insert Requirement
'***************************************************************************************************
Function fn_DL_AddNewRequirement(ByVal lngProjectId, ByVal strName, ByVal strSummary, ByVal strStatusCode)

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "INSERT INTO Requirement"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "project_id,"
    strSQL = strSQL & " " & Chr(13) & "requirement_name,"
    strSQL = strSQL & " " & Chr(13) & "summary,"
    strSQL = strSQL & " " & Chr(13) & "status_code"
    strSQL = strSQL & " " & Chr(13) & ")"
    strSQL = strSQL & " " & Chr(13) & "VALUES"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "" & lngProjectId & ","
    strSQL = strSQL & " " & Chr(13) & "'" & strName & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strSummary & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & "TENTATIVE" & "'"
    strSQL = strSQL & " " & Chr(13) & ")"
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
      
    'If Err.Number <> 0 Then
    '   fn_DL_AddNewRequirement = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_AddNewRequirement = 0

End Function


' Insert Requirement
'***************************************************************************************************
Function fn_DL_AddNewRequirementFile (ByVal lngReqId, ByVal strDescrip, ByVal strOldFileName, ByVal lngFileSize, ByVal strContentType, strNewFileName)

   Dim rsTmp           'As New ADODB.Recordset
    
   Set rsTmp = Server.CreateObject("ADODB.Recordset")
    '**************************************************
    			
   rsTmp.Open "RequirementFile", fn_GetConnectionString, 2, 2
   rsTmp.AddNew ' create a new record
   
   rsTmp("requirement_id") = lngReqId		
   rsTmp("oldfilename") = strOldFileName
   rsTmp("filesize") = lngFileSize
   rsTmp("contenttype") = strContentType	
   rsTmp("descrip") = strDescrip		
   ' Commit the changes and close
   rsTmp.Update
   strNewFileName = "REQFILE" & "__R" & CStr(lngReqId) & "_F" & CStr(rsTmp("requirementfile_id")) & Mid(strOldFileName,InStrRev(strOldFileName,"."),4)
   rsTmp("newfilename") = strNewFileName
   rsTmp.Update
   rsTmp.Close
    
   Set rsTmp = Nothing
   
   fn_DL_AddNewRequirementFile = 0

End Function


' Get Requirement Detail
'***************************************************************************************************
Function fn_DL_GetRequirementFiles(lngReqId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "rf.requirementfile_id," 
    strSQL = strSQL & " " & Chr(13) & "rf.oldfilename," 
    strSQL = strSQL & " " & Chr(13) & "rf.newfilename," 
    strSQL = strSQL & " " & Chr(13) & "rf.descrip"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "RequirementFile rf"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "rf.requirement_id = " & lngReqId
     
    'Response.Write "strSQL = " & strSQL
    'Response.End 
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetRequirementFiles = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetRequirementFiles = 0

End Function   'fn_DL_GetRequirementFiles



' Detail Requirement File
'***************************************************************************************************
Function fn_DL_DeleteRequirementFile(lngReqId, lngRFId)

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "DELETE FROM RequirementFile" 
    strSQL = strSQL & " " & Chr(13) & "WHERE requirementfile_id = " & lngRFId
    strSQL = strSQL & " " & Chr(13) & "and requirement_id = " & lngReqId
     
    'Response.Write "strSQL = " & strSQL
    'Response.End 
        
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
     
    'If Err.Number <> 0 Then
    '   fn_DL_DeleteRequirementFile = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_DeleteRequirementFile = 0

End Function   'fn_DL_DeleteRequirementFile


' Get Requirement Detail
'***************************************************************************************************
Function fn_DL_GetRequirementFile(lngReqId, lngRFId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "rf.requirementfile_id," 
    strSQL = strSQL & " " & Chr(13) & "rf.oldfilename," 
    strSQL = strSQL & " " & Chr(13) & "rf.newfilename," 
    strSQL = strSQL & " " & Chr(13) & "rf.descrip"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "RequirementFile rf"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "rf.requirementfile_id = " & lngRFId
    strSQL = strSQL & " " & Chr(13) & "and rf.requirement_id = " & lngReqId
     
    'Response.Write "strSQL = " & strSQL
    'Response.End 
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetRequirementFile = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetRequirementFile = 0

End Function   'fn_DL_GetRequirementFile



' Get Project Issue Types
'***************************************************************************************************
Function fn_DL_GetProjectIssueTypes(arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "code,"
    strSQL = strSQL & " " & Chr(13) & "descrip" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "ProjectIssueType"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "ORDER BY descrip"
    
    
    'Response.Write "SQL = " & strSQL
    'Response.End
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetProjectIssueTypes = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetProjectIssueTypes = 0

End Function   'fn_DL_GetProjectIssueTypes


' Get Project Issue Descriptions
'***************************************************************************************************
Function fn_DL_GetProjectIssueDescriptions(lngProjectId, strIssueType, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "projectissue_id,"
    strSQL = strSQL & " " & Chr(13) & "title," 
    strSQL = strSQL & " " & Chr(13) & "resolved_datetime" 
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "ProjectIssue"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "project_id = " & lngProjectId
    strSQL = strSQL & " " & Chr(13) & "and issuetype_code = '" & strIssueType & "'"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "ORDER BY insert_datetime"
    
    
    'Response.Write "SQL = " & strSQL
    'Response.End
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetProjectIssueDescriptions = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetProjectIssueDescriptions = 0

End Function   'fn_DL_GetProjectIssueDescriptions


' Update Issue
'***************************************************************************************************
Function fn_DL_EditIssue(lngIssueId, strTitle, strDetails, dtResolved)

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "UPDATE ProjectIssue" & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "SET title = '" & strTitle & "'," & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "details = '" & strDetails & "'" & vbCrLf
    If dtResolved <> "" Then
       strSQL = strSQL & " " & Chr(13) & ",resolved_datetime = '" & dtResolved & "'" & vbCrLf
    End If
    strSQL = strSQL & " " & Chr(13) & "WHERE projectissue_id = " & lngIssueId
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
      
    'If Err.Number <> 0 Then
    '   fn_DL_EditIssue = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_EditIssue = 0

End Function


' Delete Issue
'***************************************************************************************************
Function fn_DL_DeleteIssue(ByVal lngIssueId)

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "DELETE FROM ProjectIssue" & vbCrLf
    strSQL = strSQL & " " & Chr(13) & "WHERE projectissue_id = " & lngIssueId
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
      
    'If Err.Number <> 0 Then
    '   fn_DL_DeleteIssue = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_DeleteIssue = 0

End Function



' Insert Issue
'***************************************************************************************************
Function fn_DL_AddNewIssue(ByVal lngProjectId, ByVal strIssueType, ByVal strTitle, ByVal strDetails)

    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = "INSERT INTO ProjectIssue"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "project_id,"
    strSQL = strSQL & " " & Chr(13) & "title,"
    strSQL = strSQL & " " & Chr(13) & "details,"
    strSQL = strSQL & " " & Chr(13) & "issuetype_code"
    strSQL = strSQL & " " & Chr(13) & ")"
    strSQL = strSQL & " " & Chr(13) & "VALUES"
    strSQL = strSQL & " " & Chr(13) & "("
    strSQL = strSQL & " " & Chr(13) & "" & lngProjectId & ","
    strSQL = strSQL & " " & Chr(13) & "'" & strTitle & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strDetails & "',"
    strSQL = strSQL & " " & Chr(13) & "'" & strIssueType & "'"
    strSQL = strSQL & " " & Chr(13) & ")"
    
    'Response.Write "strSQL = " & strSQL
    'Response.End
    
    dbconTmp.Execute(strSQL)
    
      
    'If Err.Number <> 0 Then
    '   fn_DL_AddNewIssue = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_AddNewIssue = 0

End Function


' Get Requirement Detail
'***************************************************************************************************
Function fn_DL_GetIssueDetail(lngIssueId, arryTmp)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    Dim strDBPwd
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
    
    strSQL = ""
    strSQL = strSQL & " " & Chr(13) & "SELECT"
    strSQL = strSQL & " " & Chr(13) & "title," 
    strSQL = strSQL & " " & Chr(13) & "details,"
    strSQL = strSQL & " " & Chr(13) & "resolved_datetime"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "FROM"
    strSQL = strSQL & " " & Chr(13) & "ProjectIssue"
    strSQL = strSQL & " " & Chr(13) & ""
    strSQL = strSQL & " " & Chr(13) & "WHERE"
    strSQL = strSQL & " " & Chr(13) & "projectissue_id = " & lngIssueId
        
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
     
      
    'If Err.Number <> 0 Then
    '   fn_DL_GetIssueDetail = Err.Number
    '   Exit Function
    'End If
    
    fn_DL_GetIssueDetail = 0

End Function   'fn_DL_GetIssueDetail

%>





















