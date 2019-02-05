<%

Function fn_DL_GetProjectInfo(lngProjectId, arryTmp1, arryTmp2)

    Dim dbconTmp
    Dim rsTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 
   
   strSQL = ""  
   strSQL = strSQL & " " & "SELECT Project.project_id, Project.account_id, Project.project_name, Project.projecttype_code, Project.down_payment, Project.status_code, Project.insert_datetime, Project.project_balance, Project.project_estimate_hours, Project.project_estimate_cost, Project.target_date, Project.completion_date, Project.mou_signed, Project.mou_signee, Project.mou_signed_date, Project.mou_body, Project.first_time, Project.accrued_hours, Project.consultant_id, MOU.mou_body, ProjectType.descrip, ProjectStatus.descrip, AccountStatus.descrip, AccountType.descrip, BillingInfo.balance, BillingInfo.balance_forward, BillingInfo.billinginfostatus_code, BillingInfo.business_name, BillingInfo.first_name, BillingInfo.last_name, BillingInfo.phone1, BillingInfo.phone2, BillingInfo.email, BillingInfo.address, BillingInfo.city, BillingInfo.state_code, BillingInfo.postal_code, BillingInfo.state_descrip, BillingInfo.country_descrip, BillingInfoStatus.descrip"
   strSQL = strSQL & " " & "FROM ProjectType INNER JOIN (ProjectStatus INNER JOIN (((BillingInfoStatus RIGHT JOIN ((AccountType INNER JOIN (AccountStatus INNER JOIN Account ON AccountStatus.accountstatus_code = Account.accountstatus_code) ON AccountType.accounttype_code = Account.accounttype_code) LEFT JOIN BillingInfo ON Account.account_id = BillingInfo.account_id) ON BillingInfoStatus.billinginfostatus_code = BillingInfo.billinginfostatus_code) INNER JOIN Project ON Account.account_id = Project.account_id) LEFT JOIN MOU ON Project.project_id = MOU.project_id) ON ProjectStatus.code = Project.status_code) ON ProjectType.projecttype_code = Project.projecttype_code"
   strSQL = strSQL & " " & "WHERE (((Project.project_id)=" & lngProjectId & "))"
   
   'Response.Write "strSQL = " & strSQL & "<br><br>"
   'Response.Write "cs = " & fn_GetConnectionString
   'Response.End

   Set rsTmp = dbconTmp.Execute(strSQL)

    With rsTmp
       If Not .EOF and Not.BOF Then
          arryTmp1 = .GetRows

       End If

    End With  'rsTmp
    
    rsTmp.Close
    Set rsTmp = Nothing   
    dbconTmp.Close
    Set dbconTmp = Nothing
    
    fn_DL_GetProjectInfo = 0      

End Function



%>