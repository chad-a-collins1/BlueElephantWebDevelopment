<!--#include file="./Utility/DBUtil.asp"-->
<%


    Dim dbconTmp
    Dim strSQL
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 

    strSQL = "UPDATE Project SET first_time = 1, status_code = 'NEW' WHERE project_id = 2"    
           
    
    dbconTmp.Execute(strSQL)
    
    dbconTmp.Close
    Set dbconTmp = Nothing
       
    
    
   Response.Write "Project was reset reset!"

%>