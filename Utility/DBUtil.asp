<%

' This function returns a href based on the value of GLOB_PRODUCTION
'**********************************************************
Function fn_GetConnectionString()

Dim strDBPath
'Const GLOB_CS = "Provider=SQLOLEDB; Data Source=GORILLA1; Initial Catalog=PSdb; User Id=webapp; Password=admin"
'Const GLOB_CS = "Provider=SQLOLEDB; Data Source=ntsql05.propagtion.net; Initial Catalog=PSdb; User Id=passwordsupportcom; Password=passwordsupport330"
'Const GLOB_CS = "Provider=SQLOLEDB; Data Source=66.34.127.254; Initial Catalog=PSdb; User Id=passwordsupportcom; Password=passwordsupport330"
'strDBpath = Server.MapPath("./Utility/Collins.mdb")
strDBpath = Server.MapPath("db/Collins.mdb")

'Response.Write "strDBPath = " & strDBPath
'Response.End

   If GLOB_PRODUCTION = "yes" Then
     'fn_GetConnectionString = "Provider=SQLOLEDB; Data Source=66.34.127.254; Initial Catalog=PSdb; User Id=passwordsupportcom; Password=passwordsupport330" 
     fn_GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"
   Else
      'fn_GetConnectionString = "Provider=SQLOLEDB; Data Source=GORILLA1; Initial Catalog=PSdb; User Id=webapp; Password=admin"
     fn_GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"
   End If

End Function


%>





















