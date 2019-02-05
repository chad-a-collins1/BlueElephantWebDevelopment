<%
  Dim strFile
  strFile = Request.ServerVariables("SERVER_NAME")
  'Response.Write "MapPath = " & strFile
  Response.Write "MapPath = " & Replace(Server.MapPath(strFile),strFile,"")
%>