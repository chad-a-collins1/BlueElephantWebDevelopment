<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="upload.asp" -->
<%
Const PAGE_EXAMPLE = "uploadexample.asp"
Const PATH_UPLOAD = "uploads\"
'Const PATH_UPLOAD = "\dev\ConsignmentShop\"

'NOTE - YOU MUST HAVE VBSCRIPT v5.0 INSTALLED ON YOUR WEB SERVER
'	   FOR THIS LIBRARY TO FUNCTION CORRECTLY. YOU CAN OBTAIN IT
'	   FREE FROM MICROSOFT WHEN YOU INSTALL INTERNET EXPLORER 5.0
'	   OR LATER.


' Create the FileUploader
Dim Uploader, File
Set Uploader = New FileUploader

On Error Resume Next

' This starts the upload process
Uploader.Upload()

If Err.Number <> 0 Then
   Response.Write "Error  - " & Err.Description
   Response.End
End If

'******************************************
' Use [FileUploader object].Form to access 
' additional form variables submitted with
' the file upload(s). (used below)
'******************************************


' Check if any files were uploaded
If Uploader.Files.Count = 0 Then
	Response.Write "File(s) not uploaded."
Else
	' Loop through the uploaded files
	For Each File In Uploader.Files.Items
		
		' Check where the user wants to save the file
		If Uploader.Form("saveto") = "disk" Then
	
	
			' Save the file
	                Response.Write "upload path = " & Replace(Server.MapPath(PAGE_EXAMPLE),PAGE_EXAMPLE,"") & PATH_UPLOAD & "<br>"
			File.SaveToDisk Replace(Server.MapPath(PAGE_EXAMPLE),PAGE_EXAMPLE,"") & PATH_UPLOAD
	
	                If Err.Number <> 0 Then
                           Response.Write "Error  - " & Err.Description
                           Response.End
                        End If
	
		ElseIf Uploader.Form("saveto") = "database" Then
			
			' Open the table you are saving the file to
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.Open "MyUploadTable", "CONNECT STRING OR ADO.Connection", 2, 2
			RS.AddNew ' create a new record
			
			RS("filename")    = File.FileName
			RS("filesize")	  = File.FileSize
			RS("contenttype") = File.ContentType
		
			' Save the file to the database
			File.SaveToDatabase RS("filedata")
			
			' Commit the changes and close
			RS.Update
			RS.Close
		End If
		
		' Output the file details to the browser
		Response.Write "File Uploaded: " & File.FileName & "<br>"
		Response.Write "Size: " & File.FileSize & " bytes<br>"
		Response.Write "Type: " & File.ContentType & "<br><br>"
	Next
	
	
	Response.Write "<br><b>Thank you for your upload " & Uploader.Form("fullname") & "</b><br>"
	
End If



If Err.Number <> 0 Then
   Response.Write "Error  - " & Err.Description
   Response.End
End If

%>