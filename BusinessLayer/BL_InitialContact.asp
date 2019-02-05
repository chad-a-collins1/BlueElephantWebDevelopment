<!--#include file="../DataLayer/DL_InitialContact.asp"-->
<%

' Call the DB function to insert contact information into the DB
'***************************************************************************************************
Function fn_BL_InitialContactInsert(strCompany, strName, strEmail, strPhone, strFax, strAddress, strCity, strState, strZip, strShort)
   
   Dim intRC
   
   strCompany = fn_EscapeSQL(strCompany) 
   strName = fn_EscapeSQL(strName) 
   strEmail = fn_EscapeSQL(strEmail) 
   strPhone = fn_EscapeSQL(strPhone) 
   strFax = fn_EscapeSQL(strFax) 
   strAddress = fn_EscapeSQL(strAddress) 
   strCity = fn_EscapeSQL(strCity) 
   strState = fn_EscapeSQL(strState) 
   strZip = fn_EscapeSQL(strZip) 
   strShort = fn_EscapeSQL(strShort) 

   
   intRC = fn_DL_InitialContactInsert(strCompany, strName, strEmail, strPhone, strFax, strAddress, strCity, strState, strZip, strShort)
       
   If intRC <> 0 Then
      fn_BL_InitialContactInsert = intRC
      Exit Function
   End If
   
   fn_BL_InitialContactInsert = 0

End Function

%>





















