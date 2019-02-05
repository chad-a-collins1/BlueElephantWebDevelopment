<!--#include file="../DataLayer/DL_admContact.asp"-->
<%

' Call the DB function to insert contact information into the DB
'***************************************************************************************************
Function fn_BL_admContactInsert(strAccountID, strCompany, strName, strEmail, strPhone, strFax, strAddress, strCity, strState, strZip)
'On Error Resume Next
   
   Dim intRC
   Dim a_splitName, strFirst, strLast
   

   strAccountID = CLng(strAccountID)


   a_splitName = Split(strName, " ")

   If UBound(a_splitName) > 0 Then
	strFirst = a_splitName(0)
	strLast = a_splitName(1)
   ElseIf UBound(a_splitName) = 0 Then
	strFirst = a_splitName(0)
	strLast = ""
   Else
	strFirst = ""
	strLast = ""
   End If
   
   
If fn_DL_ContactInsert(strAccountID, strFirst, strLast, strPhone, strFax, strEmail, strAddress, strCity, strZip, strState) <> 0 Then
	fn_BL_admContactInsert = "success"
Else
	fn_BL_admContactInsert = "failed"
End If

  
 'Call fn_DL_admBillingInfoInsert(strCompany, ) <------ Finish adding contact info for BILLINGINFO Table







       
'   If intRC <> 0 Then
'      fn_BL_admContactInsert = intRC
'      Exit Function
'   End If
   
'   fn_BL_admContactInsert = 0

End Function

%>





















