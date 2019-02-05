<!-- #include file="../DataLayer/DL_adm_setAccount.asp"-->
<%

'
'************************************************************************
Function fn_BL_adm_InitContactToAccount(lngCntctID, intConsultantId, arryTmp)

On Error Resume Next

   Dim strAcctNum , strUID, strFakePwd, strRealPwd, Company, lngAcctId
   Dim ContactName, Email, Phone, Fax, Address,	City, State, Zip, ProjectTypeCode, StatusCode

   If  fn_DL_adm_InitContactToAccount(lngCntctID, arryTmp) <> 0 Then
	
	Response.Write "Datalayer call failed"
		
   Else		

	If IsArray(arryTmp) Then

		Company = arryTmp(1,0)
   		ContactName = arryTmp(2,0)
   		Email = arryTmp(9,0)
   		Phone = arryTmp(7,0)
   		Fax = arryTmp(8,0)
   		Address = arryTmp(3,0)
   		City = arryTmp(4,0)
   		State = arryTmp(5,0)
   		Zip = arryTmp(6,0)   

		Company = Replace(Company,"'","''")
   		ContactName = Replace(ContactName,"'","''")
   		Email = Replace(Email,"'","''")
   		Phone = Replace(Phone,"'","''")
   		Fax = Replace(Fax,"'","''")
   		Address = Replace(Address,"'","''")
   		City = Replace(City,"'","''")
   		State = Replace(State,"'","''")
   		Zip = Replace(Zip,"'","''")

		ProjectTypeCode = "OTHER"
   		StatusCode = "NEW"

	Else

   		Response.Write "Bad Array"

	End If

   End If

   strAcctNum = fn_BL_admCreateAccountNumber
   strUID = Email
   strFakePwd = fn_GetRandomAlphaNumeric(7, 10)
   strRealPwd = fn_IncryptString(strFakePwd)

   Call fn_DL_initialAccountInsert(strAcctNum, strUID, strFakePwd, strRealPwd, lngAcctId)
   'lngAcctId = 16
   
 If Err.Number <> 0 Then
  Response.Write "after DL_initialAccountInsert"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if     

   Call fn_BL_admContactInsert(lngAcctId, Company, ContactName, Email, Phone, Fax, Address, City, State, Zip)
   
 If Err.Number <> 0 Then
  Response.Write "after BL_admContactInsert"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if     
   
   Call fn_BL_admBillingInfoInsert(lngAcctId, Company, ContactName, Email, Phone, Fax, Address, City, State, Zip)

 If Err.Number <> 0 Then
  Response.Write "after BL_admBillingInfoInsert"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if  
   
   Call fn_BL_admProjectInsert(lngAcctId, ProjectTypeCode, StatusCode, intConsultantId)
   
 If Err.Number <> 0 Then
  Response.Write "after BL_admProjectInsert"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if  
   
   Call fn_DL_admUpdateInitContact(lngCntctID, intConsultantId)


 If Err.Number <> 0 Then
  Response.Write "after fn_DL_admUpdateInitContact"
  Response.Write "<br>Err=" & Err.Description
  Response.End
End if  

End Function





'*************************************************************************
Function fn_BL_admCreateUserLogin(strContactName, strUID, strPasswd)
Dim strTmpRnd1
Dim strTmpRnd2

	strTmpRnd1 = fn_GetRandomAlphaNumeric(2, 4)
	strTmpRnd2 = fn_GetRandomAlphaNumeric(7, 10)

	strUID = strContactName & strTmpRnd1
	strPasswd = strTmpRnd2

End Function





Function fn_BL_admCreateAccountNumber()
Dim strTmpRnd1
Dim strTmpRnd2

	strTmpRnd1 = fn_GetRandomAlphaNumeric(12,16)
	fn_BL_admCreateAccountNumber = strTmpRnd1

End Function























%>