<%
Function fn_DL_admBillingInfoInsert(AccountID, BusinessName, FirstName, LastName, Phone1, Phone2, Email, Address, City, PostalCode, StateDescrip, CountryDescrip)

'On Error Resume Next

	    Dim dbconTmp
	    Dim rs
	    Dim strSQL
	   
	    Set dbconTmp = Server.CreateObject("ADODB.Connection")
	    dbconTmp.ConnectionString = fn_GetConnectionString
	    dbconTmp.Open 

	Set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM BillingInfo"		
	rs.Open strSQL, dbconTmp, 3, 3

rs.AddNew    
rs.Fields("account_id") = AccountID
rs.Fields("balance") = 0.00
rs.Fields("balance_forward") = 0.00
rs.Fields("billinginfostatus_code") = "NORM"
rs.Fields("business_name") = BusinessName 
If  FirstName = "" Then
    'nothing
Else
    rs.Fields("first_name") = FirstName
End If
If  LastName = "" Then
    'nothing
Else
    rs.Fields("last_name") = LastName 
End If
rs.Fields("phone1") = Phone1 
rs.Fields("phone2") = Phone2 
rs.Fields("email") = Email 
rs.Fields("address") = Address 
rs.Fields("city") = City 
rs.Fields("postal_code") = PostalCode 
rs.Fields("state_descrip") = StateDescrip 
rs.Fields("country_descrip") = CountryDescrip
rs.Update



	    rs.Close
	    Set rs = Nothing  

	    dbconTmp.Close
	    Set dbconTmp = Nothing
       
     
	fn_DL_admBillingInfoInsert = 1

End Function
%>