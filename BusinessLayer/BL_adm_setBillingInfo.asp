<!-- #include file="../DataLayer/DL_adm_setBillingInfo.asp"-->
<%

Function fn_BL_admBillingInfoInsert(AccountID, Company, ContactName, Email, Phone, Fax, Address, City, State, Zip)


Dim a_splitName, FirstName, LastName, CountryDescrip

a_splitName = Split(ContactName, " ")

If UBound(a_splitName) = 1 Then
   FirstName = a_splitName(0)
   LastName = a_splitName(1)
ElseIf UBound(a_splitName) = 0 or UBound(a_splitName) > 1 Then
   FirstName = a_splitName(0)
   LastName = ""
Else
   FirstName = ""
   LastName = ""
End If

CountryDescrip = "USA"

If fn_DL_admBillingInfoInsert(AccountID, Company, FirstName, LastName, Phone, Fax, Email, Address, City, Zip, State, CountryDescrip) = 1 Then
   fn_BL_admBillingInfoInsert = "billinginfoINSERT success"
Else
   fn_BL_admBillingInfoInsert = "billinginfoINSERT failed"
End If	



End Function
%>