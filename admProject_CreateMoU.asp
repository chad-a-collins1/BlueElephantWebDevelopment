<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!--#include file="Utility/Util.asp"-->
<!--#include file="Utility/DBUtil.asp"-->
<!--#include file="BusinessLayer/BL_adm_getProjectInformation.asp" -->
<%

    Dim lngRC
    Dim lngProjectID, lngAccountID, lngBillingInfoID
    Dim strProjectName
    Dim blnMOUSigned
    Dim strMOUSignee
    Dim dtMOUSigned
    Dim curDownPayment
    Dim strProjectStatus
    Dim dtStart
    Dim dtTargetDate
    Dim dtCompletionDate, dtInsertDateTime
    Dim curProjectBalance
    Dim intEstHours
    Dim curEstCost
    Dim blnFirstTime
    Dim strStatusCode
    Dim blnFreeze
    Dim blnLockEdit        
    Dim strProjectStatusDescrip, strProjectStatusCode
    Dim strProjectTypeDescrip, strProjectTypeCode   
    Dim strBusinessName
    Dim arryTmp1    
    Dim arryTmp2
    Dim tmpInvoiceCOunt
    Dim blobMOUbody
    Dim first_name 
    Dim last_name 
    Dim phone1 
    Dim phone2
    Dim email 
    Dim address
    Dim city 
    Dim state_code
    Dim postal_code 
    Dim state_descrip 
    Dim country_dscrip 
    Dim StatusDescrip  
    Dim dbconTmp
    Dim rsTmp
    Dim rsReq
    Dim strSQL
    Dim strReq
    
    Set dbconTmp = Server.CreateObject("ADODB.Connection")
    dbconTmp.ConnectionString = fn_GetConnectionString
    dbconTmp.Open 


    Set rsReq = Server.CreateObject("ADODB.RecordSet")
    strReq = "SELECT * FROM Requirement WHERE status_code = " & "'" & "APPROVED" & "'" & " ORDER BY requirement_id ASC"
    rsReq.Open strReq, dbconTmp, 3, 3	

    
    lngProjectID = Request.QueryString("pID")

    lngRC = fn_BL_GetProjectInfo(lngProjectID, arryTmp1, arryTmp2)
    If lngRC <> 0 Then
       Call sub_HandleLogicError(lngRC)
    End If
    
   
   lngAccountID = arryTmp1(1,0)
   strProjectName = arryTmp1(2,0)
   strProjectTypeCode = arryTmp1(3,0)
   curDownPayment = arryTmp1(4,0)
   strProjectStatusCode = arryTmp1(5,0)
   dtInsertDateTime = arryTmp1(6,0)
   curProjectBalance = arryTmp1(7,0)
   intEstHours = arryTmp1(8,0)    
   curEstCost = arryTmp1(9,0)       
   dtTargetDate = arryTmp1(10,0)  
   dtCompletionDate = arryTmp1(11,0)
   blnMOUSigned = arryTmp1(12,0)
   strMOUSignee = arryTmp1(13,0) 
   dtMOUSigned = arryTmp1(14,0)   
   blobMOUbody = arryTmp1(15,0)  
   blnFirstTime = arryTmp1(16,0)
'   accrued_hours = arryTmp1(17,0)
'   consultant_id = arryTmp1(18,0)
'   MOU_Body = arryTmp1(19,0)
'   strProjectTypeDescrip = arryTmp1(20,0)   
'   strProjectStatusDescrip = arryTmp1(21,0) 
'   AccountStatus.descrip = arryTmp1(22,0)
'   AccountType.descrip = arryTmp1(23,0)
'   balance = arryTmp1(24,0)   
'   balance_forward = arryTmp1(25,0)
'   billinginfostatus_code = arryTmp1(26,0)
   strBusinessName = arryTmp1(27,0)
   first_name = arryTmp1(28,0)
   last_name = arryTmp1(29,0)
   phone1 = arryTmp1(30,0)
   phone2 = arryTmp1(31,0)
   email = arryTmp1(32,0)
   address = arryTmp1(33,0)
   city = arryTmp1(34,0)
   state_code = arryTmp1(35,0)
   postal_code = arryTmp1(36,0)
   state_descrip = arryTmp1(37,0)
   country_dscrip = arryTmp1(38,0)
   StatusDescrip = arryTmp1(39,0) 
 

%>


<html>
<head>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META name=VI60_DTCScriptingPlatform content="Server (ASP)">
<META name=VI60_defaultClientScript content=JavaScript>

<title>Memorandum of Understanding</title>  	


</head>
<body>
<%


Response.Write "<br><br>"
Response.Write "<center><b><h3>Memorandum of Understanding</h3></b></center>"
Response.Write "<br><br>"


	Response.Write "This document, known as the Memorandum of Understanding (" & """MOU""" & "), details the relationship between Bay Area Consulting, a Texas Limited Liability Partnership (" & """BAC""" & ") and"

	Response.Write " " & strBusinessName & " "

	Response.Write "referred to in this document as the " & """ THE CLIENT""" & "."
	Response.Write "<br><br>"	
	Response.Write "WHEREAS THE CLIENT wishes to obtain services from the Service Provider as specified by Schedule "'"A"'" of this document;"
	Response.Write "<br><br>"
	Response.Write "WHEREAS BAC has agreed to provide THE CLIENT with the services described herinbelow, in return for good and valuable consideration;"
	Response.Write "<br><br>"
	Response.Write "WHEREAS the Parties wish to evidence their agreement in writing;"
	Response.Write "<br><br>"
	Response.Write "WHEREAS the Parties are duly authorized and have the capacity to enter into and perform this Agreement;"
	Response.Write "<br><br>"
	Response.Write "NOW THEREFORE, THE PARTIES AGREE AS FOLLOWS:"
	Response.Write "<br><br>"
	Response.Write "1.	Consulting Services."
	Response.Write "<br><br>"
	Response.Write "a)	BAC" & "'" & "s Duties.  In response to specific requests by the CLIENT, BAC will provide to the CLIENT the consulting services described in a mutually agreeable Work Order."
	Response.Write "<br><br>"
	Response.Write "b)	Hours of Work.  BAC will devote the amount of time necessary to complete the services requested pursuant to this Agreement."  
	Response.Write "<br><br>"
	Response.Write "c)	Other Engagements.  BAC may accept other consulting assignments, and engage in other business activities, so long as they do not interfere with his obligations under this Agreement."
	Response.Write "<br><br>"
	Response.Write "2.	Compensation."
	Response.Write "<br><br>"
	Response.Write "a)	Consulting Fee.  The CLIENT will compensate BAC at a rate of _______ per hour.  BAC will be paid within 30 days of receiving an invoice for the services performed by BAC."
	Response.Write "<br><br>"
	Response.Write "3.	Confidential Information."
	Response.Write "<br><br>"
	Response.Write "a)	General Restrictions.  The CLIENT and BAC agree to maintain the confidentiality of each other’s trade secrets and confidential business information disclosed during the term of this Agreement, except as authorized by the party that disclosed the information.  Upon completion of the consulting services, the parties will return all confidential materials and equipment provided during the term of this Agreement, except as authorized by the party that provided the materials or equipment.  Each party is responsible for identifying all trade secrets, confidential business information and confidential materials."
	Response.Write "<br><br>"
	Response.Write "b)	Limitation.  Nothing in this Agreement or any other agreement (i) shall restrict any party’s use or disclosure of information that is or becomes publicly known through lawful means, that was rightfully in that party" & "'" & "s possession or part of his general knowledge prior to the term of this Agreement, or that is disclosed to that party without confidential or proprietary restrictions by a person who rightfully possesses the information; or (ii) shall prevent any party from responding to a lawful subpoena or court order."
	Response.Write "<br><br>"
	Response.Write "4.	Independent Contractor Status.	BAC is an independent contractor and is responsible for the payment of all taxes applicable to the consulting services."
	Response.Write "<br><br>"
	Response.Write "5.	Indemnification.	The CLIENT agrees to indemnify, defend and hold harmless BAC against any liability incurred by BAC within the course and scope of the consulting services provided by BAC."
	Response.Write "<br><br>"
	Response.Write "6.	Arbitration.  All claims that BAC and the CLIENT may have against each other in any way related to the subject matter, interpretation, application, or alleged breach of this Agreement (""Arbitrable Claims"") shall be resolved by arbitration in Texas in accordance with the rules of the American Arbitration Association, as amended.  Arbitration shall be final and binding upon the parties and shall be the exclusive remedy for all Arbitrable Claims."
	Response.Write "<br><br>"
	Response.Write "7.	Miscellaneous Provisions."
	Response.Write "<br><br>"
	Response.Write "a)	Integration.  The parties agree that all agreements and understandings between the parties concerning the subject matter of this Agreement are embodied in this Agreement and any Work Order to which the parties agreed.  This Agreement shall supersede all prior or contemporaneous agreements and understandings between the parties, with respect to any subject covered by this Agreement, except as otherwise provided in this Agreement."
	Response.Write "<br><br>"
	Response.Write "b)	Amendments; Waivers.  This Agreement may not be amended except by an instrument in writing, signed by each of the parties.  No failure to exercise and no delay in exercising any right under this Agreement shall operate as a waiver thereof."
	Response.Write "<br><br>"
	Response.Write "c)	Assignment; Successors and Assigns.  Neither party shall assign or otherwise transfer any rights or obligations under this Agreement, without the written consent of the other party.  Subject to the foregoing, this Agreement shall be binding upon and shall inure to the benefit of the parties’ respective heirs, successors, attorneys, and permitted assigns."
	Response.Write "<br><br>"
	Response.Write "d)	Severability.  If any provision of this Agreement, or its application to any person, place, or circumstance, is held by an arbitrator or a court of competent jurisdiction to be invalid, unenforceable, or void, such provision shall be enforced to the greatest extent permitted by law, and the remainder of this Agreement and such provision as applied to other persons, places, and circumstances shall remain in full force and effect."
	Response.Write "<br><br>"
	Response.Write "e)	Governing Law.  This Agreement shall be governed by and construed in accordance with the law of the State of Texas."
	Response.Write "<br><br>"
	Response.Write "f)	Interpretation.  This Agreement shall be construed as a whole, according to its fair meaning, and not in favor of or against any party.  Captions are used for reference purposes only and should be ignored in the interpretation of the Agreement."
	Response.Write "<br><br>"
	Response.Write "The parties have duly executed this Agreement as of " & Date()
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "APPENDIX " & """A"""
	Response.Write "<br><br>"
	Response.Write ""
	
	Do While Not rsReq.EOF

	   Response.Write rsReq.Fields("summary") & "<br><br>"
		
	rsReq.MoveNext
	Loop

	Response.Write "<br><br>"
	Response.Write ""


rsReq.close
Set rsReq = Nothing

dbconTmp.close
Set dbconTmp = Nothing









%>
</body>
</html>
































