<% @LANGUAGE = "VBScript"%>
<% Response.Buffer = True %>

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



strDBpath = Server.MapPath("Collins.mdb")


Set conn = Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBpath & ";"

strRS3 = "SELECT * FROM Account WHERE username = " & "'" & "crooksoft" & "'" 'Request.QueryString("uid")
Set rs3 = Server.CreateObject("ADODB.Recordset")
rs3.Open strRS3, conn, 3, 3


strRS1 = "SELECT * FROM BillingInfo WHERE account_id = " & rs3.Fields("account_id")
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open strRS1, conn, 3, 3



Response.Write "<br><br>"
Response.Write "<center><b><h3>Memorandum of Understanding</h3></b></center>"
Response.Write "<br><br>"


	Response.Write "This document, known as the Memorandum of Understanding (''MOU''), details the relationship between Bay Area Consulting (''BAC''), a Texas Limited Liability Partnership and"
	Response.Write " " & rs1.Fields("business_name") 
	Response.Write " (" & """THE CLIENT""" & ")."
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
	Response.Write "a)	BAC’s Duties.  In response to specific requests by the CLIENT, BAC will provide to the CLIENT the consulting services described in a mutually agreeable Work Order."
	Response.Write "<br><br>"
	Response.Write "b)	Hours of Work.  BAC will devote the amount of time necessary to complete the services requested pursuant to this Agreement.  BAC’s services will be performed during regular business hours, unless mutually agreed otherwise."
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
	Response.Write "b)	Limitation.  Nothing in this Agreement or any other agreement (i) shall restrict any party’s use or disclosure of information that is or becomes publicly known through lawful means, that was rightfully in that party’s possession or part of his general knowledge prior to the term of this Agreement, or that is disclosed to that party without confidential or proprietary restrictions by a person who rightfully possesses the information; or (ii) shall prevent any party from responding to a lawful subpoena or court order."
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
	Response.Write "The parties have duly executed this Agreement as of"
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""
	Response.Write "<br><br>"
	Response.Write ""























	rs1.Close
	Set rs1 = Nothing

'	rs2.Close
'	Set rs2 = Nothing

	rs3.Close
	Set rs3 = Nothing
	
	conn.Close
	Set conn = Nothing

%>
</body>
</html>
































