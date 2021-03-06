<%
@LANGUAGE="VBScript"
ENABLESESSIONSTATE="True"
%>
<%
    Option Explicit
%>
<!-- #include file="Utility/adovbs.inc" -->
<!-- #include file="Utility/Util.asp" -->
<!-- #include file="Utility/DBUtil.asp" -->
<!-- #include file="Utility/upload.asp" -->
<!-- #include file="BusinessLayer/BL_Portal.asp" -->
<!-- #include file="Utility/Random.asp" -->
<!-- #include file="DataLayer/DL_Portal.asp" -->
<%
Const ACTION_BEGINPROJECT = "aq_924lsng"
Const ACTION_QUESTION = "aq_489gfeh"
Const ACTION_PROJECTSTATUS = "aps_5gbfu"
Const ACTION_VALIDATECHOOSE = "avc_giu2t6"
Const ACTION_VALIDATEQUESTION = "avqu_27fvb"
Const ACTION_REQUIREMENT = "areq_13t7v"
Const ACTION_VALIDATE_EDIT_REQUIREMENT = "aver_1538gb78"
Const ACTION_VALIDATE_ADDNEW_REQUIREMENT = "avan_1538gb78"
Const ACTION_EDIT_REQUIREMENT = "aer_fhse25"
Const ACTION_ADDNEW_REQUIREMENT = "aan_fhse25"
'Const ACTION_PROJECTFIRSTTIME = "apf_478dhfv"
Const ACTION_PROJECTMESSAGE = "apm_138ejhb"
Const OPTION_BEFOREQUESTION = "obq_37fhv"
Const OPTION_AFTERANSWER = "oaa_37fhv"
Const OPTION_FREEZE = "of_sdvjbt"
Const ACTION_REQUIREMENT_ATTACH_FILE = "araft79bvygf"
Const ACTION_VALIDATE_REQUIREMENT_ATTACH_FILE = "avraft79bvygf"
Const ACTION_DELETE_REQUIREMENT_FILE = "adrf368rv"
Const ACTION_DOWNLOAD_REQUIREMENT_FILE = "adf_t378fve"
Const STATUS_PEND = "PEND"
Const ACTION_DISCREPREPORT = "adr2_fh537b"
Const ACTION_CHANGEREQUEST = "acr_fh537b"
Const ACTION_ISSUES_MAIN = "aim_fh537b"
Const ACTION_ISSUE = "ais_akejg"
Const ACTION_EDIT_ISSUE = "aei_487brh"
Const ACTION_ADDNEW_ISSUE = "ani_fe6hgr33"
Const ACTION_VALIDATE_ADDNEW_ISSUE = "av4ni_4gr"
Const ACTION_VALIDATE_EDIT_ISSUE = "avei_537fhbe"

'Sub Main
'*********************************************
Sub Main()

'On Error Resume Next
  
   Dim strQry
   Dim strAct
   Dim strLP
   Dim lngRC
   Dim lngAcctId
   Dim arryTmp
   Dim strMsg
   Dim blnWelcome
   Dim lngReqId
   Dim lngIssueId
   
   lngRC = fn_CheckForLogin
   If lngRC <> 0 Then
      Response.Redirect fn_CreateURL(True,PAGE_LOGIN,arryVar)
   End If
   
   strAct = Request(QSVAR & "1")

   Select Case strAct
            
      Case ACTION_MAIN: 

          ' Get project descriptions and if there is only one project then redirect to project status
          '*******************************************************************************************
          strMsg = ""
          lngAcctId = Session("sess_lngAcctId")
          lngRC = fn_BL_GetProjectDescriptions(lngAcctId, arryTmp)
          If lngRC <> 0 Then
             Call sub_HandleLogicError(lngRC)
          End If
          
          If IsArray(arryTmp) Then
               If Ubound(arryTmp, 2) = 0 Then
                  Dim arryVar(2)
                  arryVar(0) = ACTION_BEGINPROJECT   ' set action
                  arryVar(1) = arryTmp(0,0)           'set project id
                  Session("sess_lngProjectId") = arryTmp(0,0) 
                  arryVar(2) = 1                   'set Welcome indicator to True
                  Erase arryTmp
                  Response.Clear
                  Response.Redirect fn_CreateURL(True,PAGE_PORTAL,arryVar)
                  Exit Sub
               End If
                
          Else
              strMsg = "There are currently no projects!"
          End If
          
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayPortalMain (strMsg, arryTmp)
          Call DisplayFooter
                        

      Case ACTION_VALIDATECHOOSE:
          Call sub_ValidateChoose                         
                        
      Case ACTION_PROJECTSTATUS: 
          blnWelcome = Request(QSVAR & "2")
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayProjectStatus (blnWelcome)      
          Call DisplayFooter
                           
      Case ACTION_PROJECTMESSAGE:
          Dim strOption
          strOption = Request(QSVAR & "2")
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayProjectMessage (strOption)      
          Call DisplayFooter
          
      Case ACTION_QUESTION:
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayProjectQuestions 
          Call DisplayFooter
      
      Case ACTION_VALIDATEQUESTION:  
          Call sub_ValidateQuestion  
          
      Case ACTION_REQUIREMENT:
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayProjectRequirements ()      
          Call DisplayFooter  
          
      Case ACTION_EDIT_REQUIREMENT:
          lngReqId = CLng(Request(QSVAR & "2"))
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayRequirementEdit (lngReqId)      
          Call DisplayFooter      
          
      Case ACTION_VALIDATE_EDIT_REQUIREMENT:  
          Call sub_ValidateRequirementEdit  
          
      Case ACTION_ADDNEW_REQUIREMENT:
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayRequirementAddNew    
          Call DisplayFooter      
          
      Case ACTION_VALIDATE_ADDNEW_REQUIREMENT:  
          Call sub_ValidateRequirementAddNew    
          
      Case ACTION_REQUIREMENT_ATTACH_FILE:
          lngReqId = CLng(Request(QSVAR & "2"))
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayRequirementAttachFile (lngReqId)   
          Call DisplayFooter       
          
      Case ACTION_VALIDATE_REQUIREMENT_ATTACH_FILE:    
          Call sub_ValidateRequirementAttachFile       
      
      Case ACTION_DELETE_REQUIREMENT_FILE:
          Dim lngRFId, strFileName  
          lngRFId = Request(QSVAR & "2")
          strFileName = Request(QSVAR & "3")
          Call sub_ValidateDeleteRequirementFile(lngRFId, strFileName)       
          
      Case ACTION_DOWNLOAD_REQUIREMENT_FILE:    
         Dim lngReqFileId, strFN, strDFN
         lngReqFileId = Request(QSVAR & "2")  
         strFN = Request(QSVAR & "3")
         strDFN = Request(QSVAR & "4")
         Call sub_ValidateDownloadRequirementFile(lngReqFileId, strFN, strDFN)  
      
      Case ACTION_ISSUES_MAIN:
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayProjectIssues ()   
          Call DisplayFooter      

      Case ACTION_ISSUE:
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayIssue ()   
          Call DisplayFooter  
      
      Case ACTION_EDIT_ISSUE:
          lngIssueId = CLng(Request(QSVAR & "2"))
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayIssueEdit (lngIssueId)      
          Call DisplayFooter      
          
      Case ACTION_VALIDATE_EDIT_ISSUE:  
          Call sub_ValidateIssueEdit  
          
      Case ACTION_ADDNEW_ISSUE:
          Call DisplayHeader(TITLE_PORTAL, PIC_PORTAL, LOGO_GRAPHIC)      
          Call DisplayIssueAddNew    
          Call DisplayFooter      
          
      Case ACTION_VALIDATE_ADDNEW_ISSUE:  
          Call sub_ValidateIssueAddNew                    
         
                        
      Case Else:
          Call sub_HandleLogicError(ERR_INVALID_ACTION)
           
   End Select
   
   
   
  'If Err.Number <> 0 Then
  '   Dim objASPError
  '   Dim strFile
  '   Dim intLine
  '   
  '   Set objASPError = Server.GetLastError()
  '   strFile = objASPError.File
  '   intLine = objASPError.Line
  '
  '   Response.Write "Error " & Err.Number & ", " & Err.Description & ", " & strFile & ", " & intLine
  '   Response.End 
  'End If 
  
  'sub_ErrorCatch
   

End Sub   'Main
 
 

'Sub DisplayPortalMain
'*********************************************
Sub DisplayPortalMain(strMsg, arryTmp)
Dim strFN
Dim strLN
Dim arryVar(1)
   
strFN = Session("sess_strFirstName")
strLN = Session("sess_strLastName")
arryVar(0) = ACTION_VALIDATECHOOSE
arryVar(1) = fn_GetRandomAlphaNumeric(11,16)
%>
<table width="600">
  <tr>
    <td align="left">
   <form action="<% = fn_CreateURL(True,PAGE_PORTAL,arryVar) %>" method="post" name="theForm">

   <table> 
     <tr>
       <%
        Dim arryVarLO(0)
        arryVarLO(0) = ACTION_LOGOUT
        Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">" & fn_InsertSpaces(20) & "Welcome " & strFN & " " & strLN & "</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVarLO) & """>LogOut</a>]</td>"
      %>
     </tr>   

    <tr>
       <td colspan=2>
         <br><br>
         <table width=200>
          <tr>
            <td align=center>
          <%
            If strMsg <> "" Then
             Response.Write "<font>" & strMsg & "</font><br>" & vbCrLf
            End if
            
            If IsArray(arryTmp) Then
               Dim i
               
               Response.Write "<select name=""lstProjects"">" & vbCrLf
               Response.Write "<option value=""none"" selected> &lt;Choose A Project&gt;" & vbCrLf
               For i = 0 to Ubound(arryTmp,2)
                  'Print arryTmp(0,i) & ", " & arryTmp(1,i) 
                  Response.Write "<option value=""" & arryTmp(0,i) & """>" & arryTmp(1,i) & vbCrLf
               Next    
               Response.Write "</select>" & vbCrLf
               
               Erase arryTmp    
               
            End If   
          %>

            <br><br>
            <input type="submit" value="Go"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      
           </td>
         </tr>
       </table>
      
      </td>
    </tr>  
   </table>

   </form>
  </td>
 </tr>
</table>   
<% 
End Sub           'DisplayPortalMain


'Sub Validate Choose Project
'*********************************************
Sub sub_ValidateChoose()
    
   Dim lngProjectId    
   Dim arryVar(2)
        
   lngProjectId = Request("lstProjects")
   
   If IsNumeric(lngProjectId) Then
        Session("sess_lngProjectId") = lngProjectId 
                
        arryVar(0) = ACTION_PROJECTSTATUS   'set action
        arryVar(1) = 0                  'set Welcome indicator to True
        Response.Redirect fn_CreateURL(True,PAGE_PORTAL,arryVar)   
   Else

       arryVar(0) = ACTION_MAIN
       Response.Redirect fn_CreateURL(True, PAGE_PORTAL, arryVar)
   
   End If                  
       
End Sub    ' sub_ValidateChoose


'Sub DisplayProjectStatus
'*********************************************
Sub DisplayProjectStatus(blnWelcome)
%>
   <table> 
     <tr>
<%     
    Dim arryVar(1)
    Dim lngRC
    Dim lngProjectId 
    Dim strProjectName
    Dim blnMOUSigned
    Dim strMOUSignee
    Dim dtMOUSigned
    Dim curDownPayment
    Dim strProjectStatus
    Dim dtStart
    Dim dtTargetDate
    Dim dtCompletionDate
    Dim curProjectBalance
    Dim intEstHours
    Dim curEstCost
    Dim blnFirstTime
    Dim strStatusCode
    Dim blnFreeze
    Dim blnLockEdit        
    Dim arryTmp1    
    Dim arryTmp2
    
    lngProjectId = Session("sess_lngProjectId")
    
    lngRC = fn_BL_GetProjectInfo(lngProjectId, arryTmp1, arryTmp2)
    If lngRC <> 0 Then
       Call sub_HandleLogicError(lngRC)
    End If
    
   strProjectName = arryTmp1(0,0) 
   blnMOUSigned = arryTmp1(1,0)
   strMOUSignee = arryTmp1(2,0) 
   dtMOUSigned = arryTmp1(3,0)
   curDownPayment = arryTmp1(4,0) 
   strProjectStatus = arryTmp1(5,0) 
   dtStart = arryTmp1(6,0)
   dtCompletionDate = arryTmp1(7,0) 
   dtTargetDate = arryTmp1(8,0) 
   curProjectBalance = arryTmp1(9,0) 
   intEstHours = arryTmp1(10,0) 
   curEstCost = arryTmp1(11,0)
   blnFirstTime = arryTmp1(12,0)
   strStatusCode = arryTmp1(13,0)
   blnFreeze = arryTmp1(14,0)
   blnLockEdit = arryTmp1(15,0)
   
   Session("sess_arryProjectInfo") = arryTmp1
   'Session("sess_blnFirstTime") = blnFirstTime
    
    If blnFirstTime = True Then
       arryVar(0) = ACTION_PROJECTMESSAGE
       arryVar(1) = OPTION_BEFOREQUESTION
       Response.Clear
       Response.Redirect fn_CreateURL(True,PAGE_PORTAL,arryVar)
       Exit Sub
    End If
   
   If blnFreeze = True Then
       arryVar(0) = ACTION_PROJECTMESSAGE
       arryVar(1) = OPTION_FREEZE
       Response.Clear
       Response.Redirect fn_CreateURL(True,PAGE_PORTAL,arryVar)
       Exit Sub
   End If 
    
   arryVar(0) = ACTION_LOGOUT
   arryVar(1) = fn_GetRandomAlphaNumeric(11,16)
    If blnWelcome Then
      Dim strFN
      Dim strLN   
      strFN = Session("sess_strFirstName")
      strLN = Session("sess_strLastName")
      Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">" & fn_InsertSpaces(20) & "Welcome " & strFN & " " & strLN & "</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar) & """>LogOut</a>]</td>"
    Else 
        Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar) & """>LogOut</a>]</td>"
    End If 

    Dim intTentative
    Dim intFinalized    
    Dim intFulfilled
    Dim intTotal
    Dim dblPercent
    
    intTentative = CInt(arryTmp2(2,0)) 
    intFinalized = CInt(arryTmp2(2,1)) 
    intFulFilled = CInt(arryTmp2(2,2)) 
    intTotal = CInt(arryTmp2(2,0)) + CInt(arryTmp2(2,1))
    
    If intFulfilled > 0 and intTotal > 0 Then
      dblPercent = intFulfilled/CDbl(intTotal)
    Else
      dblPercent = 0
    End if 
%>
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="top">
    <span class="smallsubtitle">Project Main</span>
     </td>
     </tr>   
   </table>
   <br>
   <table>
     <tr>
       <td align="right"><font><b>Project Name:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = strProjectName %></td>
     </tr>
     
   <% 
      Dim strStatus
      
         If arryTmp2(2,0) = 0 Then 
            If blnMOUSigned = "Yes" Then
               strStatus = "MoU has been signed and work has begun, " & intFinalized & " of " & intTotal & " requirements have been fulfilled, Percent Complete = " & FormatPercent(dblPercent,0) & "."
            Else
               strStatus = "All requirements have been finalized but Estimation or MoU is still pending."
            End if
          Else 
            strStatus = "No work has been done, there are " & intTentative & " tentative requirement(s) that must be resolved before an estimate can be delivered and a contract(MoU) can be written."
          End If
         'strStatus = ""  
   %>
     <tr>
       <td align="right"><font><b>Project Status:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = strProjectStatus %></td>
     </tr>
     <tr>
       <td align="right"><font><b></b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left width=300><% = strStatus %></td>
     </tr>
     <tr>
       <td align="right"><font><b>Project Start:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = dtStart %></td>
     </tr>

<%
        Response.Write "<tr>"
        Response.Write "<td align=""right""><font><b>" & arryTmp2(1,0) & ":</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=""left"">" & arryTmp2(2,0) & "</td>"
        Response.Write "</tr>"
        Response.Write "<tr>"
        Response.Write "<td align=""right""><font><b>" & arryTmp2(1,1) & ":</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=""left"">" & arryTmp2(2,1) & "</td>"
        Response.Write "</tr>"
        Response.Write "<tr>"
        Response.Write "<td align=""right""><font><b>" & arryTmp2(1,2) & ":</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=""left"">" & arryTmp2(2,2) & "</td>"
        Response.Write "</tr>"

%>     

     <tr>
       <td align="right"><font><b>Estimated Work (hrs):</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = intEstHours %></td>
     </tr>
     <tr>
       <td align="right"><font><b>Est Completion Date:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = dtTargetDate %></td>
     </tr>
     <tr>
       <td align="right"><font><b>Estimated Cost (USD):</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% If Not IsNull(curEstCost) Then Response.Write "$" & curEstCost End if %></td>
     </tr>
    
     <tr>
       <td align="right"><font><b>MOU Signed:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = blnMOUSigned %></td>
     </tr>
     <tr>
       <td align="right"><font><b>MOU Signed By:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = strMOUSignee %></td>
     </tr>
     <tr>
       <td align="right"><font><b>MOU Date Signed:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = dtMOUSigned %></td>
     </tr>     
         
     <tr>
       <td align="right"><font><b>Project Balance (USD):</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left>$<% = curProjectBalance %></td>
     </tr>
     <tr>
       <td align="right"><font><b>Down Payment Paid (USD):</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left>$<% = curDownPayment %></td>
     </tr>

     <tr>
       <td align="right"><font><b>Project Completion Date:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=left><% = dtCompletionDate %></td>
     </tr>
                   


   </table>
   

  <p>
  *Note: If all requirements have been APPROVED by the client, then the MoU can be generated.
   If the MoU has been signed by the client and delivered to BayAreaConsulting.biz via mail or fax,
   and BayAreaConsulting.biz has recieved 25% down payment, The project status can be updated as a percentage complete
   based on the requirements that we have fullfilled out of the total requirements that were approved.
  </p>

<% 
   Erase arryTmp1
   Erase arryTmp2

End Sub           'DisplayProjectStatus



'Sub DisplayProjectMessage
'*********************************************
Sub DisplayProjectMessage(strOption)

%>
<!-- <table width="600">
  <tr>
    <td align="left"> -->

   <table> 
     <tr>
  <%   
    Dim arryVarLO(0)          
    arryVarLO(0) = ACTION_LOGOUT
    Dim strFN
    Dim strLN  
    Dim blnFirstTime
    Dim strStatusCode 
    blnFirstTime = Session("sess_blnFirstTime")
    arryVarLO(0) = ACTION_LOGOUT
    strFN = Session("sess_strFirstName")
    strLN = Session("sess_strLastName")
    Dim lngRC
    Dim lngProjectId
                      
    Select Case strOption
      
      Case OPTION_BEFOREQUESTION, OPTION_FREEZE:     
          Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">" & fn_InsertSpaces(20) & "Welcome " & strFN & " " & strLN & "</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVarLO) & """>LogOut</a>]</td>"
      
      Case OPTION_AFTERANSWER:
          lngProjectId = Session("sess_lngProjectId")
          lngRC = fn_BL_UpdateFirstTime(lngProjectId)
          If lngRC <> 0 Then
             Call sub_HandleLogicError(ERR_EDIT_USERDATA)
          End If
          Session("sess_blnFirstTime") = False
          strStatusCode = STATUS_PEND
          lngRC = fn_BL_UpdateProjectStatus(lngProjectId, strStatusCode)
          If lngRC <> 0 Then
             Call sub_HandleLogicError(ERR_EDIT_USERDATA)
          End If

          Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">" & fn_InsertSpaces(20) & "</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVarLO) & """>LogOut</a>]</td>"
      
      Case Else:    
          Call sub_HandleLogicError(ERR_INVALID_ACTION)
          
    End Select  
  %>      
     </tr>   
   </table>
   <%
      If blnFirstTime = True Then
         Response.Write "<span class=""smallsubtitle""><br><br>Project Welcome</span>"
      End If
   %>
   <table>     

    <tr>
       <td colspan=3>
         <br><br>
         <table>
          <tr>
            <td align=left>
<%
   
    Dim arryVarQ(0)
    Dim strBlurb

    Select Case strOption
      
      Case OPTION_BEFOREQUESTION:
          Response.Write "<font color=#4A75BC size=2><b>"
          Response.Write "<p>Welcome! This password protected area is your project portal and will be the place where you will work with us to make your vision into a reality. "
          Response.Write "<p>This portal will guide you through the different phases of our development process by providing you with important messages about the status of your project, critical instructions on what to do and the tools to help <i>you</i> provide <i>us</i> with the specific things we need to know to give you exactly what you want."
          Response.Write "<p>Now, since this your first time to use the portal, let's move forward to the first step of our process. You will first need to answer some initial questions about your project by clicking the link below."
          Response.Write "</font><br><br>"

          arryVarQ(0) = ACTION_QUESTION
          Response.Write "To Continue to the Project Questions Page [<a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVarQ) & """>Click Here</a>]."
        
      Case OPTION_AFTERANSWER:
          strStatusCode = STATUS_PEND
          
          lngRC = fn_BL_GetProjectStatusBlurb(strStatusCode, strBlurb)
          If lngRC <> 0 Then
             Call sub_HandleLogicError(ERR_GET_USERDATA_SPEC)
          End If
          Response.Write "<font color=#4A75BC size=2>Your questions were submitted successfully!<br>" & strBlurb & "</font><br><br>"
          
          'arryVarQ(0) = ACTION_PROJECTSTATUS
          'Response.Write "To continue to the Project Main Page <a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVarQ) & """>Click Here</a>."
      
      Case OPTION_FREEZE:
          Dim arryTmp  
          arryTmp = Session("sess_arryProjectInfo")
          strStatusCode = arryTmp(13,0)
          lngRC = fn_BL_GetProjectStatusBlurb(strStatusCode, strBlurb)
          If lngRC <> 0 Then
             Call sub_HandleLogicError(ERR_GET_USERDATA_SPEC)
          End If
          Response.Write "<font>" & strBlurb & "</font><br><br>"
          
          Erase arryTmp
                       
      Case Else:    
          Call sub_HandleLogicError(ERR_INVALID_ACTION)
          
    End Select  
%>                        

           </td>
         </tr>
       </table>
      
      </td>
    </tr>  
   </table>

 <!-- </td>
 </tr>
</table>  -->
<% 
End Sub           'DisplayProjectMessage



'Sub DisplayProjectQuestions
'*********************************************
Sub DisplayProjectQuestions()
Dim arryVar(0)
Dim arryVar2(0)
Dim lngProjectId
   
arryVar(0) = ACTION_VALIDATEQUESTION
%>
<!-- <table width="600">
  <tr>
    <td align="left"> -->
   <form action="<% = fn_CreateURL(True,PAGE_PORTAL,arryVar) %>" method="post" name="theForm">

   <table> 
     <tr>
  <%   
    arryVar2(0) = ACTION_LOGOUT
    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"

  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       'Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="top">
    <span class="smallsubtitle"><font color="#4A75BC" size="2"><b>Project Initial Questions</b></font></span></td></tr>
    <tr><td colspan="3"><b><font color="red" size="2">Please try to answer all of the following questions. If you feel that the question does not apply to your project then just type in "NA" for your answer.</font></b>
     </td>
     </tr>   
   </table>
   <table>     

    <tr>
       <td colspan=3>
         <br>
         <table width=200>
          <tr>
            <td align=left>
            <%
              Dim arryQuestions
              Dim lngRC
              lngProjectId = Session("sess_lngProjectId")
              lngRC = fn_BL_GetProjectQuestions (lngProjectId, arryQuestions)
              If lngRC <> 0 Then
                   Call sub_HandleLogicError(lngRC)
              End If
              
              If IsArray(arryQuestions) Then
                 
                 Session("sess_arryQuestions") = arryQuestions
                 
                 Dim i
                 Dim intUB
                 intUB = UBound(arryQuestions,2)
                 For i = 0 to intUB
                    If i <> 0 Then
                       Response.Write "<br>"
                    End If
                    Response.Write "<font color=#0000FF size=2>" & arryQuestions(0,i) & ":</font>" & vbCrLf
                    Response.Write "<textarea name=""txtAnswer" & CStr(i+1) & """ cols=65 rows=4>" & arryQuestions(3,i) & "</textarea>" & vbCrLf
                    If i <> intUB Then
                       Response.Write "<br>"
                    End If
                    
                 Next
                 Response.Write "<input type=""hidden"" name=""intQuestionCount"" value=""" & CStr(intUB + 1) & """>"
                 Response.Write "<br><br>" & vbCrLf
                 Response.Write "<input type=""submit"" value=""Submit""  onmouseover=""this.className='submitbuttonon2'"" onmouseout=""this.className='submitbutton2'"" class=""submitbutton2"">"
                 
                 Erase arryQuestions
                 
              End If
              
            %>
           </td>
         </tr>
       </table>
      
      </td>
    </tr>  
   </table>

   </form>
 <!-- </td>
 </tr>
</table>  -->
<% 
End Sub           'DisplayProjectQuestions





'Sub Validate Question
'*********************************************
Sub sub_ValidateQuestion()
    
   Dim lngProjectId    
   Dim arryQuestions
   Dim arryVar(1)
   ReDim arryAns(0)
   Dim intQC
   Dim lngRC
   Dim i
        
   intQC = Request.Form("intQuestionCount")
   intQC = CInt(intQC)
   
   If intQC > 1 Then
      ReDim arryAns(intQC - 1)
   End If
   For i = 0 to (intQC - 1)
      arryAns(i) = Request.Form("txtAnswer" & CStr(i+1))
   Next
   
   lngProjectId = Session("sess_lngProjectId")
   arryQuestions = Session("sess_arryQuestions")
   
   lngRC = fn_BL_InsertProjectAnswers(lngProjectId, arryQuestions, arryAns)
   If lngRC <> 0 Then
      Call sub_HandleLogicError(lngRC)
   End If
   
   arryVar(0) = ACTION_PROJECTMESSAGE
   arryVar(1) = OPTION_AFTERANSWER
   Response.Clear
   Response.Redirect fn_CreateURL(True,PAGE_PORTAL,arryVar)
                    
       
End Sub    ' sub_ValidateQuestion







'Sub DisplayProjectRequirements
'*********************************************
Sub DisplayProjectRequirements()
Dim arryVar(1)
Dim arryVar2(1)
   

%>

   <table height="50"> 
   <tr>
  <%   
    arryVar2(0) = ACTION_LOGOUT
    arryVar2(1) = fn_GetRandomAlphaNumeric(11,16)

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"

  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="top">
    <%
       Dim arryVarAN(0)
       arryVarAN(0) = ACTION_ADDNEW_REQUIREMENT
       Response.Write "<span class=""smallsubtitle"">Project Requirements</span>" & fn_InsertSpaces(20) & "<a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVarAN) & """ class=""biglink2"">Add New Requirement</a><br><br>"
    %>
     </td>
     </tr>   
   </table>
   <table>     

    <tr>
       <td>
         <table>
             <tr>
                <td colspan="3">

                <table width="500" border="1">
                  <tr>
                     <td width="125" align="center"><b>Status</b></td><td width="375" align="center"><b>Requirement Title</b></td>
                  </tr>
                </table>
                <br>
                </td>
             </tr>
            <%
              Dim arryReqs
              Dim lngRC
              Dim lngProjectId
              
              lngProjectId = Session("sess_lngProjectId")
              lngRC = fn_BL_GetRequirementDescriptions (lngProjectId, arryReqs)
              If lngRC <> 0 Then
                   Call sub_HandleLogicError(lngRC)
              End If
              
              If IsArray(arryReqs) Then
                 arryVar(0) = ACTION_EDIT_REQUIREMENT
                 Dim i
                 Dim intUB
                 intUB = UBound(arryReqs,2)
                 For i = 0 to intUB
                    arryVar(1) = arryReqs(0,i)
                    Response.Write "<tr>" & vbCrLf
                    Response.Write "<td align=""center"" width=""125""><font size=""2"">" & arryReqs(2,i) & "</font></td><td width""30""><img src=""picts/spacer.gif"" width=""30""></td><td width=""375"" align=""left""><font size=""2""><a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVar) & """ class=""biglink2"">" & arryReqs(1,i) & "</a></font></td>" & vbCrLf

                    Response.Write "</tr>" & vbCrLf
                 Next

                 Erase arryReqs
              End If
              
            %>
       </table>
      
      </td>
    </tr>  
   </table>
 
<% 
End Sub           'DisplayProjectRequirements



'Sub DisplayRequirementEdit
'*********************************************
Sub DisplayRequirementEdit(lngReqId)
Dim arryVar(1)
Dim arryVar2(0)
Dim arryVarVE(0)
arryVar2(0) = ACTION_LOGOUT
arryVarVE(0) = ACTION_VALIDATE_EDIT_REQUIREMENT
%>
   <table> 
     <tr>
  <%   

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"
  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="center">
     <table>
     <tr>
     
    <%
       Dim arryVarRAF(1)
       arryVarRAF(0) = ACTION_REQUIREMENT_ATTACH_FILE
       arryVarRAF(1) = lngReqId
       Response.Write "<td nowrap><span class=""smallsubtitle"">Requirement Detail</span>" & fn_InsertSpaces(60) & "</td><td><a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVarRAF) & """ class=""biglink2""><img src=""picts/attach.gif"" border=""0""></a></td><td><a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVarRAF) & """ >View Attached Files</a></td>"
    %>
     </tr>
     </table>
     </td>
     </tr>   
   </table>
   <form name="form1" action="<% = fn_CreateUrl(True,PAGE_PORTAL,arryVarVE) %>" method="post">
   <table cellspacing="5">     

    

            <%
              Dim arryReq
              Dim lngRC
              
              lngRC = fn_BL_GetRequirementDetail (lngReqId, arryReq)
              If lngRC <> 0 Then
                   Call sub_HandleLogicError(lngRC)
              End If
              
              Dim strName
              Dim strSummary
              Dim dtLastUpdate
              Dim dtApproved
              Dim dtChadDanApproved
              Dim strStatus
              Dim dtComplete
                               
              If IsArray(arryReq) Then
                 
                 strName = arryReq(0,0)
                 Session("sess_strReqTitle") = strName
                 strSummary = arryReq(1,0)
                 dtLastUpdate = arryReq(2,0)
                 dtApproved = arryReq(3,0)
                 dtChadDanApproved = arryReq(4,0)
                 strStatus = arryReq(5,0)
                 dtComplete = arryReq(6,0)
                 
                 Erase arryReq
              End If
                 

              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Title:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><input type=""text"" name=""txtName"" size=""40"" maxlength=""50"" value=""" & strName & """></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Status:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><font size=""2"">" & strStatus & "</font></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf            
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td valign=""top"" align=""right""><font size=""2""><b>Summary:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td colspan=""2"" width=""375"" align=""left""><textarea name=""txtSummary"" cols=45 rows=5>" & strSummary & "</textarea></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Date Last Updated:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><font size=""2"">" & dtLastUpdate & "</font></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Date Approved:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><font size=""2"">" & dtApproved & "</font></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Date Fulfilled:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><font size=""2"">" & dtComplete & "</font></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              
            %>
    <tr>
      <td align="center" colspan="3">
        <input type="hidden" name="txtReqId" value="<% = lngReqId %>">
        <input type="submit" name="cmdSubmit" value="Update"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="submit" name="cmdSubmit" value="Delete"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      
      </td>
    </tr>
 
   </table>
   </form>
 
<% 
End Sub           'DisplayRequirementEdit


'Sub Validate Requirement
'*********************************************
Sub sub_ValidateRequirementEdit()
    
   Dim lngProjectId    
   Dim arryVar(0)
        
   Dim strName
   Dim strSummary
   Dim lngRC
   Dim lngReqId
   Dim strSubmit
   
   strSubmit = Request.Form("cmdSubmit")
   
   Select Case strSubmit
   
      Case "Delete":
         lngReqId = Request.Form("txtReqId")
   
         lngRC = fn_BL_DeleteRequirement (lngReqId)
         If lngRC <> 0 Then
            Call sub_HandleLogicError(ERR_DELETE_USERDATA)
         End If      
      
      Case "Update":
         strName = Request.Form("txtName")
         strSummary = Request.Form("txtSummary")
         lngReqId = Request.Form("txtReqId")
   
         lngRC = fn_BL_EditRequirement (lngReqId, strName, strSummary)
         If lngRC <> 0 Then
            Call sub_HandleLogicError(ERR_EDIT_USERDATA)
         End If
   
      Case Else:
         Call sub_HandleLogicError(ERR_INVALID_ACTION)
        
   End Select
   
   arryVar(0) = ACTION_REQUIREMENT
   Response.Redirect fn_CreateUrl(False,PAGE_PORTAL,arryVar)
                    
       
End Sub    ' sub_ValidateRequirementEdit




'Sub DisplayRequirementAddNew
'*********************************************
Sub DisplayRequirementAddNew()
Dim arryVar(1)
Dim arryVar2(0)
Dim arryVarVAN(0)
arryVar2(0) = ACTION_LOGOUT
arryVarVAN(0) = ACTION_VALIDATE_ADDNEW_REQUIREMENT
%>
   <table> 
     <tr>
  <%   

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"
  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="top">
    <span class="smallsubtitle">Add New Requirement</span>
     </td>
     </tr>   
   </table>
   <form name="form1" action="<% = fn_CreateUrl(True,PAGE_PORTAL,arryVarVAN) %>" method="post">
   <table cellspacing="5">     
            <%
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Title:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><input type=""text"" name=""txtName"" size=""40"" maxlength=""50""></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf           
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td valign=""top"" align=""right""><font size=""2""><b>Summary:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td colspan=""2"" width=""375"" align=""left""><textarea name=""txtSummary"" cols=45 rows=5></textarea></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
            %>
    <tr>
      <td align="center" colspan="3">
        <input type="submit" name="cmdSubmit" value="Submit"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      </td>
    </tr>
 
   </table>
   </form>
 
<% 
End Sub           'DisplayRequirementAddNew


'Sub Validate Requirement Add New
'*********************************************
Sub sub_ValidateRequirementAddNew()
 
   Dim arryVar(0)
        
   Dim strName
   Dim strSummary
   Dim lngRC
   Dim lngProjectId      

   strName = Request.Form("txtName")
   strSummary = Request.Form("txtSummary")
   lngProjectId = Session("sess_lngProjectId")
   
   lngRC = fn_BL_AddNewRequirement (lngProjectId, strName, strSummary)
   If lngRC <> 0 Then
      Call sub_HandleLogicError(ERR_ADDNEW_USERDATA)
   End If
  
   
   arryVar(0) = ACTION_REQUIREMENT
   Response.Redirect fn_CreateUrl(False,PAGE_PORTAL,arryVar)
                    
       
End Sub    ' sub_ValidateRequirementAddNew


'Sub Display Requirement Attach File
'*********************************************
Sub DisplayRequirementAttachFile(lngReqId)
Dim arryVar(1)
Dim arryVar2(0)
Dim arryVarVR(1)
Dim arryVarRAF(0)
arryVar2(0) = ACTION_LOGOUT
arryVarRAF(0) = ACTION_VALIDATE_REQUIREMENT_ATTACH_FILE
arryVarVR(0) = ACTION_EDIT_REQUIREMENT
arryVarVR(1) = lngReqId
Session("sess_lngReqId") = lngReqId

%>
   <table> 
     <tr>
  <%   

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"
  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="top">
    <span class="smallsubtitle">Requirement - Attachments</span>
     </td>
     </tr>   
     <tr>
       <td colspan="3">
         <%
           Response.Write "Requirement Title : [ <a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVarVR) & """>" & Session("sess_strReqTitle") & "</a> ]" 
         %>
       </td>
     </tr>
   </table>
   <FORM METHOD="POST" ENCTYPE="multipart/form-data" name="form1" action="<% = fn_CreateUrl(True,PAGE_PORTAL,arryVarRAF) %>">
   <table cellspacing="5">     
            <%
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>File Description:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><input type=""text"" name=""txtFileDescrip"" size=""40"" maxlength=""50""></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf           
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td valign=""top"" align=""right""><font size=""2""><b>Choose A File:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td colspan=""2"" width=""375"" align=""left""><input type=""FILE"" size=""50"" name=""FILE1""></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
            %>
    <tr>
      <td align="center" colspan="3">
        <input type="hidden" name="txtReqId" value="<% = lngReqId %>">
        <input type="submit" name="cmdSubmit" value="Attach" onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      </td>
    </tr>
 
   </table>
   </form>
   <table>
   <%
   
    Dim arryFiles
    Dim lngRC
    Dim i
    
    lngRC = fn_BL_GetRequirementFiles (lngReqId, arryFiles)
    If lngRC <> 0 Then
         Call sub_HandleLogicError(lngRC)
    End If
    
    If IsArray(arryFiles) Then 
       Dim arryVarDRF(2)
       Dim arryVarDF(3)
       arryVarDRF(0) = ACTION_DELETE_REQUIREMENT_FILE
       arryVarDF(0) = ACTION_DOWNLOAD_REQUIREMENT_FILE
       For i = 0 to UBound(arryFiles,2)
          arryVarDRF(1) = CStr(arryFiles(0,i))    
          arryVarDRF(2) = CStr(arryFiles(2,i))    
          arryVarDF(1) = CStr(arryFiles(0,i))  
          arryVarDF(2) = CStr(arryFiles(2,i))  
          arryVarDF(3) = CStr(arryFiles(1,i))   
          Response.Write "<tr>" & vbCrLf
          Response.Write "   <td>[<a onclick=""return confirmDelete()"" href=""" & fn_CreateUrl(True,PAGE_PORTAL,arryVarDRF) & """ title=""Delete"" alt=""Delete"" onMouseOver=""window.status=this.alt;return true;"" onMouseOut=""window.status='';return true;""><font color=""RED""><b>X</b></font></a>]&nbsp;&nbsp;</td><td>[<a href=""" & fn_CreateUrl(True,PAGE_PORTAL,arryVarDF) & """ title=""Download"" alt=""Download"" onMouseOver=""window.status=this.alt;return true;"" onMouseOut=""window.status='';return true;"">" & arryFiles(1,i) & "</a>] - " & arryFiles(3,i) & "</td>" & vbCrLf
          Response.Write "</tr>" & vbCrLf
       Next
    Else
          Response.Write "<tr>" & vbCrLf
          Response.Write "   <td>No Files Attached</td>" & vbCrLf
          Response.Write "</tr>" & vbCrLf
    End If 
   %>
   
   </table>
 
<% 
End Sub           'DisplayRequirementAttachFile


'Sub Validate Requirement Attach File
'*********************************************
Sub sub_ValidateRequirementAttachFile()
 
   Dim arryVar(1) 
   Dim lngRC
   Dim lngProjectId      

   lngProjectId = Session("sess_lngProjectId")
       
    'NOTE - YOU MUST HAVE VBSCRIPT v5.0 INSTALLED ON YOUR WEB SERVER
    '	   FOR THIS LIBRARY TO FUNCTION CORRECTLY. YOU CAN OBTAIN IT
    '	   FREE FROM MICROSOFT WHEN YOU INSTALL INTERNET EXPLORER 5.0
    '	   OR LATER.


    ' Create the FileUploader
    Dim Uploader, File
    Set Uploader = New FileUploader
    
    ' This starts the upload process
    Uploader.Upload()
    
    '******************************************
    ' Use [FileUploader object].Form to access 
    ' additional form variables submitted with
    ' the file upload(s). (used below)
    '******************************************

    ' Check if any files were uploaded
    If Uploader.Files.Count = 0 Then
    	' Do Nothing
    Else
        Dim strPath
        Dim strNewFileName
        Dim strDescrip
        Dim lngReqId
        
        lngReqId = Uploader.Form("txtReqId")
        strDescrip = Uploader.Form("txtFileDescrip")
        strPath = Replace(Server.MapPath(PAGE_PORTAL),PAGE_PORTAL,"") & FOLDER_UPLOAD & "\"
        
    	' Loop through the uploaded files
    	For Each File In Uploader.Files.Items
    		
		'Save the file
                lngRC = fn_BL_AddNewRequirementFile (lngReqId, strDescrip, File.FileName, File.FileSize, File.ContentType, strNewFileName)
                If lngRC <> 0 Then
                   Call sub_HandleLogicError(ERR_ADDNEW_USERDATA)
                End If
		File.FileName = strNewFileName
		File.SaveToDisk strPath
    	
    	Next
    	
    End If
  
   
   arryVar(0) = ACTION_REQUIREMENT_ATTACH_FILE
   arryVar(1) = lngReqId
   Response.Redirect fn_CreateUrl(False,PAGE_PORTAL,arryVar)
                    
       
End Sub    ' sub_ValidateRequirementAttachFile



'Sub Validate Delete Requirement File
'*********************************************
Sub sub_ValidateDeleteRequirementFile(lngRFId, strFileName)
 
   Dim arryVar(1)
        
   Dim lngRC
   Dim lngReqId   

   lngReqId = Session("sess_lngReqId")
   
   lngRC = fn_BL_DeleteRequirementFile (lngReqId, lngRFId, strFileName)
   If lngRC <> 0 Then
      Call sub_HandleLogicError(ERR_DELETE_USERDATA)
   End If
  
   
   arryVar(0) = ACTION_REQUIREMENT_ATTACH_FILE
   arryVar(1) = lngReqId
   Response.Redirect fn_CreateUrl(True,PAGE_PORTAL,arryVar)
                    
       
End Sub    ' sub_ValidateDeleteRequirementFile



'Sub Validate Download Requirement File
'*********************************************
Sub sub_ValidateDownloadRequirementFile(lngRFId, strFileName, strDisplayName)
         
   Dim lngRC
   Dim lngReqId   

   lngReqId = Session("sess_lngReqId")   
   lngRC = fn_BL_CheckRequirementFile (lngReqId, lngRFId)
   If lngRC <> 0 Then
      Call sub_HandleLogicError(lngRC)
   End If
   
   'PrintStop Replace(Server.MapPath(PAGE_PORTAL),PAGE_PORTAL,"") & FOLDER_UPLOAD & "\" & strFileName
   Call sub_DownloadUploadedFile (Replace(Server.MapPath(PAGE_PORTAL),PAGE_PORTAL,"") & FOLDER_UPLOAD & "\" & strFileName, strDisplayName)                    
       
End Sub    ' sub_ValidateDownloadRequirementFile



'Sub Display Project Issues
'*********************************************
Sub DisplayProjectIssues()
Dim arryVar(1)
Dim arryVar2(0)
arryVar(0) = ACTION_ISSUE
arryVar2(0) = ACTION_LOGOUT

%>
   <table> 
     <tr>
  <%   

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"
  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="top">
    <span class="smallsubtitle">Project Issues</span>
     </td>
     </tr>   
     <tr>
       <td colspan="3">
       </td>
     </tr>
   </table>
     <FORM name="form22" METHOD="POST" action="<% = fn_CreateUrl(True,PAGE_PORTAL,arryVar) %>">
   <table cellspacing="5">     
 <%
   
    Dim arryTypes
    Dim lngRC
    Dim i
    
    lngRC = fn_BL_GetProjectIssueTypes (arryTypes)
    If lngRC <> 0 Then
         Call sub_HandleLogicError(lngRC)
    End If
    Session("sess_arryIssueTypes") = arryTypes
    
    
    If IsArray(arryTypes) Then 
       Response.Write "<select name=""lstIssueTypes"">" & vbCrLf
       Response.Write "<option value=""none"" selected> &lt;Choose An Issue Type&gt;" & vbCrLf
       For i = 0 to Ubound(arryTypes,2)
          Response.Write "<option value=""" & arryTypes(0,i) & """>" & arryTypes(1,i) & vbCrLf
       Next    
       Response.Write "</select>" & vbCrLf        
    Else
       Response.Write "<tr>" & vbCrLf
       Response.Write "   <td></td>" & vbCrLf
       Response.Write "</tr>" & vbCrLf
    End If 
   %>
    <tr>
      <td align="center" colspan="3">
        <input onClick="return CheckChoose(document.form22.lstIssueTypes.value,'none','Please choose an issue type.')" type="submit" name="cmdSubmit" value="Go" onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      </td>
    </tr>
 
   </table>
   </form> 
 
<% 
End Sub           'DisplayProjectIssues


'Sub Display Issue
'*********************************************
Sub DisplayIssue()
Dim arryVar(2), arryVar2(0), strIssueType, strTypeDesc
'arryVar(0) = 
arryVar2(0) = ACTION_LOGOUT

strIssueType = Request.Form("lstIssueTypes")
If strIssueType = "" Then
   strIssueType = Request("x2")
End If
strTypeDesc = fn_GetIssueTypeDescrip(strIssueType)
%>
   <table> 
     <tr>
  <%   

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"
  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="top">
    <%
       Dim arryVarAN(1)
       arryVarAN(0) = ACTION_ADDNEW_ISSUE
       arryVarAN(1) = strIssueType
       Response.Write "<span class=""smallsubtitle"">Project Issues - " & strTypeDesc & "</span>" & fn_InsertSpaces(20) & "<a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVarAN) & """ class=""biglink2"">Add New Issue</a><br><br>"
    %>
   </td>
   </tr>   
   <tr>
      <td colspan="3">
      </td>
   </tr>  
   
   </table>
 
   <table>     
  <tr>
  <td>
  <table>
      <tr>
         <td colspan="3">

         <table width="500" border="1">
           <tr>
              <td width="125" align="center"><b>Status</b></td><td width="375" align="center"><b>Issue Title</b></td>
           </tr>
         </table>
         <br>
         </td>
      </tr>
<%
   
    Dim arryIss
    Dim lngRC
    Dim i
    Dim strResolvedDate
    
    lngRC = fn_BL_GetProjectIssueDescriptions (Session("sess_lngProjectId"), strIssueType, arryIss)
    If lngRC <> 0 Then
         Call sub_HandleLogicError(lngRC)
    End If
              
    If IsArray(arryIss) Then
       arryVar(0) = ACTION_EDIT_ISSUE
       Dim intUB
       intUB = UBound(arryIss,2)
       For i = 0 to intUB
          arryVar(1) = arryIss(0,i)
          arryVar(2) = strIssueType  
          strResolvedDate = arryIss(2,i)
          If IsNull(strResolvedDate) or strResolvedDate = "" Then
             strResolvedDate = "Unresolved"
          Else
             strResolvedDate = "Resolved"
          End If
          Response.Write "<tr>" & vbCrLf
          Response.Write "<td align=""center"" width=""125""><font size=""2"">" & strResolvedDate & "</font></td><td width""30""><img src=""picts/spacer.gif"" width=""30""></td><td width=""375"" align=""left""><font size=""2""><a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVar) & """ class=""biglink2"">" & arryIss(1,i) & "</a></font></td>" & vbCrLf
          Response.Write "</tr>" & vbCrLf
       Next

       Erase arryIss
    End If
              
%>
   </table>
      
   </td>
   </tr>  
 
   </table>
 
<% 
End Sub           'DisplayIssue


'Sub DisplayIssueEdit
'*********************************************
Sub DisplayIssueEdit(lngIssueId)
Dim arryVar(1)
Dim arryVar2(0)
Dim arryVarVE(0)
Dim strIssueType, strTypeDesc
arryVar2(0) = ACTION_LOGOUT
arryVarVE(0) = ACTION_VALIDATE_EDIT_ISSUE
strIssueType = Request("x3")
%>
   <table> 
     <tr>
  <%   

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"
  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr>
   <td colspan="3" valign="center">
     <table>
     <tr>
     
    <%
       Dim arryVarI(1)
       arryVarI(0) = ACTION_ISSUE
       arryVarI(1) = strIssueType
       strTypeDesc = fn_GetIssueTypeDescrip(strIssueType)
       Response.Write "<td nowrap><span class=""smallsubtitle"">Project Issues - [<a href=""" & fn_CreateUrl(True,PAGE_PORTAL,arryVarI) & """>" & strTypeDesc & "</a>] - Issue Detail</span></td>"
    %>
     </tr>
     </table>
     </td>
     </tr>   
   </table>
   <form name="form1" action="<% = fn_CreateUrl(True,PAGE_PORTAL,arryVarVE) %>" method="post">
   <table cellspacing="5">     

    

            <%
              Dim arryIssue
              Dim lngRC
              
              lngRC = fn_BL_GetIssueDetail (lngIssueId, arryIssue)
              If lngRC <> 0 Then
                   Call sub_HandleLogicError(lngRC)
              End If
              
              Dim strTitle
              Dim strDetails
              Dim dtResolved
                               
              If IsArray(arryIssue) Then
                 
                 strTitle = arryIssue(0,0)
                 'Session("sess_strIssueTitle") = strTitle
                 strDetails = arryIssue(1,0)
                 dtResolved = arryIssue(2,0)
                 
                 Erase arryIssue
              End If
                 
              Dim strChecked
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Title:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><input type=""text"" name=""txtTitle"" size=""40"" maxlength=""50"" value=""" & strTitle & """></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              Response.Write "<tr>" & vbCrLf
              If Not IsNull(dtResolved) and dtResolved <> "" Then
                 strChecked = "checked"
              Else 
                  dtResolved = "Unresolved"  
              End If
              'Response.Write "<td colspan=""2""><font size=""2""><b>Resolved:</b></font>&nbsp;&nbsp;<input type=""checkbox"" name=""chkResolved"" value=""1"" " & strChecked & "></td>" & vbCrLf
              Response.Write "<td><font size=""2""><b>Resolved Y/N:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""350"" align=""left""><input type=""checkbox"" name=""chkResolved"" value=""1"" " & strChecked & "></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td><font size=""2""><b>Date Resolved:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""350"" align=""left""><font size=""2"">" & dtResolved & "</font></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf            
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td valign=""top"" align=""right""><font size=""2""><b>Details:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><textarea name=""txtDetails"" cols=45 rows=5>" & strDetails & "</textarea></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              
            %>
    <tr>
      <td align="center" colspan="4">
        <input type="hidden" name="txtIssueId" value="<% = lngIssueId %>">
        <input type="hidden" name="txtIssueType" value="<% = strIssueType %>">
        <input type="submit" name="cmdSubmit" value="Update"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="submit" name="cmdSubmit" value="Delete"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      
      </td>
    </tr>
 
   </table>
   </form>
 
<% 
End Sub           'DisplayIssueEdit


'Sub Validate Issue
'*********************************************
Sub sub_ValidateIssueEdit()
    
   Dim lngProjectId, arryVar(1), lngRC, lngIssueId, strSubmit, strIssueType, strTitle, strDetails, blnResolved
   
   strSubmit = Request.Form("cmdSubmit")
   lngIssueId = Request.Form("txtIssueId")
   strIssueType = Request.Form("txtIssueType")
   
   Select Case strSubmit
   
      Case "Delete":  
         lngRC = fn_BL_DeleteIssue (lngIssueId)
         If lngRC <> 0 Then
            Call sub_HandleLogicError(ERR_DELETE_USERDATA)
         End If      
      
      Case "Update":
         strTitle = Request.Form("txtTitle")
         strDetails = Request.Form("txtDetails")
         blnResolved = Request.Form("chkResolved")
   
         lngRC = fn_BL_EditIssue (lngIssueId, strTitle, strDetails, blnResolved)
         If lngRC <> 0 Then
            Call sub_HandleLogicError(ERR_EDIT_USERDATA)
         End If
   
      Case Else:
         Call sub_HandleLogicError(ERR_INVALID_ACTION)
        
   End Select
   
   arryVar(0) = ACTION_ISSUE
   arryVar(1) = strIssueType
   Response.Redirect fn_CreateUrl(False,PAGE_PORTAL,arryVar)
                    
       
End Sub    ' sub_ValidateIssueEdit




'Sub DisplayIssueAddNew
'*********************************************
Sub DisplayIssueAddNew()
Dim arryVar(1)
Dim arryVar2(0)
Dim arryVarVAN(0)
Dim strIssueType
arryVar2(0) = ACTION_LOGOUT
arryVarVAN(0) = ACTION_VALIDATE_ADDNEW_ISSUE

strIssueType = Request("x2")
%>
   <table> 
     <tr>
  <%   

    Response.Write "<td width=""125""><span class=""smalltitle"">Client Portal</span></td><td width=""400"">&nbsp;</td><td>[<a href=""" & fn_CreateURL(True,PAGE_LOGIN,arryVar2) & """>LogOut</a>]</td>"
  %>      
   </tr>
   <tr>
   <td colspan="3">
    <%
       Call DisplayProjectNavLinks
    %>
    </td>
   </tr>
   <tr> 
    <%
       Dim strTypeDesc
       Dim arryVarI(1), arryTypes, i
       
       strTypeDesc = fn_GetIssueTypeDescrip(strIssueType)
       arryVarI(0) = ACTION_ISSUE
       arryVarI(1) = strIssueType
       Response.Write "<td nowrap><span class=""smallsubtitle"">Project Issues - </span><b>[<a href=""" & fn_CreateUrl(True,PAGE_PORTAL,arryVarI) & """>" & strTypeDesc & "</a>]</b><span class=""smallsubtitle""> - Add New Issue</span></td>"
    %> 
     
     </tr>   
   </table>
   <form name="form1" action="<% = fn_CreateUrl(True,PAGE_PORTAL,arryVarVAN) %>" method="post">
   <table cellspacing="5">     
            <%
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td align=""right""><font size=""2""><b>Title:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td width=""375"" align=""left""><input type=""text"" name=""txtTitle"" size=""40"" maxlength=""50""></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf           
              Response.Write "<tr>" & vbCrLf
              Response.Write "<td valign=""top"" align=""right""><font size=""2""><b>Details:</b></font></td><td width""10""><img src=""picts/spacer.gif"" width=""10""></td><td colspan=""2"" width=""375"" align=""left""><textarea name=""txtDetails"" cols=45 rows=5></textarea></td>" & vbCrLf
              Response.Write "</tr>" & vbCrLf
              
            %>
    <tr>
      <td align="center" colspan="3">
        <%
          Response.Write "<input type=""hidden"" name=""txtIssueType"" value=""" & strIssueType & """>"
        %>
        <input type="submit" name="cmdSubmit" value="Submit"  onmouseover="this.className='submitbuttonon2'" onmouseout="this.className='submitbutton2'" class="submitbutton2">
      </td>
    </tr>
 
   </table>
   </form>
 
<% 
End Sub           'DisplayIssueAddNew


'Sub Validate Requirement Add New
'*********************************************
Sub sub_ValidateIssueAddNew()
 
   Dim arryVar(1)
        
   Dim strTitle, strDetails, strIssueType
   Dim lngRC
   Dim lngProjectId      

   strIssueType = Request("txtIssueType")
   strTitle = Request.Form("txtTitle")
   strDetails = Request.Form("txtDetails")
   lngProjectId = Session("sess_lngProjectId")
   
   lngRC = fn_BL_AddNewIssue (lngProjectId, strIssueType, strTitle, strDetails)
   If lngRC <> 0 Then
      Call sub_HandleLogicError(ERR_ADDNEW_USERDATA)
   End If
  
   arryVar(0) = ACTION_ISSUE
   arryVar(1) = strIssueType
   Response.Redirect fn_CreateUrl(False,PAGE_PORTAL,arryVar)
                    
       
End Sub    ' sub_ValidateIssueAddNew


Sub DisplayProjectNavLinks ()
%>
   <IMG SRC="picts/paragraph-line.jpg" HEIGHT=3 WIDTH=100% border="0" vspace="7"><BR>
     <table height="15">
       <tr>
         <td>
           <%
             Dim arryVar1(1)
             arryVar1(0) = ACTION_PROJECTSTATUS
             arryVar1(1) = 0
             Response.Write "[ <a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVar1) & """>Main</a> ]" & fn_InsertSpaces(7) 
             Dim arryVar2(0)
             arryVar2(0) = ACTION_REQUIREMENT
             Response.Write "[ <a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVar2) & """>Requirements</a> ]" & fn_InsertSpaces(7) 
             Dim arryVar3(1)
             arryVar3(0) =  ACTION_ISSUES_MAIN
             arryVar3(1) = 0
             Response.Write "[ <a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVar3) & """>Issues</a> ]" & fn_InsertSpaces(7)
             'Dim arryVar3(1)
             'arryVar3(0) =  ACTION_QUESTION
             'arryVar3(1) = 0
             'Response.Write "[ <a href=""" & fn_CreateURL(True,PAGE_PORTAL,arryVar3) & """>Questions</a> ]" 
           %>
         </td>
       </tr>
     </table>
   <IMG SRC="picts/paragraph-line.jpg" HEIGHT=3 WIDTH=100% border="0" vspace="7">
<%
End Sub

' Get Issue Type Description
'***********************************************
Function fn_GetIssueTypeDescrip(strIssueType)

   Dim arryTypes, i, strTypeDesc
       
   arryTypes = Session("sess_arryIssueTypes")
   For i = 0 to UBound(arryTypes,2)
      If strIssueType = arryTypes(0,i) Then
         strTypeDesc = arryTypes(1,i)
         Exit For
      End If
   Next
       
   fn_GetIssueTypeDescrip = strTypeDesc  
       
End Function       


' Call the Main Sub Routine
'*******************************
Call Main


%>


