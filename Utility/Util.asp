<!--#include file="Constants.asp"-->
<!--#include file="Random.asp"-->
<!--#include file="EncryptDecrypt.asp"-->
<%

' This function returns a href based on the value of GLOB_PRODUCTION
'**********************************************************
Function fn_GetHomePageHREF()

   If GLOB_PRODUCTION = "yes" Then
      fn_GetHomePageHREF = HREF_PROD_HOME
   Else
      fn_GetHomePageHREF = HREF_TEST_HOME
   End If

End Function


' This function returns a href based on the value of GLOB_PRODUCTION
'**********************************************************
Function fn_GetSSLHREF(strPage)

   If GLOB_PRODUCTION = "yes" Then
      fn_GetSSLHREF = HREF_SSL & strPage
   Else
      fn_GetSSLHREF = strPage
   End If

End Function


' Create a URL string
'*******************************************************
Function fn_CreateURL(ByVal blnSSL, ByVal strPage, ByVal arryVar)
  
  Dim strTmp
  Dim i
    
  If blnSSL Then
    strPage = fn_GetSSLHREF(strPage) 
  End if
  
  If Not ISArray(arryVar) Then 
      strTmp = strPage
  Else
      strTmp = strPage & "?" & QSVAR & "1=" & arryVar(0)
      For i = 1 to UBound(arryVar)
         strTmp = strTmp & "&" & QSVAR & CStr(i+1) & "=" & arryVar(i)
      Next
  End If
  
  fn_CreateURL = strTmp
  
End Function


'Sub Display Header
'*********************************************
Sub DisplayHeader(strTitle, strPic, strLogo)
%><!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html><head>

<!-- CHANGE THE NEXT THREE LINES -->

<title><% = strTitle %></title>
<meta name="Description" content="Blue Elephant Web Development">
<meta name="KeyWords" content="Blue Elephant, website development, software development, database, design">

<!-- CHANGE THE ABOVE THREE LINES -->

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META http-equiv="Content-Language" content="en">
<META name="revisit-after" content="15 days">
<META name="robots" content="index, follow">
<META name="Rating" content="General">
<META name="Robots" content="All">



<link rel=StyleSheet href="./Utility/corporatestyle.css" type="text/css" media="screen">
<script language="JavaScript" src="./Utility/validation.js"></script>
<script language="JavaScript" src="./Utility/blockerror.js"></script>


</head>

<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#339966" VLINK="#663399" ALINK="#666666" leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0 marginheight=0 marginwidth=0>

<!-- PAGE TABLE-->
<TABLE cellpadding=0 cellspacing=0 border=0 width="100%" HEIGHT="100%"><tr><td ALIGN="LEFT" VALIGN="TOP">

<script language="JavaScript" src="./Utility/<% = strLogo %>"></script>

<img src="picts/stretchbar.jpg" width="100%" height="28"><br>


<TABLE cellpadding=0 cellspacing=0 border=0 width="100%"><tr><td width="720">
<IMG SRC="picts/elephant.gif" width="1002" height="80"><BR>
</td><td bgcolor="#000000" width="2">
<IMG SRC="picts/spacer.gif" width="2" height="80"><BR>
</td><td bgcolor="#000000" background="picts/top-picture-end.gif">
&nbsp;<br>
</td></tr><tr><td bgcolor="#999999" background="picts/bluebar.gif" colspan="3">
 <IMG SRC="picts/spacer.gif" HEIGHT=26 WIDTH=15 border="0"><br>
</td></tr></table>


<script language="JavaScript" src="./Utility/menu.js"></script>



<!-- OUTER TABLE-->
<TABLE cellpadding=0 cellspacing=0 border=0 width="800"><tr><td ALIGN="CENTER" VALIGN="TOP">

<!-- SPLIT TABLE   background="picts/background-main.jpg"  -->
<TABLE cellpadding=0 cellspacing=0 border=0 height="100%" width="100%">
  <tr>
    <td ALIGN=LEFT VALIGN=TOP height="100%">
      <table cellpadding=0 cellspacing=0 border=0 height="100%">
         <tr>
           <td>
             <img src="picts/fadein2.jpg" width="200" height="30" border=0><br>
           </td>
         </tr>  
         <tr>
           <td height="100%">
             <TABLE cellpadding=0 cellspacing=0 height="100%" background="picts/fademiddle.jpg">
               <tr>
                 <td>   
                    <img src="picts/fademiddle.jpg" WIDTH=200 height=350 border=0>
                 </td>
               </tr>
             </table>                  
           </td>
         </tr>   
         <tr>
           <td>                
             <img src="picts/fadeout2.jpg" width="200" height="30" border=0>
           </td>
         </tr>       
      </table>
      
    </td>
    <td ALIGN=CENTER VALIGN=TOP bgcolor="#FFFFFF">


<!-- CONTENT TABLE -->
<TABLE cellpadding=0 cellspacing=10 border=0 bgcolor="#FFFFFF"><tr><td ALIGN=LEFT VALIGN=TOP>
<% 
End Sub   'DisplayHeader


'Sub Display Footer
'*********************************************
Sub DisplayFooter()
%>


</td></tr> </table>
<!-- CONTENT TABLE -->

</td></tr></table>
<!-- SPLIT TABLE -->

</td></tr></table>
<!-- OUTER TABLE-->



<!-- PAGE TABLE-->
</TD></TR><TR><TD VALIGN="BOTTOM">

<img src="picts/stretchbar.jpg" width="100%" height="28"><br>


<!-- COPYRIGHT -->
<TABLE cellpadding=0 cellspacing=0 border=0 width="100%" background="picts/bottom-shade.gif"><tr><td WIDTH="10">
<IMG SRC="picts/spacer.gif" HEIGHT=45 WIDTH=10 border="0"><BR>

</td><td ALIGN="left" VALIGN="center">

<script language="JavaScript" src="./Utility/copyright.js"></script>


</td><td ALIGN="right" VALIGN="center">



</TD><td WIDTH="10">

<IMG SRC="picts/spacer.gif" HEIGHT=10 WIDTH=10 border="0"><BR>

</td></tr></table>
<!-- COPYRIGHT -->

</td></tr></table>
<!-- END PAGE TABLE-->

</BODY>
</HTML>
<% 
End Sub  'Display Footer



'Function fn_GetNowAsYYYYMMDD_HHMMSS
'*********************************************
Function fn_GetNowAsYYYYMMDD_HHMMSS(dtDateTime)
  fn_GetNowAsYYYYMMDD_HHMMSS = Year(dtDateTime) & "-" & Right("0" & Month(dtDateTime),2) & "-" & Right("0" & day(dtDateTime),2) & " " _
                  & Right("0" & Hour(dtDateTime),2) & ":" & Right("0" & Minute(dtDateTime),2) & ":" & Right("0" & Second(dtDateTime),2)
End Function


'Function fn_GetDateAsYYYYMMDD
'*********************************************
Function fn_GetDateAsYYYYMMDD(dtDate)
   fn_GetDateAsYYYYMMDD = Year(dtDate) & "-" & Right("0" & Month(dtDate),2) & "-" & Right("0" & Day(dtDate),2)
End Function


'Function fn_GetDateAsMMDDYYYY
'*********************************************
Function fn_GetDateAsMMDDYYYY(dtDate)
   fn_GetDateAsMMDDYYYY = Right("0" & Month(dtDate),2) & "/" & Right("0" & Day(dtDate),2) & "/" & Year(dtDate)
End Function


'Function fn_GetTimeAsHHMMSS
'*********************************************
Function fn_GetTimeAsHHMMSS()
  Dim dt1 'As DateTime
  dt1 = Time()
  fn_GetTimeAsHHMMSS = Right("0" & Hour(dt1),2) & ":" & Right("0" & Minute(dt1),2) & ":" & Right("0" & Second(dt1),2)
End Function




'  Remove any unneccassary Session Variables
'********************************************************
Sub sub_RemoveSessionVariables

     Dim arryTmp
     Dim strTmp
     
     arryTmp = Session("sess_arryCSE")
     If IsArray(arryTmp) Then
        Session.Contents.Remove("sess_arryCSE")
        Erase arryTmp
     End If
     
     
     arryTmp = Session("sess_arryMonthInfo")
     If IsArray(arryTmp) Then
        Session.Contents.Remove("sess_arryMonthInfo")
        Erase arryTmp
     End If
     
     
     arryTmp = Session("sess_arryDayInfo")
     If IsArray(arryTmp) Then
        Session.Contents.Remove("sess_arryDayInfo")
        Erase arryTmp
     End If
     
     'Session("sess_arryUserData") 
     
     strTmp = Session("sess_dtCurDate")
     If strTmp <> "" Then
        Session.Contents.Remove("sess_dtCurDate")
     End If

End Sub




'Sub Check For Login
'*********************************************
Function fn_CheckForLogin()

   fn_CheckForLogin = -1

   If Session("sess_blnLoggedIn") = True Then
       fn_CheckForLogin = 0
   End If

End Function




' Inserts Spaces
'******************************************************
'*******************************************************
Function fn_InsertSpaces(intNumSpaces)
   fn_InsertSpaces = ""
   
   If intNumSpaces = 0 or Not IsNumeric(intNumSpaces) Then
      Exit Function
   End If

   Dim i
   For i = 1 to intNumSpaces
      fn_InsertSpaces = fn_InsertSpaces & "&nbsp;"
   Next

End Function


' Unescape SQL  - Replace any escaped single quotes with an actual single quote
'******************************************************
'*******************************************************
Function fn_UnescapeSQL(strIn)
  
   strIn = Replace(strIn,"''","'")
   fn_UnescapeSQL = strIn

End Function


' Escape SQL  - Replace single quotes with the escape char
'******************************************************
'*******************************************************
Function fn_EscapeSQL(ByVal strIn)
   
   strIn = Replace(strIn,"'","''")
   fn_EscapeSQL = strIn

End Function


' Catch Asp Error and Transfer to ASP Error Page
'******************************************************
'*******************************************************
Sub sub_ErrorCatch()

    If Err.Number <> 0 Then
       'Response.Write Err.Number
       'Response.End
       Server.Transfer PAGE_500100ERROR
    End If
  
End Sub


' Handle a Logic Error and Transfer to the Error Page
'******************************************************
'*******************************************************
Sub sub_HandleLogicError(intErr)

    'Session("sess_intError") = intErr
    Response.Redirect PAGE_ERROR & "?" & QSVAR & "1=" & intErr
  
End Sub

' Set the users agent  (i.e. internet explorer or Netscape)
'************************************************************************************
Sub sub_SetUserAgent()
 
       Dim strUA
       Dim intLoc
       Dim strTmp
       
       strUA = Request.ServerVariables("HTTP_USER_AGENT")
       intLoc = InStr(strUA,"MSIE") 
       If intLoc > 0 Then
          strTmp = USERAGENT_IE
       Else
          intLoc = InStr(strUA,"Netscape") 
          If intLoc > 0 Then
             strTmp = USERAGENT_NN
          Else
             intLoc = InStr(strUA,"Mozilla") 
             If intLoc > 0 Then
                strTmp = USERAGENT_NN
             Else
                strTmp = USERAGENT_OTHER
             End If
          End If
       End If   
       
       Session("sess_strUserAgent") = strTmp

End Sub

Sub Print (strTmp)

     Response.Write strTmp & "<br>"
  
End Sub

Sub PrintStop (strTmp)

     Response.Write strTmp & "<br>"
     Response.End
  
End Sub


' Sub Download Uploaded File
'************************************************************************
Sub sub_DownloadUploadedFile(strFilePathAndName, strDisplayFileName)

    Dim strFNAP, sContentType, sFileType, sNameFile, strChunk, strTmp, i
    Dim objFSO, objTS
    
    'strFN = server.MapPath(strFileName)
    'strFN = Replace(Server.MapPath(PAGE_PORTAL),PAGE_PORTAL,"") & FOLDER_UPLOAD & "\" & strFileName
    strFNAP = strFilePathAndName
    
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set objTS = objFSO.OpenTextFile(strFNAP)
    
    sFileType= Right(strFNAP,4)
    
    'Select Case sFileType
    '      Case ".asf"
    '          sContentType = "video/x-ms-asf"
    '      Case ".avi"
    '          sContentType = "video/avi"
    '      Case ".doc"
    '          sContentType = "application/msword"
    '      Case ".zip"
    '          sContentType = "application/zip"
    '      Case ".css"
    '          sContentType = "text/css"
    '      Case ".pdf"
    '          sContentType = "application/pdf"
    '      Case ".xls"
    '          sContentType = "application/vnd.ms-excel"
    '      Case ".gif"
    '          sContentType = "image/gif"
    '      Case ".jpg", "jpeg"
    '          sContentType = "image/jpeg"
    '      Case ".wav"
    '          sContentType = "audio/wav"
    '      Case ".mp3"
    '          sContentType = "audio/mpeg3"
    '      Case ".mpg", "mpeg"
    '          sContentType = "video/mpeg"
    '      Case ".rtf"
    '          sContentType = "application/rtf"
    '      Case ".htm", "html"
    '          sContentType = "text/html"
    '      Case ".asp", ".asa"
    '          sContentType = "text/asp"
    '      Case Else
    '          sContentType = "application/x-msdownload"
    'End Select
    
    sContentType = "application/x-msdownload"
    
    sNameFile = Right(strFNAP,Len(strFNAP)-InstrRev(strFNAP,"\"))
    
    Response.ContentType = sContentType
    'Response.AddHeader "Content-Disposition", "attachment;filename=" & sNameFile
    Response.AddHeader "Content-Disposition", "attachment;filename=" & strDisplayFileName
    
    Response.Buffer = True
    Do While Not objTS.AtEndOfStream
      strChunk = objTS.Read(32)
      strTmp = ""
      For i = 1 to Len(strChunk)
           strTmp = strTmp & ChrB(Asc(Mid(strChunk, i, 1)))
      Next
      Response.BinaryWrite strTmp
      Response.Flush
    Loop
    objTS.Close
    Set objTS = Nothing
    Set objFSO = Nothing 

End Sub   'sub_DownloadUploadedFile

%>





















