VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTMLClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ReportId As Long
Private blnTemplateRead As Boolean
Private LabNo As String
Private Specialty As String
Private HospNo As String
Private pForename As String
Private pSurname As String
Private Clinician As String
Private Destination As String
Private Collected As String
Private Received As String
Private Reported As String
Private SampText As String
Private htmlHeader As String
Private htmlTrailer As String

Private Sub Class_Initialize()
   ReadTemplate
End Sub

Private Property Get Comments() As String
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim objBuff As StringBuffer
    
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Comments"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   RS.Filter = "Comment_Type = 'D'"
   If RS.RecordCount > 0 Then
      Set objBuff = New StringBuffer
      objBuff.Append "<BR><STRONG>Report Comments:</STRONG><BR>"
      Do Until RS.EOF
          objBuff.Append RS!Service_Report_Comment
          objBuff.Append "<BR>"
          RS.MoveNext
      Loop
      Comments = objBuff.Value
      Set objBuff = Nothing
    End If
    
    RS.Close
    Set objBuff = Nothing
    Set RS = Nothing
    Set iceCmd = Nothing
End Property

Public Function CreateMessage(RepId As Long) As String
   On Error GoTo procEH
   ReportId = RepId
   CreateMessage = ReportBody(RepId)
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.CreateMessage"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Property Get FileHeader() As String
   FileHeader = htmlHeader
End Property

Public Property Get FileTrailer() As String
   FileTrailer = htmlTrailer
End Property

Public Property Get InvalidChars() As Variant
   ReDim strArray(2)
   strArray(0) = ""
   strArray(1) = ""
   InvalidChars = strArray
End Property

Public Property Get ReleaseChar() As String
   ReleaseChar = ""
End Property

Public Property Get RetainGrave() As Boolean
   RetainGrave = False
End Property

Private Sub HealthParties()
   Dim hpData As New clsHealthParties
   
   hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
   hpData.LTSIndex = repListRS!EDI_LTS_Index
   hpData.Read repListRS!EDI_Report_Index, True
   
   Clinician = hpData.HP902Name
   
   Destination = repListRS!EDI_Loc_Nat_Code_To
   Set hpData = Nothing
End Sub

Private Property Get Investigation() As String
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim m As Integer
   Dim intResultCount As Integer
   Dim objBuff As StringBuffer
   Dim blnFirst As Boolean
   
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Sample"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   Collected = RS!Collection_DateTime & ""
   Received = RS!Collection_DateTimeReceived & ""
   Reported = RS!DateTime_Of_Report
   SampText = RS!Sample_ID & " (" & RS!Sample_Text & ")"
   
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Investigation"
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , RS!Sample_Index)
      .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
      RS.Close
      Set RS = .Execute
      .Parameters.Delete ("Abnormal")
      .Parameters.Delete ("SampId")
      .Parameters.Delete ("LTSIndex")
   End With
   blnFirst = True
    
   Set objBuff = New StringBuffer
   objBuff.Append "<TABLE border=0 style=""WIDTH: 90%"" align=""center"" id=""tblResults"">"
   objBuff.Append "<TR>"
   objBuff.Append "<TD width=""51%""></TD>"
   objBuff.Append "<TD width=""4%""></TD>"
   objBuff.Append "<TD width=""15%"">"
   
   If blnFirst Then objBuff.Append "<STRONG>Result</STRONG>"
   objBuff.Append "</TD>"
   objBuff.Append "<TD width=""15%"">"
   If blnFirst Then objBuff.Append "<STRONG>Units</STRONG>"
   objBuff.Append "</TD>"
   objBuff.Append "<TD width=""15%"">"
   If blnFirst Then objBuff.Append "<STRONG>Range</STRONG>"
   objBuff.Append "</TD></TR>"
   blnFirst = False
                
   Do Until RS.EOF
      
      objBuff.Append "<TR><TD><STRONG>"
      objBuff.Append Trim(RS!Investigation_Requested)
      objBuff.Append "</STRONG></TD></TR>"
      
      If RS!Comment_Marker Then
         With iceCmd
            .CommandText = "ICELABCOMM_Report_InvestComments"
            .Parameters(0).Value = RS!Investigation_Index
            Set RS2 = .Execute
         End With
         Do Until RS2.EOF
            objBuff.Append "<TR><TD Colspan=5 STYLE=""padding-left:10px"">"
            objBuff.Append RTrim(RS2!Service_Investigation_Comment)
            objBuff.Append "</TD></TR>"
            RS2.MoveNext
         Loop
         RS2.Close
         Set RS2 = Nothing
      End If
      
      With iceCmd
'         .ActiveConnection = ICEcon
'         .CommandType = adCmdStoredProc
         .CommandText = "ICELABCOMM_Report_Results"
         .Parameters(0).Value = RS!Investigation_Index
         .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
'         .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RS!Investigation_Index)
         Set RS2 = .Execute
         .Parameters.Delete ("LTSIndex")
      End With
      
      Do Until RS2.EOF
         objBuff.Append "<TR style=""background-color:gainsboro""><TD STYLE=""padding-left:10px"">"
         objBuff.Append RS2!Result_Rubric
         objBuff.Append "</TD><TD align=""center"">"
         If RS2!Abnormal_Flag Then
             objBuff.Append "<FONT COLOR='RED'>*</FONT>"
         End If
         
         objBuff.Append "</TD><TD>"
         
         If RS2!Abnormal_Flag Then
             objBuff.Append "<FONT COLOR='RED'>"
             objBuff.Append RS2!Result
             objBuff.Append "</FONT>"
         Else
             objBuff.Append RS2!Result
         End If
         
         objBuff.Append "</TD><TD>"
         objBuff.Append RS2!UOM_Code
         objBuff.Append "</TD><TD>"
         objBuff.Append RS2!Lower_Range & " to " & RS2!Upper_Range
         objBuff.Append "</TD></TR>"
          ' Get comments
'         Set iceCmd = Nothing
         If RS2!Comment_Marker Then
            With iceCmd
               .CommandText = "ICELABCOMM_Report_ResultComment"
               .Parameters(0).Value = RS2!Result_Index
               Set RS3 = .Execute
            End With
            
            Do Until RS3.EOF
               objBuff.Append "<TR><TD Colspan=5 STYLE=""Font-Family:Courier"">"
               objBuff.Append RS3!Service_Result_Comment
               objBuff.Append "</TD></TR>"
               RS3.MoveNext
            Loop
            RS3.Close
         End If
         RS2.MoveNext
      Loop
      
      RS2.Close
      RS.MoveNext
   Loop
   Investigation = objBuff.Value
   
   RS.Close
   
   Set RS3 = Nothing
   Set RS2 = Nothing
   Set RS = Nothing
   Set iceCmd = Nothing
   Set objBuff = Nothing
   Exit Property
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.Investigation"
   eClass.Add Err.Number, Err.Description, Err.Source
End Property

Public Property Get Notification() As Boolean
   Notification = False
End Property

Private Property Get Patient() As String
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim objBuff As New StringBuffer
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Patient"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   LabNo = RS!Service_Report_Id
   Specialty = RS!Service_Report_Type
   HospNo = RS!Hospital_Number
   pForename = RS!Forename
   pSurname = RS!Surname
   
   With objBuff
      .Append "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=2>"
      .Append "<TR><TD width=""18%""><NOBR><STRONG>Patient Name:</STRONG></NOBR></TD><TD width=""33%""><NOBR>"
      .Append RS!Forename
      .Append " "
      .Append RS!Surname
      .Append "</NOBR></TD><TD width=""22%""><STRONG><NOBR>NHS Number:</NOBR></STRONG></TD><TD width=""27%""><NOBR>"
      .Append IIf(Trim(RS!New_NHS_No & "") = "", "Not Known", RS!New_NHS_No)
      .Append "</NOBR></TD></TR><TR><TD><STRONG><NOBR>Date of Birth:</NOBR></STRONG></TD><TD><NOBR>"
      .Append RS!Date_Of_Birth
      .Append "</NOBR></TD><TD><STRONG><NOBR>Hospital Number:</NOBR></STRONG></TD><TD><NOBR>"
      .Append IIf(Trim(RS!Hospital_Number & "") = "", "Not Known", RS!Hospital_Number)
      .Append "</NOBR></TD></TR><TR><TD><STRONG>Address:</STRONG></TD><TD colspan=3><NOBR>"
      .Append RS!Pat_Addr_Line1
      If RS!Pat_Addr_Line2 <> "" Then
          .Append ", "
          .Append RS!Pat_Addr_Line2
      End If
      If RS!Pat_Addr_Line3 <> "" Then
          .Append ", "
          .Append RS!Pat_Addr_Line3
      End If
      If RS!Pat_Addr_Line4 <> "" Then
          .Append ", "
          .Append RS!Pat_Addr_Line4
      End If
      If RS!Pat_Addr_Line5 <> "" Then
          .Append ", "
          .Append RS!Pat_Addr_Line5
      End If
      If RS!Pat_PostCode <> "" Then
          .Append ", "
          .Append RS!Pat_PostCode
      End If
      .Append "</NOBR></TD></TR></TABLE>"
      
      Patient = .Value
   End With
End Property

Public Sub ReadTemplate()
   On Error GoTo procEH
   Dim i As Integer
   Dim oText As TextStream
   Dim oFile As FileSystemObject
'   Dim oWebReport As WebReport
   Dim sBuffer As New StringBuffer
   
   'Open template to a string
   
   htmlHeader = ""
   htmlBody = ""
   htmlTrailer = ""
   
   eClass.FurtherInfo = "Reading Report template (" & App.Path & "\report.htm)"
   
   Set oText = fs.OpenTextFile(App.Path & "\" & "Report.htm", ForReading, False)
   For i = 0 To 7
      htmlHeader = htmlHeader & oText.ReadLine & vbCrLf
   Next i
   
   For i = 0 To 39
      htmlBody = htmlBody & oText.ReadLine & vbCrLf
   Next i
   
   Do Until oText.AtEndOfStream
      htmlTrailer = htmlTrailer & oText.ReadLine & vbCrLf
   Loop
   
   oText.Close
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.ReadTemplate"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function ReportBody(RepId As Long) As String
   On Error GoTo procEH
   Dim strReport As String
   
   strReport = htmlBody
   HealthParties
'  Start replacing tags like this <WC@XXXXX></WC@XXXXX>
   strReport = Replace$(strReport, "<WC@PATIENT></WC@PATIENT>", Patient)
   strReport = Replace$(strReport, "<WC@LABNO></WC@LABNO>", LabNo)
   strReport = Replace$(strReport, "<WC@SPECIALTY></WC@SPECIALTY>", Specialty)
   strReport = Replace$(strReport, "<WC@LOCATION></WC@LOCATION>", Destination)
   strReport = Replace$(strReport, "<WC@CLINICIAN></WC@CLINICIAN>", Clinician)
   strReport = Replace$(strReport, "<WC@COMMENTS></WC@COMMENTS>", Comments)
   strReport = Replace$(strReport, "<WC@INVESTIGATIONS></WC@INVESTIGATIONS>", Investigation)
   strReport = Replace$(strReport, "<WC@SAMPLE></WC@SAMPLE>", SampText)
   strReport = Replace$(strReport, "<WC@COLLECTED></WC@COLLECTED>", Collected)
   strReport = Replace$(strReport, "<WC@RECEIVED></WC@RECEIVED>", Received)
   strReport = Replace$(strReport, "<WC@REPORTED></WC@REPORTED>", Reported)
   
   ReportBody = strReport
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.ReportBody"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function
'
'Public Sub Prepare_HTML_File(RepNo As Long)
'   On Error GoTo procEH
'   Dim i As Integer
'   Dim strReport As String
'   Dim oText As TextStream
'   Dim oFile As FileSystemObject
'   Dim oWebReport As WebReport
'   Dim sBuffer As New StringBuffer
'
'   'Open template to a string
'
'   eClass.FurtherInfo = "Reading Report template (" & App.Path & "\report.htm)"
'   If htmlHeader = "" Then
'      Set oText = fs.OpenTextFile(App.Path & "\" & "Report.htm", ForReading, False)
'      For i = 0 To 7
'         htmlHeader = htmlHeader & oText.ReadLine & vbCrLf
'      Next i
'      For i = 0 To 31
'         htmlBody = htmlBody & oText.ReadLine & vbCrLf
'      Next i
'      Do Until oText.AtEndOfStream
'         htmlTrailer = htmlTrailer & oText.ReadLine & vbCrLf
'      Loop
'      oText.Close
'   End If
'
'   strReport = htmlBody
'
'   eClass.FurtherInfo = "Access web reporting class"
'   Set oWebReport = New WebReport
'   'Dynamicaly pass in the report number here(ignore that it says get by patient)
'   oWebReport.GetByPatient RepNo
'
'   ' start replacing tags like this <WC@XXXXX></WC@XXXXX>
'   strReport = Replace$(strReport, "<WC@PATIENT></WC@PATIENT>", oWebReport.Patient)
'   strReport = Replace$(strReport, "<WC@LABNO></WC@LABNO>", oWebReport.LabNo)
'   strReport = Replace$(strReport, "<WC@SPECIALTY></WC@SPECIALTY>", oWebReport.Specialty)
'   strReport = Replace$(strReport, "<WC@LOCATION></WC@LOCATION>", oWebReport.Destination)
'   strReport = Replace$(strReport, "<WC@SAMPLE></WC@SAMPLE>", "")
'   strReport = Replace$(strReport, "<WC@CLINICIAN></WC@CLINICIAN>", oWebReport.Clinician)
'   strReport = Replace$(strReport, "<WC@COMMENTS></WC@COMMENTS>", oWebReport.Comments)
'   strReport = Replace$(strReport, "<WC@INVESTIGATIONS></WC@INVESTIGATIONS>", oWebReport.Investigations)
'   strReport = Replace$(strReport, "<WC@COLLECTED></WC@COLLECTED>", oWebReport.Collected)
'   'strReport = Replace$(strReport, "<WC@RECEIVED></WC@RECEIVED>", oWebReport.Received)
'   strReport = Replace$(strReport, "<WC@REPORTED></WC@REPORTED>", oWebReport.Reported)
'   htmlBuf = htmlBuf & strReport
'   Exit Sub
'
'procEH:
'   If eClass.Behaviour = -1 Then
'      Stop
'      Resume
'   End If
'   eClass.CurrentProcedure = "CommonCode.Prepare_HTML_File"
'   eClass.Add Err.Number, Err.Description, Err.Source
'End Sub

Public Sub SendAsHTML(Mailaddress As String, _
                      HtmlFile As String)
   On Error GoTo procEH
   Dim objApp As Outlook.Application
   Dim objNameSpace As Outlook.NameSpace
   Dim objMailItem As Outlook.MailItem

'  Send using Outlook 2000
   
   eClass.FurtherInfo = "Preparing outlook"
   
   Set objApp = New Outlook.Application
   Set objNameSpace = objApp.GetNamespace("MAPI")
   Set objMailItem = objApp.CreateItem(olMailItem)
   
   objMailItem.Recipients.Add Mailaddress
   
   objMailItem.Subject = "HTML Reports for " & Mailaddress
   
   objMailItem.htmlBody = HtmlFile
   objMailItem.Send
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.SendAsHTML"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SendHTMLCopy(ReportIndex As Long, _
                        Mailaddress As String)
   On Error GoTo procEH
   Dim strHTML As String
   Dim objApp As New Outlook.Application
   Dim objNameSpace As Outlook.NameSpace
   Dim objMailItem As Outlook.MailItem
   
   ReportId = ReportIndex
   ReadTemplate
   strHTML = htmlHeader & ReportBody(ReportId) & htmlTrailer
   
'  Send using Outlook 2000
   eClass.FurtherInfo = "Preparing to send via Outlook"
   
   Set objNameSpace = objApp.GetNamespace("MAPI")
   Set objMailItem = objApp.CreateItem(olMailItem)
   objMailItem.Recipients.Add Mailaddress
   
   objMailItem.Subject = "Copy Report - Requesting Clinician: " & Clinician & " (" & _
                         msgControl.CurrentPractice & "/" & repListRS!EDI_Loc_Nat_Code_To & _
                         " Trust: " & msgControl.Organisation & "/" & msgControl.Trust & ")"
   
   msgControl.LogReportMessage -1, "HTML copy of report " & LabNo & " for " & pForename & " " & pSurname & _
                                   " sent to: " & Mailaddress
   objMailItem.htmlBody = strHTML
   objMailItem.Send
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.SendHTMLCopy"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
