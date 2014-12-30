Attribute VB_Name = "HTMLMod"
Option Explicit


'This routine sends individual e-Mails with report as an HTML bodypart
'It requires following references added to project.
'Webreporting.dll
'Microsoft Outlook 9.0 Object Library
'Microsoft sCRIPTING RUNTIME
'It also requires the template Report.HTM to be in the app.path

Public Sub Send_HTML_E_Mail(RepNo As Long, _
                            Mailaddress As String, _
                            HistoryFile As String)
   On Error GoTo procEH
   Dim strReport As String
   Dim oText As TextStream
   Dim oFile As FileSystemObject
   Dim oWebReport As WebReport
   Dim objApp As Outlook.Application
   Dim objNameSpace As Outlook.NameSpace
   Dim objMailItem As Outlook.MailItem
   
   eClass.FurtherInfo = "Reading Report template (" & App.Path & "\Report.htm)"
   If Mailaddress = "" Or Mailaddress = Null Then Exit Sub
   'Open template to a string
   Set oFile = New FileSystemObject
   Set oText = oFile.OpenTextFile(App.Path & "\" & "Report.htm", ForReading, False)
   strReport = oText.ReadAll

   eClass.FurtherInfo = "Creating Wbe Reporting class"
   Set oWebReport = New WebReport
   'Dynamicaly pass in the report number here(ignore that it says get by patient)
   oWebReport.GetByPatient RepNo

   ' start replacing tags like this <WC@XXXXX></WC@XXXXX>
   strReport = Replace$(strReport, "<WC@PATIENT></WC@PATIENT>", oWebReport.Patient)
   strReport = Replace$(strReport, "<WC@LABNO></WC@LABNO>", oWebReport.LabNo)
   strReport = Replace$(strReport, "<WC@SPECIALTY></WC@SPECIALTY>", oWebReport.Specialty)
   strReport = Replace$(strReport, "<WC@LOCATION></WC@LOCATION>", oWebReport.Destination)
   strReport = Replace$(strReport, "<WC@SAMPLE></WC@SAMPLE>", "")
   strReport = Replace$(strReport, "<WC@CLINICIAN></WC@CLINICIAN>", oWebReport.Clinician)
   strReport = Replace$(strReport, "<WC@COMMENTS></WC@COMMENTS>", oWebReport.Comments)
   strReport = Replace$(strReport, "<WC@INVESTIGATIONS></WC@INVESTIGATIONS>", oWebReport.Investigations)
   strReport = Replace$(strReport, "<WC@COLLECTED></WC@COLLECTED>", oWebReport.Collected)
   'strReport = Replace$(strReport, "<WC@RECEIVED></WC@RECEIVED>", oWebReport.Received)
   strReport = Replace$(strReport, "<WC@REPORTED></WC@REPORTED>", oWebReport.Reported)

   ' Send using Outlook 2000
   eClass.FurtherInfo = "Preparing to send via Outlook"
   Set objApp = New Outlook.Application
   Set objNameSpace = objApp.GetNamespace("MAPI")
   Set objMailItem = objApp.CreateItem(olMailItem)
   objMailItem.Recipients.Add Mailaddress
   ahReport = ahReps.GetReportByID(RepNo)
   With ahReport.Patient
      objMailItem.Subject = "Report For " & oWebReport.Clinician & " - " & .HospNo & " " & .Forename & "," _
                              & .Surname
      Open HistoryFile For Append As #1
      Print #1, "HTML copy of report " & oWebReport.LabNo & " for " & .Forename & " " & .Surname & _
                " sent to: " & oWebReport.Clinician & " (" & Mailaddress & ")"
      Close #1
   End With
   objMailItem.htmlBody = strReport
   objMailItem.Send
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "CommonCode.Send_HTML_Email"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Prepare_HTML_File(RepNo As Long)
   On Error GoTo procEH
   Dim i As Integer
   Dim strReport As String
   Dim oText As TextStream
   Dim oFile As FileSystemObject
   Dim oWebReport As WebReport
   Dim sBuffer As New StringBuffer
   
   'Open template to a string
   
   eClass.FurtherInfo = "Reading Report template (" & App.Path & "\report.htm)"
   If htmlHeader = "" Then
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
   End If

   strReport = htmlBody
   
   eClass.FurtherInfo = "Access web reporting class"
   Set oWebReport = New WebReport
   'Dynamicaly pass in the report number here(ignore that it says get by patient)
   oWebReport.GetByPatient RepNo

   ' start replacing tags like this <WC@XXXXX></WC@XXXXX>
   strReport = Replace$(strReport, "<WC@PATIENT></WC@PATIENT>", oWebReport.Patient)
   strReport = Replace$(strReport, "<WC@LABNO></WC@LABNO>", oWebReport.LabNo)
   strReport = Replace$(strReport, "<WC@SPECIALTY></WC@SPECIALTY>", oWebReport.Specialty)
   strReport = Replace$(strReport, "<WC@LOCATION></WC@LOCATION>", oWebReport.Destination)
   strReport = Replace$(strReport, "<WC@SAMPLE></WC@SAMPLE>", "")
   strReport = Replace$(strReport, "<WC@CLINICIAN></WC@CLINICIAN>", oWebReport.Clinician)
   strReport = Replace$(strReport, "<WC@COMMENTS></WC@COMMENTS>", oWebReport.Comments)
   strReport = Replace$(strReport, "<WC@INVESTIGATIONS></WC@INVESTIGATIONS>", oWebReport.Investigations)
   strReport = Replace$(strReport, "<WC@COLLECTED></WC@COLLECTED>", oWebReport.Collected)
   strReport = Replace$(strReport, "<WC@RECEIVED></WC@RECEIVED>", oWebReport.Received)
   strReport = Replace$(strReport, "<WC@REPORTED></WC@REPORTED>", oWebReport.Reported)
   htmlBuf = htmlBuf & strReport
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "CommonCode.Prepare_HTML_File"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SendAsHTML(Mailaddress As String, _
                      HistoryFile As String)
   On Error GoTo procEH
   Dim objApp As Outlook.Application
   Dim objNameSpace As Outlook.NameSpace
   Dim objMailItem As Outlook.MailItem

'  Send using Outlook 2000
   htmlBuf = htmlHeader & htmlBuf & htmlTrailer
   
   eClass.FurtherInfo = "Preparing outlook"
   Set objApp = New Outlook.Application
   Set objNameSpace = objApp.GetNamespace("MAPI")
   Set objMailItem = objApp.CreateItem(olMailItem)
   objMailItem.Recipients.Add Mailaddress
'   ahReport = ahReps.GetReportByID(RepNo)
   With ahReport.Patient
      objMailItem.Subject = "HTML Reports for " & Mailaddress
   End With
   
   eClass.FurtherInfo = "Creating history file (" & HistoryFile & ")"
   Open HistoryFile For Output As #1
   Print #1, htmlBuf
   Close #1
   objMailItem.htmlBody = htmlBuf
   objMailItem.Send
   htmlHeader = ""
   htmlBuf = ""
   htmlTrailer = ""
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "CommonCode.SendAsHTMLFile"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub


