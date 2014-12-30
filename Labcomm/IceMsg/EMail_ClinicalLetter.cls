VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMail_ClinicalLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
'
'   Project Name : ICEMsg
'       Identification : ICEMsg.vbp
'   Class Name : EMail_ClinicalLetter
'       Identification : EMail_ClinicalLetter.cls
'
'   Copyright (c) 2003 - Anglia Healthcare Systems Ltd
'
'   Author :
'       Simon Jones, 17th November 2003
'
'
'                                    www.ahsl.com
' *************************************************************************

Option Explicit

Private oDataRS As ADODB.Recordset

Private SMTPAddress As String

Public Enum enmEMailType
    FullLetter = 0
'    Notify = 1
End Enum
Private EType As enmEMailType

Private Type EMailDetails
    SMTPAddress As String
    Subject As String
    Body As String
End Type
Private EMail As EMailDetails

Public Function CreateMessage(ByVal ReportId As Long) As Variant

    Dim iceCmd As New ADODB.Command
    Dim oRecordset As New ADODB.Recordset
    Dim blnDoingTrans As Boolean
    Dim LetterType As String

    On Error GoTo procEH
    msgControl.ReportStatus = 0

    eClass.FurtherInfo = "Get Letter Type for E-Mail Message"
    With iceCmd
        .ActiveConnection = iceCon
        .CommandType = adCmdStoredProc
        .CommandText = "ICEMSG_Get_Letter_Type"
        .Parameters.Append .CreateParameter("ServiceReportID", adInteger, adParamInput, , ReportId)
        Set oRecordset = .Execute
        blnDoingTrans = True
    End With

    If oRecordset.BOF And oRecordset.EOF Then
        LetterType = ""
        msgControl.ReportStatus = RS_DATA_INTEGRITY
        msgControl.LogReportMessage RS_DATA_INTEGRITY, "No details available for Report ID provided"
    Else
        LetterType = oRecordset!LetterType
    End If
    oRecordset.Close
    blnDoingTrans = False
    Set oRecordset = Nothing
    Set iceCmd = Nothing
    
    'frmMain.NewFileRequired
        
    If LetterType <> "" Then
        Select Case LetterType
            Case "DISCHARGE", "DOR", "DTH", "TCT"
                eClass.FurtherInfo = "Build Discharge Letter E-Mail"
                CreateMessage = EMail_Discharge(ReportId)
            Case Else
                msgControl.ReportStatus = RS_GENERAL
                msgControl.LogReportMessage RS_GENERAL, "E-Mail generation of the Letter Type (" & oDataRS!LetterType & ") not enabled"
        End Select
    End If

    If msgControl.ReportStatus > 0 Then
        CreateMessage = "Errors encountered when generating message 'See comments for description'"
    End If

    Exit Function

procEH:
    If eClass.Behaviour = -1 Then
        Stop
        Resume
    End If
    eClass.CurrentProcedure = "EMail_ClinicalLetter.CreateMessage"
    eClass.Add Err.Number, Err.Description, Err.Source
    If blnDoingTrans Then oRecordset.Close
    Set oRecordset = Nothing
    Set iceCmd = Nothing

End Function

Public Property Get FileHeader() As String

    FileHeader = ""

End Property

Public Property Get FileTrailer() As String

    FileTrailer = ""

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

Public Property Get EMail_Type() As enmEMailType

    EMail_Type = EType

End Property

Public Property Let EMail_Type(ByVal vNewValue As enmEMailType)

    EType = vNewValue

End Property

Private Function EMail_MainData(ByVal ServiceReportID As Long) As Boolean

    Dim iceCmd As New ADODB.Command
    Dim RSFile As String
    Dim blnDoingTrans As Boolean

    On Error GoTo procEH

    With iceCmd
        .ActiveConnection = iceCon
        .CommandType = adCmdStoredProc
        .CommandText = "ICEMSG_Get_Discharge_Letter_Details"
        .Parameters.Append .CreateParameter("ServiceReportID", adInteger, adParamInput, , ServiceReportID)
        Set oDataRS = .Execute
        blnDoingTrans = True
    End With

    If oDataRS.BOF And oDataRS.EOF Then
        msgControl.ReportStatus = RS_DATA_INTEGRITY
        msgControl.LogReportMessage RS_DATA_INTEGRITY, "No details available for Report ID provided"
        EMail_MainData = False
    Else
        RSFile = fs.BuildPath(App.Path, "oDataRS,ADTG")
        If fs.FileExists(RSFile) Then
            fs.DeleteFile RSFile
        End If

        Set oDataRS.ActiveConnection = Nothing
        oDataRS.Save RSFile, adPersistADTG
        oDataRS.Close
        blnDoingTrans = False
        oDataRS.Open RSFile

        EMail_MainData = True
    End If

    Set iceCmd = Nothing
    Exit Function
    
procEH:
    If eClass.Behaviour = -1 Then
        Stop
        Resume
    End If
    eClass.CurrentProcedure = "EMail_ClinicalLetter.EMail_MainData"
    eClass.Add Err.Number, Err.Description, Err.Source
    EMail_MainData = False

End Function

Private Sub Class_Initialize()

    Set oDataRS = New ADODB.Recordset

   sDOMAIN = Read_Ini_Var("LETTER", "DOMAIN", INIFile)
   
   defUserIndex = Read_Ini_Var("LETTER", "UserIndex", INIFile)
   If defUserIndex = "" Or IsNumeric(defUserIndex) = False Then
      defUserIndex = -1
   End If

End Sub

Private Sub Class_Terminate()
    
    If oDataRS.State <> 0 Then oDataRS.Close
    Set oDataRS = Nothing

End Sub

Private Function EMail_Discharge(ByVal ReportId As Long) As String
    
    Dim objViewDischarge As Object
    Dim strHyperlinkID As String
    Dim blnGoodSMTP As Boolean
    Dim Success As Boolean
    Dim strString As String
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim ErrorCode As String
    Dim ErrorDescription As String

    On Error GoTo procEH

    eClass.FurtherInfo = "Get Data for XML Message"
    If Not EMail_MainData(ReportId) Then Exit Function

    ' Remove the curly brackets from the HyperlinkID
    If DBVERSION < 330 Or DBVERSION > 4999 Then strHyperlinkID = Mid(oDataRS!HyperlinkID, 2, Len(oDataRS!HyperlinkID) - 2)

    If repListRS!EDI_SMTP_Active And repListRS!GP_SMTP <> "" Then
        EMail.SMTPAddress = repListRS!GP_SMTP
        blnGoodSMTP = True
    ElseIf repListRS!EDI_SMTP_Mail <> "" Then
        EMail.SMTPAddress = repListRS!EDI_SMTP_Mail
        blnGoodSMTP = True
    Else
        blnGoodSMTP = False
    End If

    If blnGoodSMTP Then
        eClass.FurtherInfo = "Build Subject for E-Mail"
        EMail.Subject = EMail_BuildSubject

        Select Case EType
            Case FullLetter
                If DBVERSION > 229 And DBVERSION < 5000 Then
                    eClass.FurtherInfo = "Call Clinical Letter Web Service"
                    strString = GetLetter(oDataRS!LetterID, oDataRS!HospitalNumber)
                    If Left(strString, 5) = "ERROR" Then
                        ' This means an error occured obtaining the letter.
                        StartPos = 8
                        EndPos = InStr(1, strString, Chr(13) & Chr(10))
                        ErrorDescription = Mid(strString, 8, EndPos - 8)
                        msgControl.ReportStatus = MS_NO_OUTPUT
                        msgControl.MessageStatus = MS_NO_OUTPUT
                        msgControl.LogReportMessage MS_NO_OUTPUT, "Error in Clinical Letter Web Service: " & ErrorDescription
                        Exit Function
                    End If
                    strString = Left(strString, Len(strString) - 14) & HTMLFooting
                Else
                    ' Now use the ICEDesktopDischarge.Letter class to obtain HTML of Clinical Letter
                    eClass.FurtherInfo = "Call ICDesktopDischarge.Letter Class"
                    Set objViewDischarge = CreateObject("ICEDesktopDischarge.Letter")
                        objViewDischarge.dbConnection = "FILE NAME=" & App.Path & "\ice.udl"
                        strString = objViewDischarge.View(oDataRS!LetterID & "," & strHyperlinkID, "", False)
                    Set objViewDischarge = Nothing
                    StartPos = InStr(1, strString, "<!--SERRCODE:")
                    If StartPos > 0 Then
                        ' This means an error occured obtaining the letter.
                        ErrorCode = Mid(strString, StartPos + 16, 4)
                        StartPos = InStr(1, strString, "<!--SERRDESC:")
                        EndPos = InStr(1, strString, ":EERRDESC-->")
                        ErrorDescription = Mid(strString, StartPos + 13, EndPos - StartPos - 13)
                        msgControl.ReportStatus = MS_NO_OUTPUT
                        msgControl.MessageStatus = MS_NO_OUTPUT
                        msgControl.LogReportMessage MS_NO_OUTPUT, "Error in HTML Letter Class Code:" & ErrorCode & " Description: " & ErrorDescription
                        Exit Function
                    End If
                    strString = HTMLHeading & strString & HTMLFooting
                End If
'            Case Notify
'                strString = EMail_BuildBody(strHyperlinkID)
        End Select
        EMail.Body = strString
        Success = EMail_UpdateDatabase
        If Success Then
            ' Finished building E-Mail message
            EMail_Discharge = "E-Mail Message complete and database updated ready for ICE...Mail to process"
        End If
        
    Else
        ' No SMTP Address to send e-mail to!
        msgControl.ReportStatus = MS_NO_OUTPUT
        msgControl.MessageStatus = MS_NO_OUTPUT
        msgControl.LogReportMessage MS_NO_OUTPUT, "No SMTP Address for Recipient (Person or Practice)"
    End If

    Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "EMail_ClinicalLetter.EMail_Discharge"
   eClass.Add Err.Number, Err.Description, Err.Source

End Function

Private Function EMail_BuildSubject() As String

    Dim sSubject As String

    sSubject = ""
    sSubject = sSubject & repListRS!EDI_Loc_Nat_Code_To & " - "
    If (Not IsNull(repListRS!GP_Code)) And repListRS!GP_Code <> "" Then sSubject = sSubject & repListRS!GP_Code & " - "
    Select Case oDataRS!SpellVisitType
        Case "O", "OP"
            sSubject = sSubject & "Outpatient Letter"
        Case "OF"
            sSubject = sSubject & "Outpatient Follow-Up Letter"
        Case "AE", "A"
            sSubject = sSubject & "Accident & Emergency Letter"
        Case "TC"
            sSubject = sSubject & "Cardiology Clinical Letter"
        Case "I", "IP"
            sSubject = sSubject & "Inpatient Letter"
        Case Else
            sSubject = sSubject & "Clinical Letter"
    End Select
    sSubject = sSubject & " - " & Format(oDataRS!SpellEndDate, "dd/mm/yyyy") & " - "
    If ((Not IsNull(oDataRS!Forename)) And oDataRS!Forename <> "") And ((Not IsNull(oDataRS!Surname)) And oDataRS!Surname <> "") Then sSubject = sSubject & Trim(oDataRS!Forename & " " & oDataRS!Surname)
    If (Not IsNull(oDataRS!DateOfBirth)) And oDataRS!DateOfBirth <> "" Then sSubject = sSubject & ", " & oDataRS!DateOfBirth
    If (Not IsNull(oDataRS!NHSNumber)) And oDataRS!NHSNumber <> "" Then sSubject = sSubject & ", " & oDataRS!NHSNumber
    If (Not IsNull(repListRS!EDI_OP_Name)) And repListRS!EDI_OP_Name <> "" Then sSubject = sSubject & " - " & repListRS!EDI_OP_Name
    If (Not IsNull(oDataRS!SpellHospital)) And oDataRS!SpellHospital <> "" Then sSubject = sSubject & " - " & oDataRS!SpellHospital

    EMail_BuildSubject = sSubject

End Function

Private Function EMail_BuildBody(sHyperlinkID As String) As String

    Dim sBody As String
    Dim sHyperlink As String

    sHyperlink = repListRS!WebAddress & "icedesktop/viewdischargeletter/viewletter.asp?hletterid=" & oDataRS!LetterID & "," & sHyperlinkID
    
    sBody = ""
    sBody = sBody & "<font style='font-family: arial; font-size: 12px'>"
    sBody = sBody & "<h4>Discharge Summary Notification</h4>"
    sBody = sBody & "<strong>Patient</strong>: " & oDataRS!Surname & ", " & oDataRS!Forename
    sBody = sBody & "<br/><strong>Date Of Birth</strong>: " & oDataRS!DateOfBirth
    sBody = sBody & "<br/><strong>NHS Number</strong>: " & oDataRS!NHSNumber
    sBody = sBody & "<br/><strong>Hospital Number</strong>: " & oDataRS!HospitalNumber
    sBody = sBody & "<p>Discharged from <strong>" & oDataRS!SpellHospital & "</strong> on <strong>" & oDataRS!SpellEndDate & "</strong></p>"
    sBody = sBody & "Click <a href='" & sHyperlink & "'><strong>HERE</strong></a> to view Discharge Summary"
    sBody = sBody & "<br/><br/>"
    sBody = sBody & "<div align=center style='background : #cccccc'><strong>Please Note</strong>"
    sBody = sBody & "<br/>Any comments should be sent to "
    sBody = sBody & oDataRS!SpellHospital ' "Norfolk & Norwich University Hospital NHS Trust"
    sBody = sBody & " as any e-mails sent to this address will not be reviewed.</div>"
    sBody = sBody & "</font>"

    EMail_BuildBody = sBody

End Function

Private Function EMail_UpdateDatabase() As Boolean

    Dim strSQL As String

    On Error GoTo procEH

    strSQL = "INSERT INTO Email_List (SMTP_Address, Subject, Body, Date_Added) VALUES ("
    strSQL = strSQL & "'" & EMail.SMTPAddress & "', '" & Replace(EMail.Subject, "'", "''") & "', '" & Replace(EMail.Body, "'", "''") & "', '" & Format(Date, "dd mmm yyyy") & "')"
    iceCon.Execute strSQL

    EMail_UpdateDatabase = True

    Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "EMail_ClinicalLetter.EMail_UpdateDatabase"
   eClass.Add Err.Number, Err.Description, Err.Source
   EMail_UpdateDatabase = False

End Function

Private Property Get HTMLHeading() As String

    HTMLHeading = "<style>body {background-color: #ffffff; font-size: 8pt; font-family: tahoma, verdana, arial;} " & _
                            "input, select, textarea, file, submit, button, span, div, table {font-size: 8pt; font-family: tahoma, verdana;}</style>" & _
                            "<table height='100%' width='100%' cellpadding='0' cellspacing='0'><tr><td height='100%' valign='top' align='center'>"

End Property

Private Property Get HTMLFooting() As String

    If DBVERSION > 229 And DBVERSION < 5000 Then
        HTMLFooting = "<table style='width: 100%;' cellspacing='0'><tr><td>&nbsp;</td></tr><tr><td>" & _
            "<div align='center' style='background : #cccccc'><strong>Please Note</strong><br/>Any comments should be sent to " & _
            oDataRS!SpellHospital & " as any e-mails sent to this address will not be reviewed.</div></td></tr></table></body></html>"
    Else
        HTMLFooting = ""
        HTMLFooting = HTMLFooting & "</td></tr>"
        HTMLFooting = HTMLFooting & "<tr><td><div align=center style='background : #cccccc'><strong>Please Note</strong>"
        HTMLFooting = HTMLFooting & "<br/>Any comments should be sent to "
        HTMLFooting = HTMLFooting & oDataRS!SpellHospital ' "Norfolk & Norwich University Hospital NHS Trust"
        HTMLFooting = HTMLFooting & " as any e-mails sent to this address will not be reviewed.</div></td></tr>"
        HTMLFooting = HTMLFooting & "</table>"
    End If

End Property

Public Property Get Notification() As Boolean

    Notification = True

End Property


