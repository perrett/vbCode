VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XML_ReportMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
'
'   Project Name : ICEMsg
'       Identification : ICEMsg.vbp
'   Class Name : XML_ReportMsg
'       Identification : XML_ReportMsg.cls
'
'   Copyright (c) 2003 - Anglia Healthcare Systems Ltd
'
'   Author :
'       Simon Jones, 17th October 2003 - 31st October 2003
'
'
'                                    www.ahsl.com
' *************************************************************************

Option Explicit
Public MessageFormat As String
Private oXMLMessage As AHSLXML.XMLMessage

Private oDataRS As ADODB.Recordset

Private SenderEDICode As String
Private ReceiverEDICode As String
Private sHTMLHeading As String
Private LetterType As String

Private Type CONSULTANT
    Code As String
    Name As String
    Specialty As String
    OrgIdentifier As String
    OrgName As String
    OrgAddressLine() As String
    OrgPostCode As String
    WardKey As String
    WardName As String
End Type
Private ConsultantDetails As CONSULTANT
Private SenderDetails As CONSULTANT

Private Sub Class_Initialize()

    Set oXMLMessage = New AHSLXML.XMLMessage
    Set oDataRS = New ADODB.Recordset

   sDOMAIN = Read_Ini_Var("LETTER", "DOMAIN", INIFile)
   
   blnRemoveImports = (Read_Ini_Var("LETTER", "RemoveImport", INIFile) = 1)
   
   defUserIndex = Read_Ini_Var("LETTER", "UserIndex", INIFile)
   If defUserIndex = "" Or IsNumeric(defUserIndex) = False Then
      defUserIndex = -1
   End If
   
   sDocManSrc = Read_Ini_Var("LETTER", "DocManSourceId", INIFile)
   
End Sub

Public Function CreateMessage(ByVal ReportId As Long) As Variant

    Dim iceCmd As New ADODB.Command
    Dim oRecordset As New ADODB.Recordset
    Dim blnDoingTrans As Boolean
    
    On Error GoTo procEH
    msgControl.ReportStatus = 0

    eClass.FurtherInfo = "Get Letter Type for XML Message"

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
    
    frmMain.NewFileRequired

    If LetterType <> "" Then
        If MessageFormat = "DOCMAN,STH XML" Then
            Select Case LetterType
                Case "DISCHARGE", "DOR", "DTH", "TCT", "DEATH"
                    eClass.FurtherInfo = "Build STH Docman Letter"
                    CreateMessage = XML_DocMan_Discharge(ReportId)
                Case Else
                    msgControl.ReportStatus = RS_GENERAL
                    msgControl.LogReportMessage RS_GENERAL, "XML generation of the Letter Type (" & LetterType & ") not enabled"
            
            End Select
            
        Else
            Select Case LetterType
                Case "DISCHARGE", "DOR", "DTH", "TCT", "DEATH"
                    eClass.FurtherInfo = "Build Discharge Letter XML"
                    CreateMessage = XML_Discharge(ReportId)
                Case Else
                    msgControl.ReportStatus = RS_GENERAL
                    msgControl.LogReportMessage RS_GENERAL, "XML generation of the Letter Type (" & LetterType & ") not enabled"
            End Select
        End If
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
    eClass.CurrentProcedure = "XML_ReportMsg.CreateMessage"
    eClass.Add Err.Number, Err.Description, Err.Source
    If blnDoingTrans Then oRecordset.Close
    Set oRecordset = Nothing
    Set iceCmd = Nothing

End Function

Private Function XML_Discharge(ByVal ReportId As Long) As String

    Dim hpData As clsHealthParties
    Dim objViewDischarge As Object
    Dim strString As String
    Dim strTemp As String
    Dim strHyperlinkID As String
    Dim intLoop As Integer
    Dim intInnerLoop As Integer
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim ErrorCode As String
    Dim ErrorDescription As String
    Dim Success As Boolean
    Dim AddressLine As Boolean

    On Error GoTo procEH
    XML_Discharge = ""
    AddressLine = False

    eClass.FurtherInfo = "Get Data for XML Message"
    If Not XML_MainData(ReportId) Then Exit Function

    ' Remove the curly brackets from the HyperlinkID
    If DBVERSION < 330 Or DBVERSION > 4999 Then strHyperlinkID = Mid(oDataRS!HyperlinkID, 2, Len(oDataRS!HyperlinkID) - 2)

    ' Now use the ICEDesktopDischarge.Letter class to obtain HTML of Clinical Letter
    Select Case LetterType
        Case "DISCHARGE", "DOR"
            HTMLHeading = "Discharge Summary"
        Case "DTH", "DEATH"
            HTMLHeading = "Death Notification"
        Case "TCT"
            HTMLHeading = "Clinical Letter"
    End Select

    If DBVERSION > 329 And DBVERSION < 5000 Then
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
    Else
        eClass.FurtherInfo = "Call ICDesktopDischarge.Letter Class"
        Set objViewDischarge = CreateObject("ICEDesktopDischarge.Letter")
            objViewDischarge.dbConnection = "FILE NAME=" & App.Path & "\ice.udl"
            strString = HTMLHeading & objViewDischarge.View(oDataRS!LetterID & "," & strHyperlinkID, "", False) & HTMLFooting
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
    End If

    InitialiseConsultant
    eClass.FurtherInfo = "Get Consultant Details for XML Message for " & oDataRS!SpellIndex
    If Not XML_ConsData(oDataRS!SpellIndex) Then Exit Function
    eClass.FurtherInfo = "Get Sender Details for XML Message for " & ReportId
    If Not XML_SenderData(ReportId) Then Exit Function

    ' Assign XML Declaration
    oXMLMessage.XMLDeclaration = "<?xml version=""1.0"" encoding=""UTF-8""?>"

    ' Assign DOCTYPE Declaration
    Select Case MessageFormat
        Case "XML,RPTMSG v1.2"
            oXMLMessage.DOCTYPEDeclaration = "<!DOCTYPE ReportMsg SYSTEM ""ReportMsg03.dtd"">"
        Case "DOCMAN,KGH XML"
            oXMLMessage.DOCTYPEDeclaration = "<!DOCTYPE ReportMsg SYSTEM ""ReportMsg_YD_01.dtd"">"
    End Select
   
    'oXMLMessage.DOCTYPEDeclaration = "<!DOCTYPE ReportMsg SYSTEM ""ReportMsg03.dtd"">"

    ' Add the Root Element
    oXMLMessage.root.Load = "<ReportMsg MsgRcptAckRequest=""Always"" MsgUrgency=""Normal"" MsgStatus=""Live""/>"
    
    ' Add MsgId Element
    Select Case MessageFormat
        Case "XML,RPTMSG v1.2"
            oXMLMessage.element.Add 0, "MsgId", repListRS!Report_Identifier
        Case "DOCMAN,KGH XML"
            oXMLMessage.element.Add 0, "MsgId", FormatGuid(GetGuid)
    End Select

    'oXMLMessage.Element.Add 0, "MsgId", repListRS!Report_Identifier

    ' Add MsgIssueDate Element, date will get applied at then of this process.
    oXMLMessage.element.Add 0, "MsgIssueDate", Format(Now, "YYYYMMDDhhmmss")
    
    ' Add MsgSender Element
    oXMLMessage.element.Add 0, "MsgSender", SenderEDICode
    
    ' Add MsgRecipient Element
    oXMLMessage.element.Add 0, "MsgRecipient", ReceiverEDICode
    
    ' Add ServiceRequester Elements
    oXMLMessage.element.Add 0, "ServiceRequester"
    oXMLMessage.element.Add 1, "IdValue", "1"
    'If (repListRS!GP_Code = oDataRS!referrer) And (oDataRS!referrer <> "") Then
    '    oXMLMessage.Element.Add 1, "IdValue", "2"
    'Else
    '    If (repListRS!GP_Code = oDataRS!RegisteredGP) And (oDataRS!RegisteredGP <> "") Then
    '        oXMLMessage.Element.Add 1, "IdValue", "1"
    '    Else
    '        If oDataRS!RegisteredGP <> "" Then
    '            oXMLMessage.Element.Add 1, "IdValue", "1"
    '        ElseIf oDataRS!referrer <> "" Then
    '            oXMLMessage.Element.Add 1, "IdValue", "2"
    '        Else
    '            oXMLMessage.Element.Add 1, "IdValue", "1"
    '        End If
    '    End If
    'End If

    ' Add ServiceProvider Elements
    oXMLMessage.element.Add 0, "ServiceProvider"
    If ConsultantDetails.Code = SenderDetails.Code Then
        oXMLMessage.element.Add 1, "IdValue", "3"
    Else
        oXMLMessage.element.Add 1, "IdValue", "4"
    End If

    ' Add PatientMatchingInfo Elements
    oXMLMessage.element.Add 0, "PatientMatchingInfo"
    ' Add PatientID Child Elements
    oXMLMessage.element.Add 1, "PatientId"
    ' NHS Number
    strTemp = IIf(IsNull(oDataRS!NHSNumber), "", oDataRS!NHSNumber)
    If strTemp <> "" Then
        oXMLMessage.element.Add 2, "Id"
        oXMLMessage.element.Add 3, "IdType", "Nhs"
        oXMLMessage.element.Add 3, "IdValue", strTemp
    End If
    ' Hospital Number
    strTemp = IIf(IsNull(oDataRS!HospitalNumber), "", oDataRS!HospitalNumber)
    If strTemp <> "" Then
        oXMLMessage.element.Add 2, "Id"
        oXMLMessage.element.Add 3, "IdType", "Provider"
        oXMLMessage.element.Add 3, "IdValue", strTemp
    End If
    ' Practice Number
    strTemp = IIf(IsNull(oDataRS!PracticeNumber), "", oDataRS!PracticeNumber)
    If strTemp <> "" Then
        oXMLMessage.element.Add 2, "Id"
        oXMLMessage.element.Add 3, "IdType", "Requester"
        oXMLMessage.element.Add 3, "IdValue", strTemp
    End If
    ' Add Person Name Child Elements
    oXMLMessage.element.Add 1, "PersonName_s"
    oXMLMessage.element.Add 2, "StructPersonName"
    strTemp = IIf(IsNull(oDataRS!Surname), "", oDataRS!Surname)
    oXMLMessage.element.Add 3, "FamilyName", strTemp
    strTemp = IIf(IsNull(oDataRS!Forename), "", oDataRS!Forename)
    oXMLMessage.element.Add 3, "GivenName", strTemp
    strTemp = IIf(IsNull(oDataRS!Midname), "", oDataRS!Midname)
    oXMLMessage.element.Add 3, "MiddleName", strTemp
    strTemp = IIf(IsNull(oDataRS!Title), "", oDataRS!Title)
    oXMLMessage.element.Add 3, "Title", strTemp
    ' Add Date Of Birth Child Element
    strTemp = IIf(IsNull(oDataRS!DateOfBirth), "", oDataRS!DateOfBirth)
    oXMLMessage.element.Add 1, "BirthDate", Format(strTemp, "yyyymmdd")
    ' Add Sex Child Element
    strTemp = IIf(IsNull(oDataRS!Sex), "", oDataRS!Sex)
    oXMLMessage.element.Add 1, "Sex", strTemp
    ' Add Address Child Elements
    oXMLMessage.element.Add 1, "Address_u"
    strTemp = IIf(IsNull(oDataRS!PostCode), "", oDataRS!PostCode)
    oXMLMessage.element.Add 2, "PostCode", strTemp
    oXMLMessage.element.Add 2, "UnstructAddress"
    For intLoop = 1 To 5
        strTemp = IIf(IsNull(oDataRS("AddressLine" & intLoop)), "", oDataRS("AddressLine" & intLoop))
        If strTemp <> "" Then
            AddressLine = True
            oXMLMessage.element.Add 3, "UnstructAddressLine", strTemp
        End If
    Next intLoop
    If Not AddressLine Then oXMLMessage.element.Add 3, "UnstructAddressLine"

    ' Add AgentsDirectory Elements
    oXMLMessage.element.Add 0, "AgentsDirectory"

    ' Add Recipient Elements
    oXMLMessage.element.Add 1, "AgentInContext"
    oXMLMessage.element.Add 2, "IdValue", "1"
    oXMLMessage.element.Add 2, "Person"
    oXMLMessage.element.Add 3, "AgentId"
    oXMLMessage.element.Add 4, "IdType", "GP"
    If IsNull(repListRS!GP_Code) Then
        Set hpData = New clsHealthParties

        hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
        hpData.LTSIndex = repListRS!EDI_LTS_Index
        hpData.Read repListRS!EDI_Report_Index, True, True

        oXMLMessage.element.Add 4, "IdValue", hpData.HP902Code

        Set hpData = Nothing
    Else
        oXMLMessage.element.Add 4, "IdValue", repListRS!GP_Code
    End If
    oXMLMessage.element.Add 2, "AgentRel"
    oXMLMessage.element.Add 3, "AgentRelType", "3"
    oXMLMessage.element.Add 3, "Organisation"
    oXMLMessage.element.Add 4, "AgentId"
    oXMLMessage.element.Add 5, "IdType", "Practice"
    oXMLMessage.element.Add 5, "IdValue", repListRS!EDI_Loc_Nat_Code_To

    ' Add Registered GP Elements
    If oDataRS!RegisteredGP <> "" Then
        oXMLMessage.element.Add 1, "AgentInContext"
        oXMLMessage.element.Add 2, "IdValue", "2"
        oXMLMessage.element.Add 2, "Person"
        oXMLMessage.element.Add 3, "AgentId"
        oXMLMessage.element.Add 4, "IdType", "GP"
        oXMLMessage.element.Add 4, "IdValue", oDataRS!RegisteredGP
        oXMLMessage.element.Add 2, "AgentRel"
        oXMLMessage.element.Add 3, "AgentRelType", "3"
        oXMLMessage.element.Add 3, "Organisation"
        oXMLMessage.element.Add 4, "AgentId"
        oXMLMessage.element.Add 5, "IdType", "Practice"
        oXMLMessage.element.Add 5, "IdValue", oDataRS!PracticeNationalCode
    End If

    ' Add Registered Referreer Elements
    'If oDataRS!referrer <> "" Then
    '    oXMLMessage.Element.Add 1, "AgentInContext"
    '    oXMLMessage.Element.Add 2, "IdValue", "2"
    '    oXMLMessage.Element.Add 2, "Person"
    '    oXMLMessage.Element.Add 3, "AgentId"
    '    If Left(oDataRS!referrer, 1) = "C" And IsNumeric(Mid(oDataRS!referrer, 2, Len(oDataRS!referrer) - 1)) Then
    '        oXMLMessage.Element.Add 4, "IdType", "Specialist"
    '    Else
    '        oXMLMessage.Element.Add 4, "IdType", "GP"
    '    End If
    '    oXMLMessage.Element.Add 4, "IdValue", oDataRS!referrer
    '    If (Left(oDataRS!referrer, 1) = "C" And IsNumeric(Mid(oDataRS!referrer, 2, Len(oDataRS!referrer) - 1))) Or (Left(oDataRS!referrer, 1) <> "C" And oDataRS!ReferrerOrganisation <> "") Then
    '        oXMLMessage.Element.Add 2, "AgentRel"
    '        oXMLMessage.Element.Add 3, "AgentRelType", "3"
    '        oXMLMessage.Element.Add 3, "Organisation"
    '        oXMLMessage.Element.Add 4, "AgentId"
    '        If Left(oDataRS!referrer, 1) = "C" And IsNumeric(Mid(oDataRS!referrer, 2, Len(oDataRS!referrer) - 1)) Then
    '            oXMLMessage.Element.Add 5, "IdType", "Provider"
    '            oXMLMessage.Element.Add 5, "IdValue", SenderDetails.OrgIdentifier
    '        Else
    '            oXMLMessage.Element.Add 5, "IdType", "Practice"
    '            oXMLMessage.Element.Add 5, "IdValue", oDataRS!ReferrerOrganisation
    '        End If
    '    End If
    'End If
    
    ' Add Attending Consultant Elements
    With ConsultantDetails
        oXMLMessage.element.Add 1, "AgentInContext"
        oXMLMessage.element.Add 2, "IdValue", "3"
        oXMLMessage.element.Add 2, "Person"
        oXMLMessage.element.Add 3, "AgentId"
        oXMLMessage.element.Add 4, "IdType", IIf(Len(.Code) = 8 And Left(.Code, 1) = "C" And IsNumeric(Mid(.Code, 2)), "Specialist", "Internal")
        oXMLMessage.element.Add 4, "IdValue", .Code
        oXMLMessage.element.Add 3, "PersonName_u"
        oXMLMessage.element.Add 4, "PersonNameType", "CU"
        oXMLMessage.element.Add 4, "String", .Name
        oXMLMessage.element.Add 2, "AgentRel"
        oXMLMessage.element.Add 3, "AgentRelType", "3"
        oXMLMessage.element.Add 3, "Organisation"
        oXMLMessage.element.Add 4, "AgentId"
        oXMLMessage.element.Add 5, "IdType", "Internal"
        oXMLMessage.element.Add 5, "IdValue", .WardKey
        oXMLMessage.element.Add 4, "PartyMedicalSpecialty"
        oXMLMessage.element.Add 5, "Code", .Specialty
        oXMLMessage.element.Add 4, "OrgName", .WardName
        oXMLMessage.element.Add 2, "AgentRel"
        oXMLMessage.element.Add 3, "AgentRelType", "5"
        oXMLMessage.element.Add 3, "Organisation"
        oXMLMessage.element.Add 4, "AgentId"
        oXMLMessage.element.Add 5, "IdType", "Provider"
        oXMLMessage.element.Add 5, "IdValue", .OrgIdentifier
        oXMLMessage.element.Add 4, "Address_u"
        oXMLMessage.element.Add 5, "PostCode", .OrgPostCode
        oXMLMessage.element.Add 5, "UnstructAddress"
        AddressLine = False
        For intLoop = 1 To 5
            If .OrgAddressLine(intLoop) <> "" Then
                AddressLine = True
                oXMLMessage.element.Add 6, "UnstructAddressLine", .OrgAddressLine(intLoop)
            End If
        Next intLoop
        If Not AddressLine Then oXMLMessage.element.Add 6, "UnstructAddressLine"
        oXMLMessage.element.Add 4, "OrgName", .OrgName
    End With

    ' Add Sending Consultant Elements
    If ConsultantDetails.Code <> SenderDetails.Code Then
        With SenderDetails
            oXMLMessage.element.Add 1, "AgentInContext"
            oXMLMessage.element.Add 2, "IdValue", "4"
            oXMLMessage.element.Add 2, "Person"
            oXMLMessage.element.Add 3, "AgentId"
            oXMLMessage.element.Add 4, "IdType", IIf(Len(.Code) = 8 And Left(.Code, 1) = "C" And IsNumeric(Mid(.Code, 2)), "Specialist", "Internal")
            oXMLMessage.element.Add 4, "IdValue", .Code
            oXMLMessage.element.Add 3, "PersonName_u"
            oXMLMessage.element.Add 4, "PersonNameType", "CU"
            oXMLMessage.element.Add 4, "String", .Name
            oXMLMessage.element.Add 2, "AgentRel"
            oXMLMessage.element.Add 3, "AgentRelType", "3"
            oXMLMessage.element.Add 3, "Organisation"
            oXMLMessage.element.Add 4, "AgentId"
            oXMLMessage.element.Add 5, "IdType", "Provider"
            oXMLMessage.element.Add 5, "IdValue", .OrgIdentifier
            oXMLMessage.element.Add 4, "Address_u"
            oXMLMessage.element.Add 5, "PostCode", .OrgPostCode
            oXMLMessage.element.Add 5, "UnstructAddress"
            AddressLine = False
            For intLoop = 1 To 5
                If .OrgAddressLine(intLoop) <> "" Then
                    AddressLine = True
                    oXMLMessage.element.Add 6, "UnstructAddressLine", .OrgAddressLine(intLoop)
                End If
            Next intLoop
            If Not AddressLine Then oXMLMessage.element.Add 6, "UnstructAddressLine"
            oXMLMessage.element.Add 4, "OrgName", .OrgName
        End With
    End If

    ' Add Referer Elements
    If oDataRS!referrer <> "" Then
        oXMLMessage.element.Add 1, "AgentInContext"
        oXMLMessage.element.Add 2, "IdValue", "5"
        oXMLMessage.element.Add 2, "Person"
        oXMLMessage.element.Add 3, "AgentId"
        If Left(oDataRS!referrer, 1) = "C" And IsNumeric(Mid(oDataRS!referrer, 2, Len(oDataRS!referrer) - 1)) Then
            oXMLMessage.element.Add 4, "IdType", "Specialist"
        Else
            oXMLMessage.element.Add 4, "IdType", "GP"
        End If
        oXMLMessage.element.Add 4, "IdValue", oDataRS!referrer
        If (Left(oDataRS!referrer, 1) = "C" And IsNumeric(Mid(oDataRS!referrer, 2, Len(oDataRS!referrer) - 1))) Or (Left(oDataRS!referrer, 1) <> "C" And oDataRS!ReferrerOrganisation <> "") Then
            oXMLMessage.element.Add 2, "AgentRel"
            oXMLMessage.element.Add 3, "AgentRelType", "3"
            oXMLMessage.element.Add 3, "Organisation"
            oXMLMessage.element.Add 4, "AgentId"
            If Left(oDataRS!referrer, 1) = "C" And IsNumeric(Mid(oDataRS!referrer, 2, Len(oDataRS!referrer) - 1)) Then
                oXMLMessage.element.Add 5, "IdType", "Provider"
                oXMLMessage.element.Add 5, "IdValue", SenderDetails.OrgIdentifier
            Else
                oXMLMessage.element.Add 5, "IdType", "Practice"
                oXMLMessage.element.Add 5, "IdValue", oDataRS!ReferrerOrganisation
            End If
        End If
    End If

    ' Add ReportedService Elements
    oXMLMessage.element.Add 0, "ReportedService"
    Select Case oDataRS!AdministrativeCategory
        Case "01"
            strTemp = "PPI"
        Case "02"
            strTemp = "PPR"
        Case "03"
            strTemp = "PAM"
        Case "04"
            strTemp = "PCT"
        Case Else
            strTemp = ""
    End Select
    oXMLMessage.element.Add 1, "PaymentCategory", strTemp
    oXMLMessage.element.Add 1, "ServiceReportStatus", "CO"
    oXMLMessage.element.Add 1, "ReportedEventItem"
    oXMLMessage.element.Add 2, "Cuid"
    oXMLMessage.element.setAttribute "IdScope", "Message"
    oXMLMessage.element.Add 2, "RelAgent"
    oXMLMessage.element.Add 3, "RelAgentRole"
    oXMLMessage.element.Add 4, "Code", "32"
    If oDataRS!referrer = "" Then
        oXMLMessage.element.Add 3, "IdValue", "2"
    Else
        oXMLMessage.element.Add 3, "IdValue", "5"
    End If
    oXMLMessage.element.Add 2, "RelAgent"
    oXMLMessage.element.Add 3, "RelAgentRole"
    oXMLMessage.element.Add 4, "Code", "21"
    oXMLMessage.element.Add 3, "IdValue", "3"
    oXMLMessage.element.Add 2, "RelDate"
    oXMLMessage.element.Add 3, "Date", Format(oDataRS!SpellStartDate, "yyyymmddhhmm")
    oXMLMessage.element.Add 3, "RelDateRole"
    oXMLMessage.element.Add 4, "Code", "81"
    oXMLMessage.element.Add 2, "RelDate"
    oXMLMessage.element.Add 3, "Date", Format(oDataRS!SpellEndDate, "yyyymmddhhmm")
    oXMLMessage.element.Add 3, "RelDateRole"
    oXMLMessage.element.Add 4, "Code", "82"
    Select Case oDataRS!SpellVisitType
        Case "IP", "I"
            strTemp = "IP"
        Case "O"
            strTemp = "FV"
        Case "OF"
            strTemp = "SV"
        Case "A", "AE", "TC"
            strTemp = "RC"
        Case Else
            strTemp = "UN"
    End Select
    oXMLMessage.element.Add 2, "ServiceEventType", strTemp
    oXMLMessage.element.Add 2, "AdministrativeOutcome"
    Select Case oDataRS!DischargeMethod
        Case "1"
            strTemp = "DC"
        Case "2"
            strTemp = "DS"
        Case "3"
            strTemp = "DH"
        Case "4"
            strTemp = "PD"
        Case "5"
            strTemp = "SB"
        Case Else
            strTemp = "UN"
    End Select
    oXMLMessage.element.Add 3, "Code", strTemp

    ' Add PatientAdminInfo Elements
    If oDataRS!RegisteredGP <> "" Then
        oXMLMessage.element.Add 0, "PatientAdminInfo"
        oXMLMessage.element.Add 1, "RelAgent"
        oXMLMessage.element.Add 2, "RelAgentRole"
        oXMLMessage.element.Add 3, "Code", "PRG"
        oXMLMessage.element.Add 2, "IdValue", "2"
    End If

    ' Add ClincialReport Elements
    oXMLMessage.element.Add 0, "ClinicalReport"
    
    If MessageFormat = "DOCMAN,KGH XML" Then
       oXMLMessage.element.Add 1, "ReportID", repListRS!Report_Identifier
   
       Select Case oDataRS!SpellVisitType
       Case "O", "OP", "OF"
          oXMLMessage.element.Add 1, "ReportType", "Outpatient Letter"
          oXMLMessage.element.Add 1, "ReportCode", "OP"
       Case "AE", "A"
          oXMLMessage.element.Add 1, "ReportType", "Accident & Emergency Letter"
          oXMLMessage.element.Add 1, "ReportCode", "A&E"
       Case "I", "IP"
          oXMLMessage.element.Add 1, "ReportType", "Inpatient Letter"
          oXMLMessage.element.Add 1, "ReportCode", "IP"
       Case Else
          oXMLMessage.element.Add 1, "ReportType", "Clinical Letter"
          oXMLMessage.element.Add 1, "ReportCode", "CL"
       End Select
       oXMLMessage.element.Add 1, "EventDate", Format(oDataRS!SpellStartDate, "yyyymmddhhmm")
       oXMLMessage.element.Add 1, "GPName", IIf(IsNull(repListRS!EDI_OP_Name), "unknown", repListRS!EDI_OP_Name)
    End If
    
    oXMLMessage.element.Add 1, "TextItem"
    oXMLMessage.element.setAttribute "RcStatus", "Current"
    oXMLMessage.element.Add 2, "Cuid"
    oXMLMessage.element.setAttribute "IdScope", "Message"
    oXMLMessage.element.Add 2, "TextMarkupIndicator", "-//IETF//DTD HTML//EN"
    If repListRS!EDI_EMIS_System Then
        strString = Replace(strString, "margin-left: 0pt; ", "")
        strString = Replace(strString, "margin-right: 0pt; ", "")
        strString = Replace(strString, "margin-top: 0pt; ", "")
        strString = Replace(strString, "margin-bottom: 0pt; ", "")
        strString = Replace(strString, "font-weight: normal; ", "")
        strString = Replace(strString, "font-style: normal; ", "")
    End If
    
    If blnRemoveImports Then
       strString = Replace(strString, "@import url(""../css/menubar.css"");", "")
    End If
    
    oXMLMessage.element.AddCDATA 2, "TextBlock", strString

    ' Finished building XML message
    XML_Discharge = oXMLMessage.MessageContents
    
    Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "XML_ReportMsg.XML_Discharge"
   eClass.Add Err.Number, Err.Description, Err.Source

End Function

Private Function XML_DocMan_Discharge(ByVal ReportId As Long) As String

    Dim hpData As clsHealthParties
    Dim objViewDischarge As Object
    Dim strString As String
    Dim strTemp As String
    Dim strHyperlinkID As String
    Dim intLoop As Integer
    Dim intInnerLoop As Integer
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim ErrorCode As String
    Dim ErrorDescription As String
    Dim Success As Boolean
    Dim AddressLine As Boolean

    On Error GoTo procEH
    XML_DocMan_Discharge = ""
    AddressLine = False

    eClass.FurtherInfo = "Get Data for XML Message"
    If Not XML_MainData(ReportId) Then Exit Function

    ' Remove the curly brackets from the HyperlinkID
    If DBVERSION < 330 Or DBVERSION > 4999 Then strHyperlinkID = Mid(oDataRS!HyperlinkID, 2, Len(oDataRS!HyperlinkID) - 2)

    ' Now use the ICEDesktopDischarge.Letter class to obtain HTML of Clinical Letter
    Select Case LetterType
        Case "DISCHARGE", "DOR"
            HTMLHeading = "Discharge Summary"
        Case "DTH", "DEATH"
            HTMLHeading = "Death Notification"
        Case "TCT"
            HTMLHeading = "Clinical Letter"
    End Select

    If DBVERSION > 329 And DBVERSION < 5000 Then
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
    Else
        eClass.FurtherInfo = "Call ICDesktopDischarge.Letter Class"
        Set objViewDischarge = CreateObject("ICEDesktopDischarge.Letter")
            objViewDischarge.dbConnection = "FILE NAME=" & App.Path & "\ice.udl"
            strString = HTMLHeading & objViewDischarge.View(oDataRS!LetterID & "," & strHyperlinkID, "", False) & HTMLFooting
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
    End If

    InitialiseConsultant
    eClass.FurtherInfo = "Get Consultant Details for XML Message for " & oDataRS!SpellIndex
    If Not XML_ConsData(oDataRS!SpellIndex) Then Exit Function
    eClass.FurtherInfo = "Get Sender Details for XML Message for " & ReportId
    If Not XML_SenderData(ReportId) Then Exit Function

    With oXMLMessage

        ' Assign XML Declaration
        .XMLDeclaration = "<?xml version=""1.0"" encoding=""UTF-8""?>"
    
        ' Assign DOCTYPE Declaration
        .DOCTYPEDeclaration = "<!DOCTYPE ReportMsg SYSTEM ""ReportMsg_EDT_03.dtd"">"
    
        ' Add the Root Element
        .root.Load = "<ReportMsg MsgRcptAckRequest=""Always"" MsgUrgency=""Normal"" MsgStatus=""Live""/>"
        
        ' Add MsgId Element
        .element.Add 0, "MsgId", ReportId & "^" & repListRS!Report_Identifier
    
        ' Add MsgIssueDate Element, date will get applied at then of this process.
        .element.Add 0, "MsgIssueDate", Format(Now, "YYYYMMDDhhmmss")
        
        ' Add MsgSender Element
        .element.Add 0, "MsgSender", SenderEDICode
        
        ' Add MsgRecipient Element
        .element.Add 0, "MsgRecipient", ReceiverEDICode
        
        ' Add ServiceRequester Elements
        .element.Add 0, "ServiceRequester"
        .element.Add 1, "IdValue", "1"
    
        ' Add ServiceProvider Elements
        .element.Add 0, "ServiceProvider"
        .element.Add 1, "IdValue", "2"
        
'        If ConsultantDetails.Code = SenderDetails.Code Then
'            .element.Add 1, "IdValue", "3"
'        Else
'            .element.Add 1, "IdValue", "4"
'        End If
    
        ' Add PatientMatchingInfo Elements
        .element.Add 0, "PatientMatchingInfo"
        ' Add PatientID Child Elements
        .element.Add 1, "PatientId"
        ' NHS Number
        strTemp = IIf(IsNull(oDataRS!NHSNumber), "", oDataRS!NHSNumber)
        .element.Add 2, "Id"
        .element.Add 3, "IdType", "Nhs"
        
        If strTemp = "" Then
            .element.Add 3, "IdValue", "9999999999"
        Else
            .element.Add 3, "IdValue", strTemp
        End If
        
        ' Add Person Name Child Elements
        .element.Add 1, "PersonName_s"
        .element.Add 2, "PersonNameType", "CU"
        
        .element.Add 2, "StructPersonName"
        strTemp = IIf(IsNull(oDataRS!Surname), "", oDataRS!Surname)
        .element.Add 3, "FamilyName", strTemp
        strTemp = IIf(IsNull(oDataRS!Forename), "", oDataRS!Forename)
        .element.Add 3, "GivenName", strTemp
        
        ' Add Date Of Birth Child Element
        strTemp = IIf(IsNull(oDataRS!DateOfBirth), "", oDataRS!DateOfBirth)
        .element.Add 1, "BirthDate", Format(strTemp, "yyyymmdd")
        
        ' Add Sex Child Element
        'strTemp = IIf(IsNull(oDataRS!Sex), "", oDataRS!Sex)
        .element.Add 1, "Sex", oDataRS!patSex
        
        .element.Add 0, "Parties"
        ' Add Recipient Elements
        .element.Add 1, "Party"
        .element.setAttribute "AgentId", "1"
        
        .element.Add 2, "HcpCode", oDataRS!RegisteredGP
        .element.setAttribute "IdType", "GP"
    
        .element.Add 2, "HcpName" ', oDataRS!RegisteredGP
        
        .element.Add 2, "OrgCode", oDataRS!PracticeNationalCode
        .element.setAttribute "IdType", "Practice"
        
        .element.Add 2, "OrgName"
        .element.Add 2, "LogicalAddress"
        
        '   Add Sender Elements
        .element.Add 1, "Party"
        .element.setAttribute "AgentId", "2"
        
        .element.Add 2, "HcpCode", ConsultantDetails.Code
        .element.setAttribute "IdType", "Internal"
        
        .element.Add 2, "HcpName", ConsultantDetails.Name
        
        .element.Add 2, "BossCode"
        .element.setAttribute "IdType", "Specialist"
        
        .element.Add 2, "BossName"
        
        .element.Add 2, "DepartmentCode", ConsultantDetails.WardKey
        .element.setAttribute "IdType", "Internal"
        
        .element.Add 2, "DepartmentName", ConsultantDetails.WardName
        
        .element.Add 2, "OrgCode", ConsultantDetails.OrgIdentifier
        .element.setAttribute "IdType", "Provider"
        .element.Add 2, "OrgName", ConsultantDetails.OrgName
        
    
    ' Add ClincialReport Elements
        .element.Add 0, "ClinicalReport"
        
        .element.Add 1, "ReportID", FormatGuid(GetGuid)
   
        Select Case oDataRS!SpellVisitType
        Case "O", "OP", "OF"
           .element.Add 1, "ReportType", "Outpatient Letter"
           .element.Add 1, "ReportCode", "OP"
        Case "AE", "A"
           .element.Add 1, "ReportType", "Accident & Emergency Letter"
           .element.Add 1, "ReportCode", "A&E"
        Case "I", "IP"
           .element.Add 1, "ReportType", "Inpatient Letter"
           .element.Add 1, "ReportCode", "IP"
        Case Else
           .element.Add 1, "ReportType", "Clinical Letter"
           .element.Add 1, "ReportCode", "CL"
        End Select
        
        .element.Add 1, "ReportSourceType", sDocManSrc
        .element.Add 1, "ReportContributor"
       
       
        .element.Add 1, "EventDate", Format(oDataRS!SpellStartDate, "yyyymmddhhmm")
        .element.Add 1, "EventDateEnd"
        .element.Add 1, "EventUrgencyStatus"
        
        .element.Add 1, "DocCreationDate", Format(Now(), "yyyymmddhhmm")
    
        .element.Add 1, "TextItem"
        .element.setAttribute "RcStatus", "Current"
        .element.Add 2, "Cuid"
        .element.setAttribute "IdScope", "Message"
        .element.Add 2, "TextMarkupIndicator", "-//IETF//DTD HTML//EN"
    
        If repListRS!EDI_EMIS_System Then
            strString = Replace(strString, "margin-left: 0pt; ", "")
            strString = Replace(strString, "margin-right: 0pt; ", "")
            strString = Replace(strString, "margin-top: 0pt; ", "")
            strString = Replace(strString, "margin-bottom: 0pt; ", "")
            strString = Replace(strString, "font-weight: normal; ", "")
            strString = Replace(strString, "font-style: normal; ", "")
        End If
    
        .element.AddCDATA 2, "TextBlock", strString

    ' Finished building XML message
        XML_DocMan_Discharge = .MessageContents
       End With
       
    Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "XML_ReportMsg.XML_Discharge"
   eClass.Add Err.Number, Err.Description, Err.Source

End Function

Public Property Let SenderID(ByVal vNewValue As String)

    SenderEDICode = "0080" & vNewValue

End Property

Public Property Let ReceiverID(ByVal vNewValue As String)

    ReceiverEDICode = "0080" & vNewValue

End Property

Private Function XML_MainData(ByVal ServiceReportID As Long) As Boolean

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
        XML_MainData = False
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

        XML_MainData = True
    End If
    
    Set iceCmd = Nothing
    Exit Function
    
procEH:
    If eClass.Behaviour = -1 Then
        Stop
        Resume
    End If
    eClass.CurrentProcedure = "XML_ReportMsg.XML_MainData"
    eClass.Add Err.Number, Err.Description, Err.Source
    Set iceCmd = Nothing
    XML_MainData = False

End Function

Private Function XML_ConsData(ByVal SpellIndex As Long) As Boolean

    Dim iceCmd As New ADODB.Command
    Dim XMLConsDataRS As New ADODB.Recordset
    Dim blnDoingTrans As Boolean
    Dim intLoop As Integer

    On Error GoTo procEH

    With iceCmd
        .ActiveConnection = iceCon
        .CommandType = adCmdStoredProc
        .CommandText = "ICEMSG_Get_Consultant_Details_for_XML"
        .Parameters.Append .CreateParameter("SpellIndex", adInteger, adParamInput, , SpellIndex)
        .Parameters.Append .CreateParameter("OrgCode", adVarChar, adParamInput, 6, orgCode)
        .Parameters.Append .CreateParameter("SystemProviderIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
        Set XMLConsDataRS = .Execute
        blnDoingTrans = True
    End With

    If XMLConsDataRS.BOF And XMLConsDataRS.EOF Then
        msgControl.ReportStatus = RS_DATA_INTEGRITY
        msgControl.LogReportMessage RS_DATA_INTEGRITY, "No Consultant details available for Spell ID provided"
        XML_ConsData = False
    Else
        With ConsultantDetails
            .Code = IIf(IsNull(XMLConsDataRS!NationalCode), XMLConsDataRS!LocalCode, XMLConsDataRS!NationalCode)
            .Name = ""
            If Not IsNull(XMLConsDataRS!Surname) Then
                If Trim(XMLConsDataRS!Surname) <> "" Then
                    .Name = XMLConsDataRS!Surname
                End If
            End If
            If Not IsNull(XMLConsDataRS!Forename) Then
                If Trim(XMLConsDataRS!Forename) <> "" Then
                    If .Name = "" Then
                        .Name = XMLConsDataRS!Forename
                    Else
                        .Name = .Name & ", " & XMLConsDataRS!Forename
                    End If
                End If
            End If
            .Specialty = IIf(IsNull(XMLConsDataRS!Specialty), "", XMLConsDataRS!Specialty)
            .OrgIdentifier = IIf(IsNull(XMLConsDataRS!orgId), "", XMLConsDataRS!orgId)
            .OrgName = IIf(IsNull(XMLConsDataRS!OrgName), "", XMLConsDataRS!OrgName)
            For intLoop = 1 To 5
                If Not IsNull(XMLConsDataRS("AddressLine" & intLoop)) Then
                    If Trim(XMLConsDataRS("AddressLine" & intLoop)) <> "" Then
                        .OrgAddressLine(intLoop) = XMLConsDataRS("AddressLine" & intLoop)
                    End If
                End If
            Next intLoop
            If Not IsNull(XMLConsDataRS!PostCode) Then
                If Trim(XMLConsDataRS!PostCode) <> "" Then
                    .OrgPostCode = XMLConsDataRS!PostCode
                End If
            End If
            .WardKey = IIf(IsNull(XMLConsDataRS!ClinicKey), "", XMLConsDataRS!ClinicKey)
            If InStr(.WardKey, .OrgIdentifier) > 0 Then
                .WardKey = Mid(.WardKey, 7)
            End If
            .WardName = IIf(IsNull(XMLConsDataRS!ClinicName), "", XMLConsDataRS!ClinicName)
            '.WardName = Mid(.WardName, 7)
        End With
        XML_ConsData = True
    End If

    XMLConsDataRS.Close
    blnDoingTrans = False
    Set iceCmd = Nothing
    Set XMLConsDataRS = Nothing
    Exit Function
    
procEH:
    If eClass.Behaviour = -1 Then
        Stop
        Resume
    End If
    eClass.CurrentProcedure = "XML_ReportMsg.XML_ConsData"
    eClass.Add Err.Number, Err.Description, Err.Source
    If blnDoingTrans Then XMLConsDataRS.Close
    Set iceCmd = Nothing
    Set XMLConsDataRS = Nothing
    XML_ConsData = False

End Function

Private Function XML_SenderData(ByVal ReportIndex As Long) As Boolean

    Dim iceCmd As New ADODB.Command
    Dim XMLConsDataRS As New ADODB.Recordset
    Dim blnDoingTrans As Boolean
    Dim intLoop As Integer

    On Error GoTo procEH

    With iceCmd
        .ActiveConnection = iceCon
        .CommandType = adCmdStoredProc
        .CommandText = "ICEMSG_Get_Sender_Details_for_XML"
        .Parameters.Append .CreateParameter("ReportIndex", adInteger, adParamInput, , ReportIndex)
        Set XMLConsDataRS = .Execute
        blnDoingTrans = True
    End With

    If XMLConsDataRS.BOF And XMLConsDataRS.EOF Then
        msgControl.ReportStatus = RS_DATA_INTEGRITY
        msgControl.LogReportMessage RS_DATA_INTEGRITY, "No Sender details available for Report ID provided"
        XML_SenderData = False
    Else
        Do Until XMLConsDataRS.EOF
            With SenderDetails
                Select Case XMLConsDataRS!Type
                    Case "903"
                        .OrgIdentifier = IIf(IsNull(XMLConsDataRS!NationalCode), "", XMLConsDataRS!NationalCode)
                        .OrgName = IIf(IsNull(XMLConsDataRS!OrgName), "", XMLConsDataRS!OrgName)
                        For intLoop = 1 To 5
                            If Not IsNull(XMLConsDataRS("AddressLine" & intLoop)) Then
                                If Trim(XMLConsDataRS("AddressLine" & intLoop)) <> "" Then
                                    .OrgAddressLine(intLoop) = XMLConsDataRS("AddressLine" & intLoop)
                                End If
                            End If
                        Next intLoop
                        If Not IsNull(XMLConsDataRS!PostCode) Then
                            If Trim(XMLConsDataRS!PostCode) <> "" Then
                                .OrgPostCode = XMLConsDataRS!PostCode
                            End If
                        End If
                    Case "906"
                        .Code = IIf(IsNull(XMLConsDataRS!NationalCode), "", XMLConsDataRS!NationalCode)
                        .Name = ""
                        If Not IsNull(XMLConsDataRS!Surname) Then
                            If Trim(XMLConsDataRS!Surname) <> "" Then
                                .Name = XMLConsDataRS!Surname
                            End If
                        End If
                        If Not IsNull(XMLConsDataRS!Forename) Then
                            If Trim(XMLConsDataRS!Forename) <> "" Then
                                If .Name = "" Then
                                    .Name = XMLConsDataRS!Forename
                                Else
                                    .Name = .Name & ", " & XMLConsDataRS!Forename
                                End If
                            End If
                        End If
                End Select
            End With
            XMLConsDataRS.MoveNext
        Loop
        XML_SenderData = True
    End If

    XMLConsDataRS.Close
    blnDoingTrans = False
    Set iceCmd = Nothing
    Set XMLConsDataRS = Nothing
    Exit Function
    
procEH:
    If eClass.Behaviour = -1 Then
        Stop
        Resume
    End If
    eClass.CurrentProcedure = "XML_ReportMsg.XML_SenderData"
    eClass.Add Err.Number, Err.Description, Err.Source
    If blnDoingTrans Then XMLConsDataRS.Close
    Set iceCmd = Nothing
    Set XMLConsDataRS = Nothing
    XML_SenderData = False

End Function

Private Property Let HTMLHeading(sTitle As String)

    sHTMLHeading = "<html><head><title>" & sTitle & "</title><style>body {background-color: #ffffff; font-size: 8pt; font-family: tahoma, verdana, arial;} " & _
                            "input, select, textarea, file, submit, button, span, div, table {font-size: 8pt; font-family: tahoma, verdana;}</style></head>" & _
                            "<body topmargin='0' leftmargin='0' rightmargin='0' bottommargin='0' bgcolor='white'>" & _
                            "<table height='100%' width='100%' cellpadding='0' cellspacing='0'><tr><td height='100%' valign='top' align='center'>"

End Property

Private Property Get HTMLHeading() As String

    HTMLHeading = sHTMLHeading

End Property

Private Property Get HTMLFooting() As String

    HTMLFooting = "</td></tr></table></body></html>"

End Property

Private Sub InitialiseConsultant()

    With ConsultantDetails
        .Code = ""
        .Name = ""
        .Specialty = ""
        .OrgIdentifier = ""
        .OrgName = ""
        ReDim .OrgAddressLine(5) As String
        .OrgPostCode = ""
        .WardKey = ""
        .WardName = ""
    End With

    With SenderDetails
        .Code = ""
        .Name = ""
        .Specialty = ""
        .OrgIdentifier = ""
        .OrgName = ""
        ReDim .OrgAddressLine(5) As String
        .OrgPostCode = ""
        .WardKey = ""
        .WardName = ""
    End With

End Sub

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

Private Sub Class_Terminate()

    If oDataRS.State <> 0 Then oDataRS.Close
    Set oDataRS = Nothing
    Set oXMLMessage = Nothing

End Sub

Public Property Get Notification() As Boolean

    Notification = False

End Property


