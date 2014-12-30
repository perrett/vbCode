VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Edifact3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements AHSLMessage

Private blnHeader As Boolean
Private blnBody As Boolean
Private msgEdi As edifact_v3.MsgMEDRPT
Private ediErr As edifact_v3.ErrorCollection
Private ICEDbCon As ADODB.Connection
Private UDLFILE As String
Private objAHSL As AHSLMsg
Private batchRef As String
Private practiceCode As String
Private lngTotSegs As Long
Private intEdi2Count As Integer
Private intTotReports As Integer
Private intSegs As Integer
'Private i As Integer
'Private j As Integer
Private sg18batt As Integer
Private sg18Seq As Integer
Private dbData As New AHSLMsgDb
Private ahRepList As AHSLReporting.ReportList
Private ahreps As AHSLReporting.Reports
Private ahReport As AHSLReporting.Report
Private txtHandler As New TextSplit
Private cData As String
Private repIndex As Long
Private sCount As Integer
Private LocalUOM As String
Private sg10ftx As Integer
Private sg16ftx As Integer
Private log1_Data As Variant
Private log2_Data As Variant
Private log3_Data As Variant
Private log4_Data As Variant

Private colSamp As New Collection

Private Function AHSLMessage_AddMessage(ReportId As Long, MsgRef As String) As Variant
   On Local Error GoTo procEH
   Dim lastInv As String
   Dim i As Integer
   Dim warnData() As String
   Dim wCnt As Integer
   Dim actCodes As Integer
   Dim sCLS As sampleClass
   
   Set colSamp = Nothing
   Set colSamp = New Collection

   ahReport = ahreps.GetReportByID(ReportId)
   repIndex = ReportId
   msgKey = MsgRef
'   actCodes = 0
   
'  Have all the tests for this message been suppressed?
   
   If dbData.SuppressionTest(CStr(ReportId), 1, log1_Data) Then
      If blnHeader Then
         blnBody = True
         intSegs = 0
         msgEdi.MsgGR.Append
         MapMsgHeaders intEdi2Count
'         Debug.Print "Headers: " & intSegs
         MapSG1 intEdi2Count, ReportId
'         Debug.Print "SG1: " & intSegs
         MapSG2 intEdi2Count, ReportId
'         Debug.Print "SG2: " & intSegs
         MapSG6_SG7 intEdi2Count
'         Debug.Print "SG6/7: " & intSegs

'        Have all the tests for this sample been suppressed?
         If ahreps.HasEntries(ahReport.Samples) Then
            sCount = 1
            For i = 0 To UBound(ahReport.Samples)
               If dbData.SuppressionTest(ahReport.Samples(i).SampleId, 2, log2_Data) Then
                  MapSG6_SG16 intEdi2Count, i
               End If
            Next i
         End If
'         Debug.Print "SG16: " & intSegs
         
         If ahreps.HasEntries(ahReport.Investigations) Then
            lastInv = ""
            sg18Seq = 0
            sg18batt = 0
            wCnt = 0
            ReadCodeError = ""
            For i = 0 To UBound(ahReport.Investigations)
'              Have all the tests for this investigation been suppressed?
               If dbData.SuppressionTest(CStr(ahReport.Investigations(i).Investigation_Index), _
                                         3, _
                                         log3_Data) Then
                  MapSG6_SG18 intEdi2Count, i
               End If
            Next i
         End If
'         Debug.Print "SG18: " & intSegs
         
         intSegs = intSegs + 1   ' add UNT
         msgEdi.MsgGR(intEdi2Count).UNT.MessageReferenceNumber.Value = intEdi2Count + 1
         msgEdi.MsgGR(intEdi2Count).UNT.NumberOfSegmentsInAMessage.Value = intSegs
         intEdi2Count = intEdi2Count + 1
         lngTotSegs = lngTotSegs + intSegs
         
         AHSLMessage_AddMessage = colRC.Count
      End If
   Else
      AHSLMessage_AddMessage = -1
   End If
   Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.AddMessage"
   eClass.FurtherInfo = "Report: " & ahReport.LabNo
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Property Let AHSLMessage_ClinicianDetails(strNewValue As String)
   cData = strNewValue
End Property

Private Property Get AHSLMessage_GetBatchHeader() As String
   AHSLMessage_GetBatchHeader = objAHSL.getHeader("ControlRef")
End Property

Private Sub AHSLMessage_Reports(thisOrg As String)
   Set ahreps = New AHSLReporting.Reports
   ahreps.SetParameters orgId, ""
   orgId = thisOrg
End Sub

Private Property Get AHSLMessage_SuppressedCodes(DataType As Integer) As Variant
   Select Case DataType
      Case 0
         AHSLMessage_SuppressedCodes = log1_Data
         log1_Data = ""
      
      Case 1
         AHSLMessage_SuppressedCodes = log2_Data
         log2_Data = ""
      
      Case 2
         AHSLMessage_SuppressedCodes = log3_Data
         log3_Data = ""
      
      Case 3
         AHSLMessage_SuppressedCodes = log4_Data
         log4_Data = ""
      
   End Select
End Property

Private Sub Class_Initialize()
   eClass.FurtherInfo = "Creating AHSLReporting object"
   Set ahreps = New AHSLReporting.Reports
   Set objAHSL = New AHSLMsg
   Set rcwarn = New ReadCodeSummary
   txtHandler.Invalid = "" '"#$@[\]^_{|}~"
   txtHandler.SubstituteCharacter = "" '"            "
End Sub

Private Sub Class_Terminate()
   Set msgEdi = Nothing
   Set ahreps = Nothing
End Sub

Private Sub AHSLMessage_Reset()
   Set msgEdi = Nothing
   objAHSL.Clear
End Sub

Private Property Get AHSLMessage_msgData() As String
   On Local Error GoTo procEH
   AHSLMessage_msgData = msgEdi.Parse
   Exit Property
   
procEH:
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.msgData"
   If eClass.Add(Err.Number, Err.Description, Err.Source) = -1 Then
      Stop
      Resume
   End If
End Property

Private Property Let AHSLMessage_SetBatchHeader(hdrItem As String, hdrData As String)
   objAHSL.setHeader hdrData, hdrItem
End Property

Private Property Get AHSLMessage_MsgType() As String
   AHSLMessage_MsgType = "Edifact v3"
End Property

Private Property Get AHSLMessage_SegmentTerminator() As String
   AHSLMessage_SegmentTerminator = Right(msgEdi.GetEncodedValue, 1)
End Property

Private Property Get AHSLMessage_ValidationErrors() As AHSLMsg
   On Local Error GoTo procEH
   Dim i As Integer
   Dim strErrPath As String
   Dim strErrDesc As String
   Dim msgErr As edifact_v3.ErrorCollection

   Set msgErr = msgEdi.Validate
   For i = 0 To msgErr.Count - 1
      objAHSL.Add msgErr.Item(i).Description, msgErr.Item(i).Path, Str(i)
   Next i
   Set AHSLMessage_ValidationErrors = objAHSL
   Exit Property

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.ValidationErrors"
   eClass.Add Err.Number, Err.Description, Err.Source
End Property

Private Sub AHSLMessage_NewBatch()
   On Local Error GoTo procEH
   blnHeader = True
   Set msgEdi = Nothing
   Set msgEdi = New edifact_v3.MsgMEDRPT
'  Build UNA & UNB segments
   batchRef = objAHSL.getHeader("ControlRef")
   With msgEdi.UNB
      .SyntaxIdentifier.SyntaxIdentifier.Value = "UNOC"
      .SyntaxIdentifier.SyntaxVersionNumber.Value = "3"
      .InterchangeSender.SenderIdentification.Value = objAHSL.getHeader("SenderID")
      .InterchangeSender.PartnerIdentificationCodeQualifier.Value = objAHSL.getHeader("S_PartnerID")
      .InterchangeRecipient.RecipientIdentification.Value = objAHSL.getHeader("RecipientID")
      .InterchangeRecipient.PartnerIdentificationCodeQualifier.Value = objAHSL.getHeader("R_PartnerID")
      .DateTimeOfPreparation.DateOfPreparation.Value = Format(Now(), "yymmdd")
      .DateTimeOfPreparation.TimeOfPreparation.Value = Format(Now(), "HHMM")
      .InterchangeControlReference.Value = batchRef
      .ApplicationReference.Value = "MEDRPT"
      
      If objAHSL.getHeader("Encrypted") = "Y" Then
         .CommunicationsAgreementId.Value = "P1"
      End If
      
      If objAHSL.getHeader("Acks") Then
         .AcknowledgementRequest.Value = "1"
      End If
      If objAHSL.getHeader("Testing") Then
         .TestIndicator.Value = "1"
      End If
   End With
'   logData = ""
   lngTotSegs = 0
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.NewBatch"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub AHSLMessage_EndBatch()
'   Dim msgedi As edifact_v3.MsgMEDRPT
   On Local Error GoTo procEH
   msgEdi.UNZ.InterchangeControlReference.Value = batchRef
'   msgedi.UNZ.InterchangeControlCount.Value = 0
   msgEdi.UNZ.InterchangeControlCount.Value = intEdi2Count
   intEdi2Count = 0
   lngTotSegs = 0
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.EndBatch"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapMsgHeaders(rpt As Integer)
   On Local Error GoTo procEH
   msgEdi.MsgGR.Item(rpt).UNH.MessageReferenceNumber.Value = rpt + 1
   msgEdi.MsgGR.Item(rpt).UNH.MessageIdentifier.MessageTypeIdentifier.Value = "MEDRPT"
   msgEdi.MsgGR.Item(rpt).BGM.DocumentMessageName.DocumentMessageNameCoded.Value = "LSR"
   msgEdi.MsgGR(rpt).DTM.DateTimePeriod.DateTimePeriodQualifier.Value = "137"
   msgEdi.MsgGR(rpt).DTM.DateTimePeriod.DateTimePeriod.Value = Format(Now(), "yyyymmddHhMm")
   msgEdi.MsgGR(rpt).DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = "203"
   intSegs = 3
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapMsgHeaders"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG1(rpt As Integer, ReportId As Long)
   On Error GoTo procEH
   Dim j As Integer
   Dim pos As Integer
   Dim intHP As Integer
   Dim intOver As Integer
   Dim nCode As String
   Dim lCode As String
   Dim strArray() As String
   
   If ahreps.HasEntries(ahReport.HealthParties) Then
      intHP = -1
      With msgEdi.MsgGR.Item(rpt).SG1
         For j = 0 To UBound(ahReport.HealthParties)
            intHP = intHP + 1
            Select Case ahReport.HealthParties(j).Type
               Case "901"
'                  Message Recipient - GP Practice/Organisation
                  .Append
                  .Item(intHP).NAD.PartyQualifier.Value = "MR"
                  nCode = Left(ahReport.HealthParties(j).NationalCode, 6)
                  lCode = Mid(ahReport.HealthParties(j).NationalCode, 7)
'                  strArray = Split(Trim(ahReport.HealthParties(j).NationalCode))
                  
                  practiceCode = lCode
                  .Item(intHP).NAD.PartyIdentificationDetails.PartyIdIdentification.Value = dbData.HealthPartiesCode(nCode, lCode)
                  .Item(intHP).NAD.PartyIdentificationDetails.CodeListQualifier.Value = "901"
                  .Item(intHP).SPR.ServiceProviderQualifier.Value = "ORG"
'                  .Item(intHp).SPR.SpecialtyDetails.SpecialtyIdentification.Value = ahReport.HealthParties.Specialty
                  intSegs = intSegs + 3
                  Debug.Print "901segs = " & intSegs
               Case "902"
'                  Message Requestor - Professional
'                 Try to determine the name and nat code of the GP if these have not been supplied
                  .Append
                  strArray = Split(cData, "|")
                  If strArray(0) = "" Then
                     cData = dbData.ClinicianDetails(ahReport.HealthParties(j).NationalCode, practiceCode)
                     strArray = Split(cData, "|")
                  End If
'                  If strArray(1) = "" Then
'                     strArray(1) = Trim(Left(ahReport.Clinician, 17))
'                  End If
                  intOver = intHP
                  With .Item(intHP)
                     .NAD.PartyQualifier.Value = "PO"
                     If strArray(0) = "" Then
'                       Clinician goes through "as is"
                        .NAD.PartyName.PartyName1.Value = Trim(ahReport.Clinician)
                        .RFF.Reference.ReferenceQualifier.Value = "AHI"
                        .RFF.Reference.ReferenceNumber.Value = strArray(1) '  Trim(ahReport.Clinician)
                        intSegs = intSegs + 1
'                       .Item(j).NAD.PartyIdentificationDetails.PartyIdIdentification.Value = Trim(ahReport.HealthParties(j).NationalCode)
                     Else
                        If Left(strArray(1), 1) = "!" Then
'                          Clinician nat code not found - send through "as is"
                           .NAD.PartyName.PartyName1.Value = Trim(ahReport.Clinician)
                           .RFF.Reference.ReferenceQualifier.Value = "AHI"
                           .RFF.Reference.ReferenceNumber.Value = Mid(strArray(1), 2)  '  Trim(ahReport.Clinician)
                           intSegs = intSegs + 1
                        Else
                           .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = Left(strArray(1), 17)
                           .NAD.PartyName.PartyName1.Value = strArray(0)
                           .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
                        End If
                     End If
                     .SPR.ServiceProviderQualifier.Value = "PRO"
                     intSegs = intSegs + 3
                  End With
                  
'                 Message Recipient - GP Individual
                  .Append
                  intHP = intHP + 1
                  With .Item(intHP)
                     .NAD.PartyQualifier.Value = "MR"
                     If strArray(0) = "" Then
                        .NAD.PartyName.PartyName1.Value = Trim(ahReport.Clinician)
                        .RFF.Reference.ReferenceQualifier.Value = "AHI"
                        .RFF.Reference.ReferenceNumber.Value = strArray(1) '  Trim(ahReport.Clinician)
                        intSegs = intSegs + 1
'                       .Item(intHp).NAD.PartyIdentificationDetails.PartyIdIdentification.Value = Trim(ahReport.HealthParties(j).NationalCode)
                     Else
'                        strArray = Split(cData, ":")
                        If Left(strArray(1), 1) = "!" Then
'                          Clinician nat code not found - send through "as is"
                           .NAD.PartyName.PartyName1.Value = Trim(ahReport.Clinician)
                           .RFF.Reference.ReferenceQualifier.Value = "AHI"
                           .RFF.Reference.ReferenceNumber.Value = Mid(strArray(1), 2) '  Trim(ahReport.Clinician)
                           intSegs = intSegs + 1
                        Else
                           .NAD.PartyName.PartyName1.Value = strArray(0)
                           .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = Left(strArray(1), 17)
                           .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
                        End If
                     End If
                     .SPR.ServiceProviderQualifier.Value = "PRO"
                     intSegs = intSegs + 3
'                    Reset Cdata for the next report
                     cData = ""
                  Debug.Print "902segs = " & intSegs
                  End With
                  
               Case "903"
'                 Service Provider - Organisation
                  .Append
                  With .Item(intHP)
                     .NAD.PartyQualifier.Value = "SLA"
                     .SPR.ServiceProviderQualifier.Value = "ORG"
                     .NAD.PartyName.PartyName1.Value = dbData.OrganisationId(ahReport.HealthParties(j).NationalCode)
                     intSegs = intSegs + 3
                  End With
                  
'                    Service Provider - Dept
                  .Append
                  intHP = intHP + 1
                  With .Item(intHP)
                     .NAD.PartyQualifier.Value = "SLA"
                     If Trim(ahReport.RepSpecialty) = "" Then
                        .NAD.PartyName.PartyName1.Value = "Pathology"
                     Else
                        .NAD.PartyName.PartyName1.Value = Trim(ahReport.RepSpecialty)
                     End If
                     .SPR.ServiceProviderQualifier.Value = "DPT"
                     intSegs = intSegs + 3
                  End With
                  Debug.Print "903segs = " & intSegs
               
               Case "906"
'                 If a 906 exists overwrite the previuos 902 PO details
'                 Note we do NOT append a new segment

'                 Try to determine the name and nat code of the GP if these have not been supplied
                  cData = dbData.ClinicianDetails(ahReport.HealthParties(j).NationalCode, practiceCode)
                  strArray = Split(cData, "|")
                  If strArray(2) = "" Then
                     strArray(2) = ahReport.Clinician
                  End If

'                 Set intHP to the overwrite value (so the S01 addition does not fail)
                  intHP = intOver

'                 Message Requestor - GP Individual
                  With .Item(intHP)
'                    Clear the previously set values
                     .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = ""
                     .NAD.PartyName.PartyName1.Value = ""
                     .NAD.PartyIdentificationDetails.CodeListQualifier.Value = ""
                     
                     If .RFF.Reference.ReferenceQualifier.Value <> "" Then
                        .RFF.Reference.ReferenceQualifier.Value = ""
                        .RFF.Reference.ReferenceNumber.Value = ""
                        intSegs = intSegs - 1
                     End If
                     
                     .NAD.PartyQualifier.Value = "PO"
                     
                     If strArray(0) = "" Then
'                       Clinician goes through "as is"
                        .NAD.PartyName.PartyName1.Value = strArray(2)   '  Trim(ahReport.Clinician)
                        If .RFF.Reference.ReferenceQualifier.Value = "" Then
                           intSegs = intSegs + 1
                        End If
                        .RFF.Reference.ReferenceQualifier.Value = "AHI"
                        .RFF.Reference.ReferenceNumber.Value = strArray(1) '  Trim(ahReport.Clinician)

                     Else
                        .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = Left(strArray(1), 17)
                        .NAD.PartyName.PartyName1.Value = strArray(2)
                        .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
                     End If
                     .SPR.ServiceProviderQualifier.Value = "PRO"
                  End With
                  Debug.Print "906segs = " & intSegs
            End Select
            .Item(intHP).S01.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "01"
         Next j
      End With
   End If
   cData = ""
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG1"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG2(rpt As Integer, ReportId As Long)
   On Local Error GoTo procEH
   Dim rTime As String
   With msgEdi.MsgGR.Item(rpt).SG2
      
      .S02.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "02"
      .GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
      .RFF.Reference.ReferenceQualifier.Value = "SRI"
      .RFF.Reference.ReferenceNumber.Value = dbData.ReportIdentifier(ReportId)
      
'      If ahReport.Status = "" Then
         .STS.StatusEvent.StatusEventCoded.Value = "UN"
'      Else
'         .STS.StatusEvent.StatusEventCoded.Value = ahReport.Status
'      End If
      intSegs = intSegs + 1
      
      rTime = Trim(ahReport.ReportProducedDate)
      If Len(rTime) <= 10 Then
         .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = "102"
         .DTM.DateTimePeriod.DateTimePeriod.Value = Format(rTime, "yyyymmdd")
      Else
         .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = "203"
         .DTM.DateTimePeriod.DateTimePeriod.Value = Format(rTime, "yyyymmddhHnN")
      End If
      
      .DTM.DateTimePeriod.DateTimePeriodQualifier.Value = "ISR"
   End With
   intSegs = intSegs + 4
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG2"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG6_SG7(rpt As Integer)
   On Local Error GoTo procEH
   Dim i As Integer
   Dim j As Integer
   Dim segTot As Integer
   Dim strTemp As String
   Dim strRepAdr(5) As String
   Dim strMsgAdr(5) As String
   Dim strArray() As String

   With msgEdi.MsgGR.Item(rpt)
      segTot = 0
      If objAHSL.getHeader("Anonymize") Then
         With .SG2.SG6.ADR.AddressDetails
            .AddressFormatCoded.Value = "US"
            .AddressComponent1.Value = "Address Line 1"
            .AddressComponent2.Value = "Address Line 2"
            .AddressComponent3.Value = "Address Line 3"
            .AddressComponent4.Value = "Address Line 4"
            .AddressComponent5.Value = "Address Line 5"
         End With
         .SG2.SG6.ADR.PostcodeIdentification.Value = "PO5T0DE"
         segTot = 1
      Else
         strRepAdr(0) = ahReport.Patient.Address.Line1
         strRepAdr(1) = ahReport.Patient.Address.Line2
         strRepAdr(2) = ahReport.Patient.Address.Line3
         strRepAdr(3) = ahReport.Patient.Address.Line4
         strRepAdr(4) = ahReport.Patient.Address.Line5
         j = 0
         For i = 0 To 4
            If Trim(strRepAdr(i)) <> "" Then
               ReDim Preserve strArray(j)
               strArray(j) = Trim(strRepAdr(i))
               j = j + 1
            End If
         Next i
         If j > 0 Then
            i = UBound(strArray)
            With .SG2.SG6.ADR.AddressDetails
               .AddressFormatCoded.Value = "US"
               .AddressComponent1.Value = strArray(0)
               If i > 0 Then
                  .AddressComponent2.Value = strArray(1)
               End If
               If i > 1 Then
                  .AddressComponent3.Value = strArray(2)
               End If
               If i > 2 Then
                  .AddressComponent4.Value = strArray(3)
               End If
               If i > 3 Then
                  .AddressComponent5.Value = strArray(4)
               End If
            End With
            segTot = 1
         End If
         
         If ahReport.Patient.Address.PostCode <> "" Then
            .SG2.SG6.ADR.PostcodeIdentification.Value = ahReport.Patient.Address.PostCode
            segTot = 1
         End If
      End If
      intSegs = intSegs + segTot
      
      With .SG2.SG6
         .S06.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "06"
         intSegs = intSegs + 1
         If objAHSL.getHeader("Anonymize") Then
            .SG7.PNA.IdentificationNumber.IdentityNumber.Value = "9999999999"
            .SG7.PNA.IdentificationNumber.IdentityNumberQualifier.Value = "OPI"
         Else
            
'           Do we have a valid NHS number?
            If ahReport.Patient.NHSNo <> "" Then
               .SG7.PNA.IdentificationNumber.IdentityNumber.Value = ahReport.Patient.NHSNo
               .SG7.PNA.IdentificationNumber.IdentityNumberQualifier.Value = "OPI"
            ElseIf ahReport.Patient.HospNo <> "" And Left(ahReport.Patient.HospNo, 3) <> "ICE" Then
   '           No NHS number but there is a valid hospital number
               .RFF.Reference.ReferenceQualifier.Value = "SSI"
               .RFF.Reference.ReferenceNumber.Value = ahReport.Patient.HospNo
               intSegs = intSegs + 1
            End If
         End If
      End With
   
      .SG2.SG6.SG7.S07.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "07"
      intSegs = intSegs + 1
      With .SG2.SG6.SG7.PNA
         .PartyQualifier.Value = "PAT"
         intSegs = intSegs + 1
         .NameComponentDetails1.NameComponentQualifier.Value = "SU"
         If objAHSL.getHeader("Anonymize") Then
            .NameComponentDetails1.NameComponent.Value = "TESTPATIENT"
            .NameComponentDetails2.NameComponentQualifier.Value = "FO"
            .NameComponentDetails2.NameComponent.Value = "EDI"
         Else
            .NameComponentDetails1.NameComponent.Value = ahReport.Patient.Surname
            If ahReport.Patient.Forename <> "" Then
               .NameComponentDetails2.NameComponent.Value = ahReport.Patient.Forename
               .NameComponentDetails2.NameComponentQualifier.Value = "FO"
            End If
            
            If ahReport.Patient.MidName <> "" Then
               .NameComponentDetails3.NameComponent.Value = ahReport.Patient.MidName
               .NameComponentDetails3.NameComponentQualifier.Value = "MI"
            End If
         End If
      End With

      With .SG2.SG6.SG7
         If objAHSL.getHeader("Anonymize") Then
            .DTM.DateTimePeriod.DateTimePeriodQualifier.Value = "329"
            .DTM.DateTimePeriod.DateTimePeriod.Value = "19291225"
            .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = "102"
            intSegs = intSegs + 1
         Else
            If ahReport.Patient.dob <> "" Then
               If ahReport.Patient.HospNo = "" Then
                  strTemp = dbData.PatientDOB(ahReport.Patient.NHSNo)
               Else
                  strTemp = dbData.PatientDOB(orgId & ahReport.Patient.HospNo)
               End If
               
               .DTM.DateTimePeriod.DateTimePeriodQualifier.Value = "329"
               strArray = Split(strTemp, "|")
               If strArray(0) = "" Then
                  .DTM.DateTimePeriod.DateTimePeriod.Value = ahReport.Patient.dob
                  .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = "102"
               Else
                  .DTM.DateTimePeriod.DateTimePeriod.Value = strArray(0)
                  .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = strArray(1)
               End If
               intSegs = intSegs + 1
            End If
         End If
         
         With .PDI.SexCoded
            If ahReport.Patient.Sex = "Male" Then
               .Value = "1"
            ElseIf ahReport.Patient.Sex = "Female" Then
               .Value = "2"
            ElseIf ahReport.Patient.Sex = "Unknown" Then
               .Value = "0"
            Else
               .Value = "9"
            End If
         End With
         intSegs = intSegs + 1
      End With
   End With
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG6_SG7"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG10(invRpt As Integer, vTxt As ClinicalText)
   On Error GoTo procEH
   Dim i As Integer
   Dim j As Integer
   Dim cmtCnt As Integer
   With msgEdi.MsgGR.Item(invRpt).SG2.SG6.SG10
'     Have we recorded these in the segment total?
      If .CIN.ClinicalInformationQualifier.Value = "" Then
'        No, so add the two header segmements
         .S10.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "10"
         .CIN.ClinicalInformationQualifier.Value = "UN"
         intSegs = intSegs + 2
      End If
      intSegs = intSegs + txtHandler.FreeText(msgEdi.MsgGR.Item(invRpt).SG2.SG6.SG10, _
                                              vTxt, _
                                              "CID", _
                                              0, _
                                              "NHS003")

'      .FTX.Append
'      .FTX.Item(sg10ftx).TextSubjectQualifier.Value = "CID"
'      For j = 0 To UBound(vTxt)
'         Select Case j
'            Case 0
'               .FTX.Item(sg10ftx).TextLiteral.FreeText1.Value = vTxt(j)
'            Case 1
'               .FTX.Item(sg10ftx).TextLiteral.FreeText2.Value = vTxt(j)
'            Case 2
'               .FTX.Item(sg10ftx).TextLiteral.FreeText3.Value = vTxt(j)
'            Case 3
'               .FTX.Item(sg10ftx).TextLiteral.FreeText4.Value = vTxt(j)
'            Case 4
'               .FTX.Item(sg10ftx).TextLiteral.FreeText5.Value = vTxt(j)
'
'         End Select
'      Next j
'      cmtCnt = cmtCnt + 1
'      intSegs = intSegs + 1
'      sg10ftx = sg10ftx + 1
   End With
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG10"
   eClass.FurtherInfo = "Investigation " & invRpt
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG6_SG16(pRpt As Integer, sRpt As Integer)
   On Local Error GoTo procEH
   Dim vTxt As Variant
   Dim i As Integer
   Dim j As Integer
   Dim cmtCnt As Integer
   Dim sCLS As sampleClass
   Dim rTime As String
   Dim s10Txt As New ClinicalText
   Dim s16Txt As New ClinicalText
   
   On Local Error GoTo NewClass
   Set sCLS = colSamp(ahReport.Samples(sRpt).SampleId & ahReport.Samples(sRpt).SampleCode & ahReport.Samples(sRpt).SampleText)
   sCLS.SampleId = ahReport.Samples(sRpt).SampleCode
   On Local Error GoTo procEH
'   Debug.Print ahReport.Samples(sRpt).SampleId & " repeats = " & pRpt & " + " & sRpt
   msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG16.Append
   With msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG16.Item(sRpt)
      .S16.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "16"
      .SEQ.SequenceInformation.SequenceNumber.Value = sRpt + 1
      .SPC.Append
      .SPC.Item(0).SpecimenCharacteristicsDetails.SpecimenCharacteristic.Value = ahReport.Samples(sRpt).SampleText
      .SPC.Item(0).SpecimenCharacteristicQualifier.Value = "TSP"
      .SPC.Item(0).SpecimenCharacteristicsDetails.CodeListQualifier.Value = ahReport.Samples(sRpt).SampleType
      
      .RFF.Append
      .RFF.Item(0).Reference.ReferenceQualifier.Value = "STI"
      .RFF.Item(0).Reference.ReferenceNumber.Value = RTrim(ahReport.Samples(sRpt).SampleId)
      If ahReport.Samples(sRpt).CollectionAmout <> "" Then
         .QTY.QuantityDetails.QuantityQualifier.Value = "SVO"
         .QTY.QuantityDetails.Quantity.Value = ahReport.Samples(sRpt).CollectionAmout
         .QTY.MeasurementUnitDetails.MeasurementUnitIdentification.Value = ahReport.Samples(sRpt).CollectoinUnits
         intSegs = intSegs + 1
      End If
         
      If Trim(ahReport.Samples(sRpt).CollectionDateTime) <> "" Then
         .DTM.Append
         With .DTM(0).DateTimePeriod
            .DateTimePeriodQualifier.Value = "SCO"
            
            rTime = Trim(ahReport.Samples(sRpt).CollectionDateTime)
            If Len(rTime) <= 10 Then
               .DateTimePeriodFormatQualifier.Value = "102"
               .DateTimePeriod.Value = Format(rTime, "yyyymmdd")
            Else
               .DateTimePeriodFormatQualifier.Value = "203"
               .DateTimePeriod.Value = Format(rTime, "yyyymmddhHnN")
            End If
         End With
         intSegs = intSegs + 5
      End If
      
      If Trim(ahReport.Samples(sRpt).CollectionDateTime_Received) <> "" Then
         .DTM.Append
         With .DTM(.DTM.Count - 1).DateTimePeriod
               .DateTimePeriodQualifier.Value = "SRI"
            
               rTime = Trim(ahReport.Samples(sRpt).CollectionDateTime_Received)
               If Len(rTime) <= 10 Then
                  .DateTimePeriodFormatQualifier.Value = "102"
                  .DateTimePeriod.Value = Format(rTime, "yyyymmdd")
               Else
                  .DateTimePeriodFormatQualifier.Value = "203"
                  .DateTimePeriod.Value = Format(rTime, "yyyymmddhHnN")
               End If
            End With
         intSegs = intSegs + 1
      End If
      
      If sRpt = 0 Then  '  We ony want to report the clinical details once
         If ahreps.HasEntries(ahReport.Comments) Then
            cmtCnt = 0
            sg10ftx = 0
            sg16ftx = 0
            
            
            
            For i = 0 To UBound(ahReport.Comments)
               If ahReport.CommentTypes(i) = "D" Then
                  s10Txt.DataIn = ahReport.Comments(i)
               Else
                  s16Txt.DataIn = ahReport.Comments(i)
               End If
            Next i
            
            If s10Txt.Entries > 0 Then
               MapSG10 pRpt, s10Txt
            End If
            
            If s16Txt.Entries > 0 Then
               intSegs = intSegs + txtHandler.FreeText(msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG16(sRpt), _
                                                       s16Txt, _
                                                       "SPC", _
                                                       0, _
                                                       "NHS003", _
                                                       9)
            End If
            
            
            
'            For i = 0 To UBound(ahReport.Comments)
'               vTxt = txtHandler.DataChunk(ahReport.Comments(i))
'               If VarType(vTxt) > 0 Then
'                  If ahReport.CommentTypes(i) = "D" Then
'                     eClass.FurtherInfo = "SG10.FTX( " & i & ")"
'                     MapSG10 pRpt, vTxt
'                  Else
'                     eClass.FurtherInfo = "SG16.FTX( " & i & ")"
'                     .FTX.Append
'                     .FTX.item(sg16ftx).TextSubjectQualifier.Value = "SPC"
'                     For j = 0 To UBound(vTxt)
'                        Select Case j
'                           Case 0
'                              .FTX.item(sg16ftx).TextLiteral.FreeText1.Value = vTxt(j)
'                           Case 1
'                              .FTX.item(sg16ftx).TextLiteral.FreeText2.Value = vTxt(j)
'                           Case 2
'                              .FTX.item(sg16ftx).TextLiteral.FreeText3.Value = vTxt(j)
'                           Case 3
'                              .FTX.item(sg16ftx).TextLiteral.FreeText4.Value = vTxt(j)
'                           Case 4
'                              .FTX.item(sg16ftx).TextLiteral.FreeText5.Value = vTxt(j)
'                        End Select
'                     Next j
'                     cmtCnt = cmtCnt + 1
'                     intSegs = intSegs + 1
'                     sg16ftx = sg16ftx + 1
'                  End If
'               End If
'            Next i
         End If
      End If
   End With
   Exit Sub
   
NewClass:
   Set sCLS = New sampleClass
   sCLS.SequenceId = sCount
   colSamp.Add sCLS, ahReport.Samples(sRpt).SampleId & ahReport.Samples(sRpt).SampleCode & ahReport.Samples(sRpt).SampleText
   sCount = sCount + 1
   Resume Next

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.FurtherInfo = "Sample No: " & sRpt
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG6_SG16"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG6_SG18(pRpt As Integer, invRpt As Integer)
   On Local Error GoTo procEH
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim intRLen As Integer
   Dim intSPos As Integer
   Dim intComments As Integer
   Dim freeTextType As String
   Dim strText As String
   Dim blnComment As Boolean
   Dim strTemp As String
   Dim strRes As String
   Dim age As String
   Dim vTxt As Variant
   Dim cmtCnt As Integer
   Dim batSeq As Integer
   Dim strArray() As String
   Dim blnStandAlone As Boolean
   Dim sCLS As sampleClass
   Dim strSI As String
   Dim pos As Integer
   Dim ftxType As String
   Dim strRC As String
   Dim rslSeg As Integer
   Dim TestType As Integer
  
   'debug.Print "SG18 = " & intSegs
   If ahReport.Patient.dob <> "" Then
      age = Str(DateDiff("yyyy", ahReport.Patient.dob, Now()))
   Else
      age = "0"
   End If
   Set sCLS = colSamp(dbData.SG18Sample(repIndex, ahReport.Investigations(invRpt).Sample_Index))
   If ahreps.HasEntries(ahReport.Investigations(invRpt).Results) Then
      If (UBound(ahReport.Investigations(invRpt).Results) = 0) And _
         (UCase(ahReport.Investigations(invRpt).Investigation_Requested) = UCase(ahReport.Investigations(invRpt).Results(0).Test) Or _
         Trim(ahReport.Investigations(invRpt).Results(0).Test) = "") Then
         SG18_StandAlone_Test pRpt, invRpt, age, sCLS.SequenceId
         'debug.Print "After StandAlone = " & intSegs
         sg18Seq = sg18Seq + 1
         blnStandAlone = True
      Else
         blnStandAlone = False
         
         SG18_BatteryHeader pRpt, invRpt, age, sCLS.SequenceId
         'debug.Print "After Battery Header = " & intSegs
         sg18Seq = sg18Seq + 1
         
         For i = 0 To UBound(ahReport.Investigations(invRpt).Results)
'           Has the output for this test been supressed?
            If dbData.SuppressionTest(CStr(ahReport.Investigations(invRpt).Results(i).ResultID), _
                                      4, _
                                      log4_Data) Then
'              Set localUOM and change later if necessary
               LocalUOM = ahReport.Investigations(invRpt).Results(i).Units
               msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG18.Append
               With msgEdi.MsgGR.Item(pRpt).SG2.SG6
                  eClass.FurtherInfo = "SG18 Battery component - repeat " & i
                  With .SG18.Item(sg18Seq)
                     .GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
                     .INV.InvestigationCharacteristicQualifier.Value = "MQ"
                     
                     .RFF.Reference.ReferenceQualifier.Value = "ARL"
                     .RFF.Reference.ReferenceNumber.Value = sg18batt + 1
                     intSegs = intSegs + 3
                     If (ahReport.Investigations(invRpt).Results(i).Status = "PR" Or _
                        ahReport.Investigations(invRpt).Results(i).Status = "SR") Then
                        .STS.StatusEvent.StatusEventCoded.Value = ahReport.Investigations(invRpt).Results(i).Status
                        intSegs = intSegs + 1
                     End If
                     
                     If ahreps.HasEntries(ahReport.Investigations(invRpt).Results(i).Comment) Then
                        intComments = UBound(ahReport.Investigations(invRpt).Results(i).Comment)
                     Else
                        intComments = -1
                     End If
                     
                     eClass.FurtherInfo = "SG18.RSL - repeat " & sg18Seq
                     strTemp = txtHandler.AnalyseRangeData(ahReport.Investigations(invRpt).Results(i).Result)
                     strArray = Split(strTemp, "|")
                     strTemp = strArray(0)
                     If UBound(strArray) > 0 Then
                        strSI = strArray(1)
                     Else
                        strSI = ""
                     End If
                     
                     ftxType = "RIT"
                     strRes = dbData.AbnormalCode(ahReport.Investigations(invRpt).Results(i).ResultID)
                     strArray = Split(strRes, "|")
                     
                     rslSeg = 0
                     If IsNumeric(strTemp) And strArray(1) <> "T" Then
                        ftxType = "SPC"
                        .RSL.ResultTypeCoded.Value = "NV"
                        .RSL.ResultDetails1.MeasurementValue.Value = strTemp
                        .RSL.ResultDetails1.MeasurementSignificanceCoded.Value = strSI
                           
'                       If the rubric is not found, method returns the supplied UOM code
                        .RSL.MeasurementUnitDetails.MeasurementUnit.Value = dbData.UOMCode(LocalUOM)
                        rslSeg = 1
                     ElseIf strTemp <> "" Then
                        .FTX.Append
                        .FTX.Item(cmtCnt).TextSubjectQualifier.Value = ftxType
                        .FTX.Item(cmtCnt).TextLiteral.FreeText1.Value = ahReport.Investigations(invRpt).Results(i).Result
                        cmtCnt = cmtCnt + 1
                        intSegs = intSegs + 1
                     End If
                     
                     If ahReport.Investigations(invRpt).Results(i).Abnormal Then
                        If strArray(0) <> "" Then
                           If InStr(1, "HI|LO|OR|PA", strArray(0)) = 0 Then
                              strArray(0) = "OR"
                           End If
                        End If
                        .RSL.ResultNormalcyIndicatorCoded.Value = strArray(0)
                        rslSeg = 1
                     End If
                     
                     intSegs = intSegs + rslSeg
                        
                     eClass.FurtherInfo = "SG18.INV repeat " & sg18Seq
      '              If a read code exists, use it, otherwise use the rubric
                                          
                     strRC = dbData.tMapResults(ahReport.Investigations(invRpt).Results(i).ResultID, _
                                                colRC, _
                                                IsNumeric(strTemp), _
                                                LocalUOM, _
                                                Left(ahReport.Patient.Sex, 1), _
                                                age, _
                                                1, _
                                                True, _
                                                blnIgnoreTestFlags)
                     
'                     strRC = dbData.MapResults(ahReport.Investigations(invRpt).Results(i).Test, _
                                               sCLS.SampleId, _
                                               colRC, _
                                               IsNumeric(strTemp), _
                                               LocalUOM, _
                                               "V2", _
                                               Left(ahReport.Patient.Sex, 1), _
                                               age, _
                                               TestType, _
                                               True)
                     strArray = Split(strRC, "|")
                     With .INV.InvestigationCharacteristicDetails
                        If UBound(strArray) >= 2 Then
                           .InvestigationCharacteristic.Value = strArray(1)
                           .InvestigationCharacteristicIdentification.Value = strArray(0)
                           .CodeListQualifier.Value = strArray(2)
                        Else
                           If blnUpper Then
                              strTemp = UCase(ahReport.Investigations(invRpt).Results(i).Test)
                           Else
                              strTemp = ahReport.Investigations(invRpt).Results(i).Test
                           End If
                           .InvestigationCharacteristic.Value = strTemp
                        End If
                     End With
                        
'                     .INV.InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = txtHandler.EscapedData(ahReport.Investigations(invRpt).Results(i).Test, False)
'                     If strRC <> "" Then
'                        strArray = Split(strRC, "|")
'                        .INV.InvestigationCharacteristicDetails.InvestigationCharacteristicIdentification.Value = strArray(0)
'                        .INV.InvestigationCharacteristicDetails.CodeListQualifier.Value = strArray(2)
'                     End If
                        
                     If intComments >= 0 Then
                        eClass.FurtherInfo = " SG18.FTX repeat " & sg18Seq
                        intSegs = intSegs + txtHandler.FreeText(msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG18.Item(sg18Seq), _
                                                                ahReport.Investigations(invRpt).Results(i), _
                                                                ftxType, _
                                                                cmtCnt)
                     End If
                     cmtCnt = 0
                  End With
   
                  If ftxType = "SPC" Then '  A numeric result
'                    Add any range data
                     With .SG18.Item(sg18Seq).SG20
                        If ahReport.Investigations(invRpt).Results(i).Range <> "" Then
                           eClass.FurtherInfo = "SG20 segment (" & sg18Seq & ")"
                           .Append
                           With .Item(0)
                              .S20.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "20"
                              With .Rnd
                                 .RangeTypeQualifier.Value = "U"
                                 If ahReport.Investigations(invRpt).Results(i).Range <> "" Then
                                    dbData.AdjustRangeData ahReport.Investigations(invRpt).Results(i).Range
                                    .RangeMinimum.Value = dbData.LowRange
                                    .RangeMaximum.Value = dbData.HighRange
                                    intSegs = intSegs + 2
                                 End If
                              End With
                           End With
                        End If
                     End With
                  End If
               End With
               sg18Seq = sg18Seq + 1
               'debug.Print "Battery(" & i & ") = " & intSegs
            End If   '  Suppression test
         Next i
         'debug.Print "After Battery = " & intSegs
         sg18batt = sg18batt + 1
      End If
   ElseIf ahreps.HasEntries(ahReport.Investigations) Then
      SG18_BatteryHeader pRpt, invRpt, age, sCLS.SequenceId, False
      sg18Seq = sg18Seq + 1
   End If
Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG6_SG18"
   eClass.FurtherInfo = "Result  " & invRpt + 1
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub SG18_StandAlone_Test(pRpt As Integer, _
                                 invRpt As Integer, _
                                 age As String, _
                                 seqId As Integer)
   On Local Error GoTo procEH
   Dim strArray() As String
   Dim strTemp As String
   Dim cmtCnt As Integer
   Dim vTxt As Variant
   Dim k As Integer
   Dim strSI As String
   Dim pos As Integer
   Dim ftxType As String
   Dim intRLen As Integer
   Dim intSPos As Integer
   Dim strRC As String
   Dim sCLS As sampleClass
   Dim strRes As String
   Dim rslSeg As Integer
   Dim j As Integer
   Dim TestType As Integer
   
   msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG18.Append
   eClass.FurtherInfo = "Repeat " & sg18Seq
   Set sCLS = colSamp(dbData.SG18Sample(repIndex, ahReport.Investigations(invRpt).Sample_Index))
   With msgEdi.MsgGR.Item(pRpt).SG2.SG6
      .SG18.Item(sg18Seq).GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
      eClass.FurtherInfo = "SG18.INV segment(0)"
      LocalUOM = ahReport.Investigations(invRpt).Results(0).Units
      
      With .SG18.Item(sg18Seq)
         eClass.FurtherInfo = "SG18.RSL repeat " & sg18Seq
         strTemp = txtHandler.AnalyseRangeData(ahReport.Investigations(invRpt).Results(0).Result)
         strArray = Split(strTemp, "|")
         strTemp = strArray(0)
         If UBound(strArray) > 0 Then
            strSI = strArray(1)
         Else
            strSI = ""
         End If
         
         ftxType = "RIT"
         strRes = dbData.AbnormalCode(ahReport.Investigations(invRpt).Results(0).ResultID)
         strArray = Split(strRes, "|")
         
         rslSeg = 0
         If IsNumeric(strTemp) And strArray(1) <> "T" Then
            ftxType = "SPC"
            .RSL.ResultTypeCoded.Value = "NV"
            .RSL.ResultDetails1.MeasurementValue.Value = strTemp
            .RSL.ResultDetails1.MeasurementSignificanceCoded.Value = strSI
                           
'           If the rubric is not found, method returns the supplied UOM code
            .RSL.MeasurementUnitDetails.MeasurementUnit.Value = dbData.UOMCode(LocalUOM)
            rslSeg = 1
         ElseIf strTemp <> "" Then
            .FTX.Append
            .FTX.Item(cmtCnt).TextSubjectQualifier.Value = ftxType
            .FTX.Item(cmtCnt).TextLiteral.FreeText1.Value = ahReport.Investigations(invRpt).Results(0).Result
            cmtCnt = cmtCnt + 1
            intSegs = intSegs + 1
         End If
         
         If ahReport.Investigations(invRpt).Results(0).Abnormal Then
            If strArray(0) <> "" Then
               If InStr(1, "HI|LO|OR|PA", strArray(0)) = 0 Then
                  strArray(0) = "OR"
               End If
            End If
            .RSL.ResultNormalcyIndicatorCoded.Value = strArray(0)
'            .RSL.ResultNormalcyIndicatorCoded.Value = IIf(strArray(0) = "", "OR", strArray(0))
            rslSeg = 1
         End If
                     
         intSegs = intSegs + rslSeg
         
         
'        Evaluate the read code here so we can check for a numeric result
         .INV.InvestigationCharacteristicQualifier.Value = "MQ"
         
'        If a read code exists, use it, otherwise use the rubric
         strRC = dbData.tMapResults(ahReport.Investigations(invRpt).Results(0).ResultID, _
                                    colRC, _
                                    IsNumeric(strTemp), _
                                    LocalUOM, _
                                    Left(ahReport.Patient.Sex, 1), _
                                    age, _
                                    1, _
                                    True, _
                                    blnIgnoreTestFlags)
'         strRC = dbData.MapResults(ahReport.Investigations(invRpt).Results(0).Test, _
                                   sCLS.SampleId, _
                                   colRC, _
                                   IsNumeric(strTemp), _
                                   LocalUOM, _
                                   "V2", _
                                   Left(ahReport.Patient.Sex, 1), _
                                   age, _
                                   TestType, _
                                   True)
         strArray = Split(strRC, "|")
         With .INV.InvestigationCharacteristicDetails
            If UBound(strArray) >= 2 Then
               .InvestigationCharacteristic.Value = strArray(1)
               .InvestigationCharacteristicIdentification.Value = strArray(0)
               .CodeListQualifier.Value = strArray(2)
            Else
               If blnUpper Then
                  strTemp = UCase(ahReport.Investigations(invRpt).Results(0).Test)
               Else
                  strTemp = ahReport.Investigations(invRpt).Results(0).Test
               End If
               .InvestigationCharacteristic.Value = strTemp
            End If
         End With
         
         .RFF.Reference.ReferenceQualifier.Value = "ASL"
         .RFF.Reference.ReferenceNumber.Value = seqId
         intSegs = intSegs + 3   ' GIS segment above
         
         If (ahReport.Investigations(invRpt).Results(0).Status = "PR" Or _
            ahReport.Investigations(invRpt).Results(0).Status = "SR") Then
            .STS.StatusEvent.StatusEventCoded.Value = ahReport.Investigations(invRpt).Results(0).Status
            intSegs = intSegs + 1
         End If

'        Check for any Investigation comments
         If ahreps.HasEntries(ahReport.Investigations(invRpt).Comment) Then
            eClass.FurtherInfo = "Investigation FTX repeat " & sg18Seq
            cmtCnt = cmtCnt + txtHandler.FreeText(msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG18.Item(sg18Seq), _
                                                  ahReport.Investigations(invRpt), _
                                                  "SPC", _
                                                  cmtCnt)
            intSegs = intSegs + cmtCnt
         End If
         
'        Now check for any result comments
         If ahreps.HasEntries(ahReport.Investigations(invRpt).Results(0).Comment) Then
            eClass.FurtherInfo = "Result FTX repeat " & sg18Seq
            intSegs = intSegs + txtHandler.FreeText(msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG18.Item(sg18Seq), _
                                                    ahReport.Investigations(invRpt).Results(0), _
                                                    ftxType, _
                                                    cmtCnt)
         End If
         cmtCnt = 0
      End With
      
      If ftxType = "SPC" Then '  A numeric result
'        Add any range data
         With .SG18.Item(sg18Seq).SG20
            If ahReport.Investigations(invRpt).Results(0).Range <> "" Then
               eClass.FurtherInfo = "SG20 segment (" & sg18Seq & ")"
               .Append
               With .Item(0)
                  .S20.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "20"
                  With .Rnd
                     .RangeTypeQualifier.Value = "U"
                     If ahReport.Investigations(invRpt).Results(0).Range <> "" Then
                        dbData.AdjustRangeData ahReport.Investigations(invRpt).Results(0).Range
                        .RangeMinimum.Value = dbData.LowRange
                        .RangeMaximum.Value = dbData.HighRange
                        intSegs = intSegs + 2
                     End If
                  End With
               End With
            End If
         End With
      End If
   End With
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.Standalone_Test"
   eClass.FurtherInfo = "Result: " & invRpt + 1
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub SG18_BatteryHeader(pRpt As Integer, _
                               invRpt As Integer, _
                               age As String, _
                               seqId As Integer, _
                               Optional SeqReqd As Boolean = True)
   On Local Error GoTo procEH
   Dim strArray() As String
   Dim strTemp As String
   Dim cmtCnt As Integer
   Dim k As Integer
   Dim vTxt As Variant
   Dim sCLS As sampleClass
   Dim strRC As String
   Dim j As Integer
   Dim TestType As Integer
   
   Set sCLS = colSamp(dbData.SG18Sample(repIndex, ahReport.Investigations(invRpt).Sample_Index))
   msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG18.Append
   eClass.FurtherInfo = "Repeat " & sg18Seq
   With msgEdi.MsgGR.Item(pRpt).SG2.SG6
      .SG18.Item(sg18Seq).GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
      eClass.FurtherInfo = "SG18.INV repeat " & sg18Seq
      
      With .SG18.Item(sg18Seq)
         .INV.InvestigationCharacteristicQualifier.Value = "MQ"
         
'        If a read code exists, use it, otherwise use the rubric
'        Changed next line from results(i).test
         strRC = dbData.tMapResults(ahReport.Investigations(invRpt).Investigation_Index, _
                                    colRC, _
                                    IsNumeric(strTemp), _
                                    LocalUOM, _
                                    Left(ahReport.Patient.Sex, 1), _
                                    age, _
                                    2, _
                                    True, _
                                    blnIgnoreTestFlags)
'         strRC = dbData.MapResults(ahReport.Investigations(invRpt).Investigation_Requested, _
                                   sCLS.SampleId, _
                                   colRC, _
                                   IsNumeric(strTemp), _
                                   LocalUOM, _
                                   "V2", _
                                   Left(ahReport.Patient.Sex, 1), _
                                   age, _
                                   TestType, _
                                   True)
         strArray = Split(strRC, "|")
         With .INV.InvestigationCharacteristicDetails
            If UBound(strArray) >= 2 Then
               .InvestigationCharacteristic.Value = strArray(1)
               .InvestigationCharacteristicIdentification.Value = strArray(0)
               .CodeListQualifier.Value = strArray(2)
            Else
               If blnUpper Then
                  strTemp = UCase(ahReport.Investigations(invRpt).Investigation_Requested)
               Else
                  strTemp = ahReport.Investigations(invRpt).Investigation_Requested
               End If
               .InvestigationCharacteristic.Value = strTemp
            End If
         End With
   
         If SeqReqd Then
            .SEQ.SequenceInformation.SequenceNumber.Value = sg18batt + 1
            intSegs = intSegs + 1
         End If
         .RFF.Reference.ReferenceQualifier.Value = "ASL"
         .RFF.Reference.ReferenceNumber.Value = seqId
         intSegs = intSegs + 3   ' GIS segment above
         If ahreps.HasEntries(ahReport.Investigations(invRpt).Comment) Then
            eClass.FurtherInfo = "FTX repeat " & sg18Seq
            intSegs = intSegs + txtHandler.FreeText(msgEdi.MsgGR.Item(pRpt).SG2.SG6.SG18.Item(sg18Seq), _
                                                    ahReport.Investigations(invRpt), _
                                                    "SPC")
         End If
      End With
   End With
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.BatteryHeader"
   eClass.FurtherInfo = "Result: " & invRpt
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
