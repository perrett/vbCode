VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Edifact_NHS002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private blnHeader As Boolean
Private blnBody As Boolean
Private RepIndex As Long
Private sg18Index As Integer
Private srcRS As ADODB.Recordset
Private msgEDI As nhs002.MsgMEDRPT
Private msgErr As nhs002.ErrorCollection
Private strArray() As String
Private ReportClass As Integer

Private Const TF_VALID_RC = &HA1

Private colSamp As New Collection

Public Property Get AllowApostrophe() As Boolean
   AllowApostrophe = True
End Property

Public Function CreateMessage(ReportId As Long) As Variant
   On Local Error GoTo procEH
   Dim msgBuf As String
   Dim I As Integer
   
   Set msgEDI = CreateObject("nhs002.MsgMEDRPT") 'New nhs002.MsgMEDRPT
   With msgEDI.UNH
      .MessageReferenceNumber.Value = msgControl.MessageId
      .MessageIdentifier.MessageTypeIdentifier.Value = "MEDRPT"
   End With
   
   msgEDI.BGM.DocumentMessageName.DocumentMessageNameCoded.Value = "LSR"
   
   With msgEDI.DTM.DateTimePeriod
      .DateTimePeriodQualifier.Value = "137"
      .DateTimePeriod.Value = Format(Now(), "yyyymmddHhMm")
      .DateTimePeriodFormatQualifier.Value = "203"
   End With
   
   msgControl.ReportStatus = 0
   
   MapSG1 ReportId
   MapSG2
   MapSG6_SG7 ReportId
   
   If repListRS!service_Report_Type = "002" Then
      sg18Index = 0
      MapSG18_Discharge ReportId
   Else
      MapSG10 ReportId
      sg18Index = 0
      MapSG6_SG16 ReportId
   End If

'   MapSG10 ReportId
'   sg18Index = 0
'   MapSG6_SG16 ReportId
   
   If msgControl.ReportStatus > 0 Then
      msgBuf = "Errors encountered when generating message. 'See comments for description'"
   Else
      With msgEDI.UNT
         .MessageReferenceNumber.Value = msgControl.MessageId
         .NumberOfSegmentsInAMessage.Value = 0
      End With
   
      Set msgErr = msgEDI.Validate
      If msgErr.Count > 0 Then
         msgControl.MessageStatus = MS_PARSE_FAIL
         msgControl.ReportStatus = MS_PARSE_FAIL
         For I = 0 To msgErr.Count - 1
            msgControl.LogReportMessage MS_PARSE_FAIL, msgControl.CurrentReport & ": " & _
                                                        msgErr(I).Path & " - " & _
                                                        msgErr(I).Description
         Next I
      End If
      
      msgBuf = msgEDI.parse
      
'      msgData.InvalidChars = "#$@[\]^_`{|}~"
'      msgData.ReplaceChars = "             "
      msgData.SegmentCount msgBuf
'      msgData.EscapeData msgBuf
   End If
   Set msgEDI = Nothing
   CreateMessage = msgBuf
   Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.AddMessage"
   eClass.FurtherInfo = "Report: " & ReportId
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function FileHeader() As String
   Dim strHeader As String
   
   strHeader = "UNA:+.? '" & _
                "UNB+UNOB:2+<SENDINGTRADER>:80+<RECEIVINGTRADER>:80+<DATE:TIME>+<INTERCHANGE>++MEDRPT"
   strHeader = Replace(strHeader, "<SENDINGTRADER>", msgControl.CurrentSender)
   strHeader = Replace(strHeader, "<RECEIVINGTRADER>", msgControl.CurrentReceiver)
   strHeader = Replace(strHeader, "<DATE:TIME>", Format(Now(), "YYMMDD:HHNN"))
   strHeader = Replace(strHeader, "<INTERCHANGE>", msgControl.CurrentInterchange)
   
   FileHeader = strHeader & "'" ' & Flags & "'"
End Function

Public Property Get FileTrailer() As String
   FileTrailer = "UNZ+" & msgControl.MsgInBatch & "+" & msgControl.CurrentInterchange & "'"
End Property

Public Property Get InvalidChars() As Variant
   ReDim strArray(2)
   strArray(0) = Chr(10) & Chr(13) & "#$@[\]^_{|}~`"
   strArray(1) = Chr(0) & Chr(0) & "             "
   InvalidChars = strArray
End Property

Public Property Get ReleaseChar() As String
   ReleaseChar = "?"
End Property

Public Property Get RetainGrave() As Boolean
   RetainGrave = False  '  (True And blnRetainGrave)
End Property

Private Property Get msgType() As String
   msgType = "Edifact v3"
End Property

Public Property Get Notification() As Boolean
   Notification = False
End Property

Private Sub MapSG1(ReportIndex As Long)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim clinNatCode As String
   Dim gpCode As String
   Dim gpName As String
   Dim gpNat As Boolean
   Dim hpData As New clsHealthParties
   Dim DownloadInd As Integer
'   Dim gp902Name As String
'   Dim gp906Code As String
'   Dim gp906Name As String
'   Dim nat902 As String
'   Dim blnNat902 As Boolean
   
'   With iceCmd
'      .ActiveConnection = iceCon
'      .CommandType = adCmdStoredProc
'      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportIndex)
'      .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , 0)
'   End With
   
   With msgEDI.SG1
      .Append
      With .Item(0).NAD
         .PartyQualifier.Value = "PO"
         .PartyIdentificationDetails.PartyIdIdentification.Value = repListRS!EDI_Loc_Nat_Code_To
         .PartyIdentificationDetails.CodeListQualifier.Value = "901"
      End With
      
      .Item(0).SPR.ServiceProviderQualifier.Value = "ORG"
                 
      ReportClass = repListRS!report_Class
      hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
      hpData.LTSIndex = repListRS!EDI_LTS_Index
      hpData.DownloadRequest = DownloadInd
      hpData.Read ReportIndex, blnUse906, True, blnClinNameUseAll
      
      If blnUse906 Then
         gpName = hpData.HP906Name
         gpCode = hpData.HP906Code
         gpNat = hpData.HP906Nat
      Else
         gpName = hpData.HP902Name
         gpCode = hpData.HP902Code
         gpNat = hpData.HP902Nat
      End If
      
      .Append
      
      With .Item(1)
         .NAD.PartyQualifier.Value = "PO"
         .NAD.PartyName.PartyName1.Value = gpName
         
         If gpNat Then
            .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = gpCode
            If Left(gpCode, 1) = "C" Then
               .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "902"
            Else
               .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
            End If
         Else
            .RFF.Append
            .RFF(0).Reference.ReferenceQualifier.Value = "AHI"
            
            If gpCode = "" Or gpCode = "~" Then
               gpCode = "UNK"
            End If
            
            .RFF(0).Reference.ReferenceNumber.Value = gpCode
         End If
         .SPR.ServiceProviderQualifier.Value = "PRO"
      End With
                  
      .Append
      With .Item(2)
         .NAD.PartyQualifier.Value = "SLA"
         .SPR.ServiceProviderQualifier.Value = "ORG"
         .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = msgControl.OrganisationCode
         .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "903"
      End With
                  
      .Append
      With .Item(3)
         .NAD.PartyQualifier.Value = "SLA"
         If repListRS!Specialty = "" Then
            .NAD.PartyName.PartyName1.Value = "Pathology"
         Else
            .NAD.PartyName.PartyName1.Value = repListRS!Specialty
         End If
         .SPR.ServiceProviderQualifier.Value = "DPT"
      End With
   End With
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS002.MapSG1"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG2()
   On Local Error GoTo procEH
   Dim rTime As String
   
   With msgEDI.SG2
      .S02.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "02"
      .GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
      .RFF.Reference.ReferenceQualifier.Value = "SRI"
      If IsNull(repListRS!Report_Identifier) Then
         .RFF.Reference.ReferenceNumber.Value = repListRS!Service_Id
      Else
         .RFF.Reference.ReferenceNumber.Value = repListRS!Report_Identifier  ' Trim(Left(repListRS!Service_ID, InStr(1, repListRS!Service_ID, " ")))
      End If
      
      .STS.StatusEvent.StatusEventCoded.Value = "UN"
      
      .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = "203"
      .DTM.DateTimePeriod.DateTimePeriod.Value = Format(repListRS!Report_Date, "yyyymmddhHnN")
      
      .DTM.DateTimePeriod.DateTimePeriodQualifier.Value = "ISR"
   End With
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHS002.MapSG2"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG6_SG7(RepId As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   
   Dim I As Integer
   Dim j As Integer
   Dim segTot As Integer
   Dim strTemp As String
   Dim strRepAdr(5) As String
   Dim strMsgAdr(5) As String
   Dim strArray() As String
   Dim fmtStr As String
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Patient"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RepId)
      Set RS = .Execute
   End With
   With msgEDI.SG2
      With .SG6
         If msgControl.AnonymizeData Then
            With .ADR.AddressDetails
               .AddressFormatCoded.Value = "US"
               .AddressComponent1.Value = "Address Line 1"
               .AddressComponent2.Value = "Address Line 2"
               .AddressComponent3.Value = "Address Line 3"
               .AddressComponent4.Value = "Address Line 4"
               .AddressComponent5.Value = "Address Line 5"
            End With
            .ADR.PostcodeIdentification.Value = "PO5T0DE"
         Else
               
            strRepAdr(0) = RS!Pat_Addr_Line1 & ""
            strRepAdr(1) = RS!Pat_Addr_Line2 & ""
            strRepAdr(2) = RS!Pat_Addr_Line3 & ""
            strRepAdr(3) = RS!Pat_Addr_Line4 & ""
            strRepAdr(4) = RS!Pat_Addr_Line5 & ""
            j = 0
            
            For I = 0 To 4
               strRepAdr(I) = Replace(strRepAdr(I), "'", "`")
               If Trim(Replace(strRepAdr(I), "~", "")) <> "" Then
                  ReDim Preserve strArray(j)
                  If Len(strRepAdr(I)) < 35 Then
                     strArray(j) = Trim(strRepAdr(I)) & Chr(255)
                  Else
                     strArray(j) = Left(strRepAdr(I), 34) & Chr(255)
                  End If
                  j = j + 1
               End If
            Next I
               
            With .ADR.AddressDetails
               .AddressFormatCoded.Value = "US"
               If j > 0 Then
                  I = UBound(strArray)
                  .AddressComponent1.Value = strArray(0)
                  If I > 0 Then
                     .AddressComponent2.Value = strArray(1)
                  End If
                  If I > 1 Then
                     .AddressComponent3.Value = strArray(2)
                  End If
                  If I > 2 Then
                     .AddressComponent4.Value = strArray(3)
                  End If
                  If I > 3 Then
                     .AddressComponent5.Value = strArray(4)
                  End If
               Else
                  .AddressComponent1.Value = "Address Unknown"
               End If
            End With
             
            If RS!Pat_PostCode <> "" Then
               .ADR.PostcodeIdentification.Value = RS!Pat_PostCode
            End If
         End If
      End With
      
      With .SG6
         .S06.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "06"
         .SG7.Append
         
         If msgControl.AnonymizeData Then
            .SG7(0).PNA.IdentificationNumber.IdentityNumber.Value = "9999999999"
            .SG7(0).PNA.IdentificationNumber.IdentityNumberQualifier.Value = "OPI"
         Else
            
'           Do we have a valid NHS number?
            If RS!New_NHS_No <> "" Then
               .SG7(0).PNA.IdentificationNumber.IdentityNumber.Value = RS!New_NHS_No
               .SG7(0).PNA.IdentificationNumber.IdentityNumberQualifier.Value = "OPI"
            End If

'           Do we have a valid hospital number?
            If RS!Hospital_Number <> "" And Left(RS!Hospital_Number, 3) <> "ICE" Then
               .RFF.Append
               .RFF(0).Reference.ReferenceQualifier.Value = "SSI"
               .RFF(0).Reference.ReferenceNumber.Value = RS!Hospital_Number
            End If
         End If
      End With
      
      With .SG6.SG7(0)
         .S07.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "07"
         With .PNA
            .PartyQualifier.Value = "PAT"
            .NameTypeCoded.Value = "US"
            .NameComponentDetails1.NameComponentQualifier.Value = "SU"
            If msgControl.AnonymizeData Then
               .NameComponentDetails1.NameComponent.Value = "TESTPATIENT"
               .NameComponentDetails2.NameComponentQualifier.Value = "FO"
               .NameComponentDetails2.NameComponent.Value = "EDI"
            Else
               .NameComponentDetails1.NameComponent.Value = Replace(RS!Surname, "'", "`")
               If RS!Forename <> "" Then
                  .NameComponentDetails2.NameComponent.Value = Replace(RS!Forename, "'", "`")
                  .NameComponentDetails2.NameComponentQualifier.Value = "FO"
               End If
               
               If RS!Midname <> "" Then
                  .NameComponentDetails3.NameComponent.Value = Replace(RS!Midname, "'", "`")
                  .NameComponentDetails3.NameComponentQualifier.Value = "MI"
               End If
            End If
         End With
         
         If msgControl.AnonymizeData Then
            .DTM.DateTimePeriod.DateTimePeriodQualifier.Value = "329"
            .DTM.DateTimePeriod.DateTimePeriod.Value = "19000101"
            .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = "102"
         Else
            If RS!Date_Of_Birth <> "" Then
               .DTM.DateTimePeriod.DateTimePeriodQualifier.Value = "329"
               
               If RS!Date_Of_Birth_Format = "602" Then
                  fmtStr = "yyyy"
               ElseIf RS!Date_Of_Birth_Format = "610" Then
                  fmtStr = "yyyymm"
               Else
                  fmtStr = "yyyymmdd"
               End If
               
               .DTM.DateTimePeriod.DateTimePeriod.Value = Format(RS!Date_Of_Birth, fmtStr)
               .DTM.DateTimePeriod.DateTimePeriodFormatQualifier.Value = IIf(IsNull(RS!Date_Of_Birth_Format), "102", RS!Date_Of_Birth_Format)
            End If
         End If
         
         If IsNull(RS!Sex) Then
            .PDI.SexCoded.Value = 0
         Else
            .PDI.SexCoded.Value = Val(RS!Sex)
         End If
      End With
   End With
   RS.Close
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHS002.MapSG6_SG7"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG10(ReportId As Long)
   On Error GoTo procEH
   Dim I As Integer
   Dim j As Integer
   Dim cmtCnt As Integer
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Comments"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   RS.Filter = "Comment_Type = 'D'"
   If RS.RecordCount > 0 Then
      With msgEDI.SG2.SG6
         .SG10.Append
         .SG10(0).S10.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "10"
         .SG10(0).CIN.ClinicalInformationQualifier.Value = "UN"
         msgData.FreeText .SG10(0), _
                          RS, _
                          "CID", _
                          0, _
                          "NHS002"
      End With
   End If
   RS.Close
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHS002.MapSG10"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG6_SG16(ReportId As Long)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim sampSeq As Integer
   Dim RS As ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim NoPrevComments As Integer
   Dim lastIndex As Long
   Dim battSeq As Long
   Dim sampleIdCode As String
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Sample"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   sampSeq = 0
   battSeq = 1
   
   Do Until RS.EOF
      msgEDI.SG2.SG6.SG16.Append
      With msgEDI.SG2.SG6.SG16.Item(sampSeq)
         .S16.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "16"
         .SEQ.SequenceInformation.SequenceNumber.Value = sampSeq + 1
         .SPC.Append
         
         With .SPC.Item(0)
            sampleIdCode = Replace(RS!Sample_Code, "'", "")
'            If repListRS!Report_Class = 2 Then
'               If sampleIdCode = "" Then
'                  sampleIdCode = "T027"
'               End If
'            End If
                  
            .SpecimenCharacteristicsDetails.SpecimenCharacteristicIdentification.Value = sampleIdCode
            .SpecimenCharacteristicQualifier.Value = "TSP"
            .SpecimenCharacteristicsDetails.CodeListQualifier.Value = "920"
         End With
         
         If Trim(RS!AnatOrig_Code & "") <> "" Then
            .SPC.Append
            With .SPC.Item(.SPC.Count - 1)
               .SpecimenCharacteristicQualifier.Value = "ALO"
               .SpecimenCharacteristicsDetails.SpecimenCharacteristicIdentification.Value = RS!AnatOrig_Code
               .SpecimenCharacteristicsDetails.SpecimenCharacteristic.Value = Trim(Left(RS!Sample_Text, 35))
               .SpecimenCharacteristicsDetails.CodeListQualifier.Value = "921"
            End With
         End If
         
         If Trim(RS!Collection_Code & "") <> "" Then
            With .PRC
               .ProcessTypeAndDescription.ProcessTypeIdentification.Value = "SCO"
               With .ProcessIdentificationDetails
                  .ProcessIdentification.Value = RS!Collection_Code
                  .CodeListQualifier.Value = "922"
                  .Process.Value = RS!Collection_Text
               End With
            End With
         End If
         
         .RFF.Append
         .RFF.Item(0).Reference.ReferenceQualifier.Value = "STI"
         .RFF.Item(0).Reference.ReferenceNumber.Value = RTrim(RS!Sample_Id)
         
         If RS!Collection_Amount <> "" Then
            .QTY.QuantityDetails.QuantityQualifier.Value = "SVO"
            .QTY.QuantityDetails.Quantity.Value = RS!Collection_Amount
            .QTY.MeasurementUnitDetails.MeasurementUnitIdentification.Value = RS!Collection_Units
         End If
            
         If RS!Collection_DateTime <> "" Then
            .DTM.Append
            With .DTM(0).DateTimePeriod
               .DateTimePeriodQualifier.Value = "SCO"
               If RS!ctime = 0 Then
                  .DateTimePeriodFormatQualifier.Value = "102"
                  .DateTimePeriod.Value = Format(RS!Collection_DateTime, "yyyymmdd")
               Else
                  .DateTimePeriodFormatQualifier.Value = "203"
                  .DateTimePeriod.Value = Format(RS!Collection_DateTime, "yyyymmddhHnN")
               End If
            End With
         End If
         
         If Trim(RS!Collection_DateTimeReceived) <> "" Then
            .DTM.Append
            With .DTM(.DTM.Count - 1).DateTimePeriod
               .DateTimePeriodQualifier.Value = "SRI"
               If RS!rTime = 0 Then
                  .DateTimePeriodFormatQualifier.Value = "102"
                  .DateTimePeriod.Value = Format(RS!Collection_DateTimeReceived, "yyyymmdd")
               Else
                  .DateTimePeriodFormatQualifier.Value = "203"
                  .DateTimePeriod.Value = Format(RS!Collection_DateTimeReceived, "yyyymmddhHnN")
               End If
            End With
         End If
         
         With iceCmd
            .CommandText = "ICELABCOMM_Report_Comments"
            Set RS2 = .Execute
         End With
         
         RS2.Filter = "Comment_Type = 'X'"
         NoPrevComments = RS2.RecordCount
         
         If (sampSeq = 0) And (RS2.RecordCount > 0) Then
            msgData.FreeText msgEDI.SG2.SG6.SG16(sampSeq), _
                             RS2, _
                             "SPC", _
                             0, _
                             "NHS003", _
                             9
         End If
         
         RS2.Filter = "Comment_Type = 'R'"
         If RS2.RecordCount > 0 Then
            msgData.FreeText msgEDI.SG2.SG6.SG16(sampSeq), _
                             RS2, _
                             "SPC", _
                             NoPrevComments, _
                             "NHS002", _
                             9
         End If
         RS2.Close
      End With
      
      With iceCmd
         If blnUseRCIndex Then
            .CommandText = "ICELABCOMM_Report_Invest_By_Index"
         Else
            .CommandText = "ICELABCOMM_Report_Invest_By_Code"
         End If
         
         .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
         .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , RS!Sample_Index)
         .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
         Set RS2 = .Execute
         .Parameters.Delete ("Abnormal")
         .Parameters.Delete ("SampId")
         .Parameters.Delete ("LTSIndex")
      End With
      
      lastIndex = sg18Index
      
      Do Until RS2.EOF
         If IsNull(RS2!Result_Index) Or RS2!Result_Recs > 1 Then
            SG18_Battery RS2, sampSeq + 1, battSeq
            battSeq = battSeq + 1
         Else
            SG18_StandAlone RS2, sampSeq + 1, battSeq
         End If
         RS2.MoveNext
      Loop
      
      RS2.Close
      
      If sg18Index > lastIndex Then
         sampSeq = sampSeq + 1
      Else
         msgControl.LogReportMessage IS_SAMPLE_SUPPRESSED, RS!Sample_Text & _
                                     "- All tests have output suppressed"
         msgEDI.SG2.SG6.SG16(sampSeq).SetEmpty
      End If
      RS.MoveNext
   Loop
   
   If RS.RecordCount = 0 Then
      msgControl.ReportStatus = RS_DATA_INTEGRITY
      msgControl.LogReportMessage RS_DATA_INTEGRITY, "No samples recorded against this report"
   ElseIf sampSeq = 0 Then
      msgControl.ReportStatus = MS_NO_OUTPUT
      msgControl.LogReportMessage RS_SUPPRESSION, "No tests/results to ouput." & _
                                  " Check for any suppression messages"
   End If
   
   RS.Close
   Set RS2 = Nothing
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS002.MapSG6_SG16"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub SG18_Test(invRS As ADODB.Recordset, _
                      RelatedBattery As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim resRS As New ADODB.Recordset
   Dim RS As New ADODB.Recordset
   Dim testCount As Integer
   Dim localUOM As String
   Dim strTemp As String
   Dim strSi As String
   Dim ftxType As String
   Dim cmtCnt As Integer
   Dim Status As enumInvResStatus
   Dim vData As Variant
   Dim vFields As Variant
   Dim testDesc As String
   Dim rcIndex As Long
   Dim rPosn As Integer
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      
      If blnUseRCIndex Then
         .CommandText = "ICELABCOMM_Report_Results_By_Index"
      Else
         .CommandText = "ICELABCOMM_Report_Results_By_Code"
      End If
      
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , invRS!Investigation_Index)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      Set resRS = .Execute
      .Parameters.Delete ("LTSIndex")
   End With
   resRS.Filter = "EDI_OP_Suppress = False"
   testCount = resRS.RecordCount
   
   If testCount = 0 Then
      Status = IS_INV_SUPPRESSED
      msgControl.LogReportMessage Status, invRS!Investigation_Code & " (" & invRS!Investigation_Requested & _
                                  ") ALL the tests in this battery have their output suppressed"
   Else
      resRS.Filter = ""
      
      resRS.MoveFirst
      
      sg18Index = sg18Index + 1
      testCount = 1
      cmtCnt = 0
      Do Until resRS.EOF
         If IsNull(resRS!EDI_RC_Index) Then
'         If IsNull(vData(12, rPosn)) Then
            rcIndex = -1
         Else
            rcIndex = resRS!EDI_RC_Index
         End If
         
         Status = 0
   
         If IsNull(resRS!EDI_Local_Test_Code) Then
'         If IsNull(vData(0, rPosn)) Then
            testDesc = resRS!Result_Code & " (" & resRS!Result_Rubric & ") "
         Else
            testDesc = Replace(resRS!EDI_Local_Test_Code, Chr(0), "") & " (" & resRS!EDI_Local_Rubric & ") "
         End If
   
         testDesc = Replace(testDesc, "'", "`")
         
'        Has the output for this test been supressed?
         If resRS!EDI_OP_Suppress Then
'         If vData(2, rPosn) Then
            Status = IS_TEST_SUPPRESSED
            msgControl.LogReportMessage Status, testDesc & "Test suppressed"
         Else
            localUOM = resRS!UOM_Code
            msgEDI.SG2.SG6.SG18.Append
            With msgEDI.SG2.SG6
               With .SG18(sg18Index)
                  .GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
                  .INV.Append
                  .INV(0).InvestigationCharacteristicQualifier.Value = "MQ"
                  
                  .RFF.Append
                  .RFF(0).Reference.ReferenceQualifier.Value = "ARL"
                  .RFF(0).Reference.ReferenceNumber.Value = RelatedBattery
                  
                  strTemp = msgData.AnalyseResultString(resRS!Result)
                  strArray = Split(strTemp, "|")
                  strTemp = strArray(0)
                  
                  If UBound(strArray) > 0 Then
                     strSi = strArray(1)
                  Else
                     strSi = ""
                  End If
                  
                  If (resRS!Status = "PR" Or resRS!Status = "FR" Or resRS!Status = "NA") Then
                     .STS.StatusEvent.StatusEventCoded.Value = resRS!Status
                  End If
                           
                  strTemp = msgData.AnalyseResultString(resRS!Result)
                  strArray = Split(strTemp, "|")
                  strTemp = strArray(0)
                  
                  If UBound(strArray) > 0 Then
                     strSi = strArray(1)
                  Else
                     strSi = ""
                  End If
                  
                  ftxType = "RIT"
                           
                  If IsNumeric(strTemp) And resRS!Result_Type <> "T" Then
                     ftxType = "SPC"
                     .RSL.ResultTypeCoded.Value = "NV"
                     .RSL.ResultDetails1.MeasurementValue.Value = strTemp
                     .RSL.ResultDetails1.MeasurementSignificanceCoded.Value = strSi
                     
                     If IsNull(resRS!EDI_Nat_UOM) Then
                        .RSL.MeasurementUnitDetails.MeasurementUnit.Value = resRS!UOM_Code
                     Else
                        .RSL.MeasurementUnitDetails.MeasurementUnitIdentification.Value = resRS!EDI_Nat_UOM
                        .RSL.MeasurementUnitDetails.CodeListQualifier.Value = "923"
                     End If
                     
                  ElseIf strTemp <> "" Then
                     .FTX.Append
                     .FTX.Item(.FTX.Count - 1).TextSubjectQualifier.Value = ftxType
                     .FTX.Item(.FTX.Count - 1).TextLiteral.FreeText1.Value = Replace(resRS!Result, "'", "`")
                  End If
                           
                  If resRS!Abnormal_Flag Then
                     If Trim(resRS!Abnormal_Code & "") <> "" Then
                        If InStr(1, "NM|UN|LO|SL|HI|SH", Trim(resRS!Abnormal_Code & "")) = 0 Then
                           If Trim(resRS!Abnormal_Code & "") = "PA" Then
                              .RSL.ResultNormalcyIndicatorCoded.Value = ""
                           Else
                              .RSL.ResultNormalcyIndicatorCoded.Value = "UN"
                           End If
                        Else
                           .RSL.ResultNormalcyIndicatorCoded.Value = resRS!Abnormal_Code
                        End If
                     End If
                  End If
                                                
                  .INV(0).InvestigationCharacteristicQualifier.Value = "MQ"
                  
                  With .INV(0).InvestigationCharacteristicDetails
                     If rcIndex = 0 Then
                        Status = IS_RC_NONE
                        If blnReadCodeInfo Then
                           msgControl.LogReportMessage Status, "Read Code suppressed by Lab system."
                        End If
                     Else
                        If IsNull(resRS!Read_V2RC) Then
'                        If IsNull(vData(5, rPosn)) Then
                           Status = IS_RC_NONE
                           If blnReadCodeInfo Then
                              msgControl.LogReportMessage IS_RC_NONE, testDesc & "Not Read-Coded"
                           End If
                        Else
                           If resRS!Read_Status = "D" Then
'                           If vData(9, rPosn) = "D" Then
                              Status = IS_RC_DELETED
                              msgControl.LogReportMessage IS_RC_DELETED, testDesc & resRS!Read_Comments
                           Else
                              If resRS!Read_Test = "F" Then
'                              If vData(8, rPosn) = "F" Then
                                 Status = IS_RC_NA
                                 msgControl.LogReportMessage IS_RC_NA, testDesc & "Read code flagged as inapplicable t a test in the bounded list."
                              End If
                              
                              If resRS!EDI_OP_Active = False Then
'                              If vData(3, rPosn) = False Then
                                 Status = Status Or IS_TEST_INACTIVE
                                 msgControl.LogReportMessage Status, testDesc & "Test not read-coded - marked as inactive."
                              End If
                           
                              If resRS!Read_Ratio <> "T" Then
'                              If vData(10, rPosn) <> "T" Then
                                 If IsNumeric(resRS!Result) Then
                                    If Trim(resRS!UOM_Code) = "" Then
                                       If Trim(resRS!EDI_Nat_UOM & "") = "" Then
                                          Status = Status Or IS_RC_REMOVED
                                          msgControl.LogReportMessage IS_RC_REMOVED, testDesc & "The bounded list specifies a UOM is required with this read code - None present."
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                     
                     If (Status And TF_VALID_RC) = 0 Then
                        .InvestigationCharacteristic.Value = Replace(resRS!Read_V2Rubric, "'", "`")
                        .InvestigationCharacteristicIdentification.Value = resRS!Read_V2RC
                        .CodeListQualifier.Value = Replace(resRS!Read_V2EDI, "'", "`")
                     Else
                        If blnUpper Then
                           strTemp = UCase(resRS!Result_Rubric)
                        Else
                           strTemp = resRS!Result_Rubric
                        End If
                        .InvestigationCharacteristic.Value = Replace(Trim(Left(strTemp, 35)), "'", "`")
                     End If
                  End With
                           
      '        Now check for any result comments
                  If resRS!Comment_Marker Then
                     With iceCmd
                        .CommandText = "ICELABCOMM_Report_ResultComment"
                        .Parameters(0).Value = resRS!Result_Index
                        Set RS = .Execute
                     End With
                  
                     RS.Filter = "Comment_Type <> 'A'"
                     msgData.FreeText msgEDI.SG2.SG6.SG18.Item(sg18Index), _
                                      RS, _
                                      ftxType, _
                                      cmtCnt, _
                                      "NHS002"
                  End If
               
                  If ftxType = "SPC" Then '  A numeric result
      '              Add any range data
                     With .SG20
                        If resRS!Lower_Range <> "" Or resRS!Upper_Range <> "" Then
                           .Append
                           With .Item(0)
                              .S20.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "20"
                              With .Rnd
                                 .RangeTypeQualifier.Value = "U"
                                 msgData.AdjustRangeData resRS!Lower_Range, resRS!Upper_Range
                                 .RangeMinimum.Value = msgData.LowRange
                                 .RangeMaximum.Value = msgData.HighRange
                              End With
                              
                              If resRS!Comment_Marker Then
                                 RS.Filter = "Comment_Type = 'A'"
                                 If RS.RecordCount > 0 Then
                                    msgData.RangeFreeText .FTX, _
                                                          RS
                                 End If
                              End If
                           End With
                        End If
                     End With
                  End If
                  
                  If RS.State = adStateOpen Then
                     RS.Close
                  End If
                  
                  sg18Index = sg18Index + 1
               End With
            End With
         End If
         
         resRS.MoveNext
         rPosn = resRS.AbsolutePosition - 1
     Loop
   End If
   
   resRS.Close
   Set resRS = Nothing
   Set iceCmd = Nothing
Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS002.SG18_Test"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub SG18_StandAlone(invRS As ADODB.Recordset, _
                            seqId As Integer, _
                            seqNo As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim resRS As New ADODB.Recordset
   Dim RS As New ADODB.Recordset
   Dim strArray() As String
   Dim localUOM As String
   Dim strTemp As String
   Dim strSi As String
   Dim ftxType As String
   Dim Status As enumInvResStatus
   Dim vData As Variant
   Dim vFields As Variant
   Dim rPosn As Integer
   Dim testDesc As String
   Dim rcIndex As Long

   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      
      If blnUseRCIndex Then
         .CommandText = "ICELABCOMM_Report_Results_By_Index"
      Else
         .CommandText = "ICELABCOMM_Report_Results_By_Code"
      End If
      
      .Parameters.Append .CreateParameter("InvId", adInteger, adParamInput, , invRS!Investigation_Index)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      Set resRS = .Execute
      .Parameters.Delete ("LTSIndex")
   End With
   
   Status = 0

'   If IsNull(resRS!EDI_RC_Index) Then
'      vFields = Array(16, 17, 19, 20, 21, 28, 29, 30, 31, 32, 33, 34, 5)
'   Else
'      vFields = Array(22, 23, 24, 25, 26, 35, 36, 37, 38, 39, 40, 41, 5)
'   End If
'
'   vData = resRS.GetRows(-1, adBookmarkFirst, vFields)
   resRS.MoveFirst
   rPosn = 0
   
   If IsNull(resRS!EDI_Local_Test_Code) Then
'   If IsNull(vData(0, rPosn)) Then
      testDesc = resRS!Result_Code & " (" & resRS!Result_Rubric & ") "
   Else
      testDesc = Replace(resRS!EDI_Local_Test_Code, Chr(0), "") & " (" & resRS!EDI_Local_Rubric & ") "
   End If
   
   testDesc = Replace(testDesc, "'", "`")
   If resRS!EDI_OP_Suppress Then
'   If vData(2, rPosn) Then
      Status = IS_INV_SUPPRESSED
      msgControl.LogReportMessage Status, testDesc & "Standalone test has output suppressed)"
   Else
   
      If IsNull(resRS!EDI_RC_Index) Then
'      If IsNull(vData(12, rPosn)) Then
         rcIndex = -1
      Else
         rcIndex = resRS!EDI_RC_Index 'vData(12, rPosn)
      End If
         
      msgEDI.SG2.SG6.SG18.Append
      eClass.FurtherInfo = "Repeat " & sg18Index
   '   Set sCLS = colSamp(dbData.SG18Sample(repIndex, ahReport.Investigations(invRpt).Sample_Index))
      With msgEDI.SG2.SG6
         .SG18.Item(sg18Index).GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
         localUOM = resRS!EDI_OP_UOM & ""
         
         With .SG18.Item(sg18Index)
            strTemp = msgData.AnalyseResultString(resRS!Result)
            strArray = Split(strTemp, "|")
            strTemp = strArray(0)
            If UBound(strArray) > 0 Then
               strSi = strArray(1)
            Else
               strSi = ""
            End If
         
            If repListRS!EDI_EMIS_System Then
               .SEQ.SequenceInformation.SequenceNumber.Value = seqNo
            End If
            
            ftxType = "RIT"
            
            If IsNumeric(strTemp) And resRS!Result_Type <> "T" Then
               ftxType = "SPC"
               .RSL.ResultTypeCoded.Value = "NV"
               .RSL.ResultDetails1.MeasurementValue.Value = strTemp
               .RSL.ResultDetails1.MeasurementSignificanceCoded.Value = strSi
               
               If IsNull(resRS!EDI_Nat_UOM) Then
                  .RSL.MeasurementUnitDetails.MeasurementUnit.Value = resRS!UOM_Code
               Else
                  .RSL.MeasurementUnitDetails.MeasurementUnitIdentification.Value = resRS!EDI_Nat_UOM
                  .RSL.MeasurementUnitDetails.CodeListQualifier.Value = "923"
               End If
               
            ElseIf strTemp <> "" Then
               .FTX.Append
               .FTX.Item(.FTX.Count - 1).TextSubjectQualifier.Value = ftxType
               .FTX.Item(.FTX.Count - 1).TextLiteral.FreeText1.Value = Replace(resRS!Result, "'", "`")
'               cmtCnt = cmtCnt + 1
            End If
            
            If resRS!Abnormal_Flag Then
               If Trim(resRS!Abnormal_Code & "") <> "" Then
                  If InStr(1, "NM|UN|LO|SL|HI|SH", Trim(resRS!Abnormal_Code & "")) = 0 Then
                     If Trim(resRS!Abnormal_Code & "") = "PA" Then
                        .RSL.ResultNormalcyIndicatorCoded.Value = ""
                     Else
                        .RSL.ResultNormalcyIndicatorCoded.Value = "UN"
                     End If
                  Else
                     .RSL.ResultNormalcyIndicatorCoded.Value = resRS!Abnormal_Code
                  End If
               End If
            End If
            
   '        Evaluate the read code here so we can check for a numeric result
            .INV.Append
            .INV(0).InvestigationCharacteristicQualifier.Value = "MQ"
            
            With .INV(0).InvestigationCharacteristicDetails
               If rcIndex = 0 Then
                  Status = IS_RC_NONE
                  If blnReadCodeInfo Then
                     msgControl.LogReportMessage Status, "Read Code suppressed by Lab system."
                  End If
               Else
                  If IsNull(resRS!Read_V2RC) Then
'                  If IsNull(vData(5, rPosn)) Then
                     Status = IS_RC_NONE
                     If blnReadCodeInfo Then
                        msgControl.LogReportMessage IS_RC_NONE, testDesc & "Not Read-Coded)"
                     End If
                  Else
                     If resRS!Read_Status = "D" Then
'                     If vData(9, rPosn) = "D" Then
                        Status = IS_RC_DELETED
                        msgControl.LogReportMessage IS_RC_DELETED, testDesc & resRS!Read_Comments
                     Else
                        If resRS!Read_Test = "F" Then
'                        If vData(8, rPosn) = "F" Then
                           Status = Status Or IS_RC_NA
                           msgControl.LogReportMessage IS_RC_NA, testDesc & "Read code flagged as inapplicable to a battery header in the bounded List."
                        End If
                        
                        If resRS!EDI_OP_Active = False Then
'                        If vData(3, rPosn) = False Then
                           Status = Status Or IS_TEST_INACTIVE
                           msgControl.LogReportMessage Status, testDesc & "Test not read-coded - marked as inactive."
                        End If
                                          
                        If resRS!Read_Ratio <> "T" Then
'                        If vData(10, rPosn) <> "T" Then
                           If IsNumeric(resRS!Result) Then
                              If Trim(resRS!UOM_Code) = "" Then
                                 If Trim(resRS!EDI_Nat_UOM & "") = "" Then
                                    Status = Status Or IS_RC_REMOVED
                                    msgControl.LogReportMessage IS_RC_REMOVED, testDesc & "The bounded list specifies a UOM is required with this read code - None present."
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               
               If (Status And TF_VALID_RC) = 0 Then
                  .InvestigationCharacteristic.Value = Replace(resRS!Read_V2Rubric, "'", "`") 'vData(6, rPosn)
                  .InvestigationCharacteristicIdentification.Value = resRS!Read_V2RC 'vData(5, rPosn)
                  .CodeListQualifier.Value = Replace(resRS!Read_V2EDI, "'", "`") 'vData(7, rPosn)
               Else
                  If blnUpper Then
                     strTemp = UCase(resRS!Result_Rubric)
                  Else
                     strTemp = resRS!Result_Rubric
                  End If
                  .InvestigationCharacteristic.Value = Replace(Trim(Left(strTemp, 35)), "'", "`")
               End If
            End With
            
            .RFF.Append
            .RFF(0).Reference.ReferenceQualifier.Value = "ASL"
            .RFF(0).Reference.ReferenceNumber.Value = seqId
            
            If (resRS!Status = "PR" Or resRS!Status = "FR" Or resRS!Status = "NA") Then
               .STS.StatusEvent.StatusEventCoded.Value = resRS!Status
            End If
   
   '        Check for any Investigation comments
            If invRS!Comment_Marker Then
               With iceCmd
                  .CommandText = "ICELABCOMM_Report_InvestComments"
                  .Parameters("InvId").Value = invRS!Investigation_Index
                  Set RS = .Execute
               End With
               
               msgData.FreeText msgEDI.SG2.SG6.SG18.Item(sg18Index), _
                                RS, _
                                "SPC", _
                                0, _
                                "NHS002"
               RS.Close
            End If
            
   '        Now check for any result comments
            If resRS!Comment_Marker Then
               With iceCmd
                  .CommandText = "ICELABCOMM_Report_ResultComment"
                  .Parameters(0).Value = resRS!Result_Index
                  Set RS = .Execute
               End With
            
               RS.Filter = "Comment_Type <> 'A'"
               msgData.FreeText msgEDI.SG2.SG6.SG18.Item(sg18Index), _
                                RS, _
                                ftxType, _
                                0, _
                                "NSH002"

            End If
         End With
         
         If ftxType = "SPC" Then '  A numeric result
   '        Add any range data
            With .SG18.Item(sg18Index).SG20
               If resRS!Lower_Range <> "" Or resRS!Upper_Range <> "" Then
                  eClass.FurtherInfo = "SG20 segment (" & sg18Index & ")"
                  .Append
                  With .Item(0)
                     .S20.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "20"
                     With .Rnd
                        .RangeTypeQualifier.Value = "U"
                           msgData.AdjustRangeData resRS!Lower_Range, resRS!Upper_Range
                           .RangeMinimum.Value = msgData.LowRange
                           .RangeMaximum.Value = msgData.HighRange
                     End With
                     
                     If resRS!Comment_Marker Then
                        RS.Filter = "Comment_Type = 'A'"
                        If RS.RecordCount > 0 Then
                           msgData.RangeFreeText .FTX, _
                                                 RS
                        End If
                     End If
                  End With
               End If
            End With
         End If
      End With
      
      If RS.State = adStateOpen Then
         RS.Close
      End If
      
      sg18Index = sg18Index + 1
   End If
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS002.SG18_Standalone"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub SG18_Battery(invRS As ADODB.Recordset, _
                         seqId As Integer, _
                         seqNo As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim strArray() As String
   Dim strTemp As String
   Dim cmtCnt As Integer
   Dim k As Integer
   Dim vTxt As Variant
   Dim strRC As String
   Dim j As Integer
   Dim TestType As Integer
   Dim Status As enumInvResStatus
   Dim vData As Variant
   Dim vFields As Variant
   Dim testDesc As String
   Dim testCount As Integer
   Dim rcIndex As Long
   Dim rPosn As Long
   Dim vBMark As Variant
   
   If IsNull(invRS!EDI_Local_Test_Code) Then
      testDesc = invRS!Investigation_Code
   Else
      testDesc = Replace(invRS!EDI_Local_Test_Code, Chr(0), "")
   End If
   
   testDesc = testDesc & " (" & invRS!Investigation_Requested & ") "
   testDesc = Replace(testDesc, "'", "`")
   
   If invRS!EDI_OP_Suppress Then
'   If vData(2, rPosn) Then
      msgControl.LogReportMessage IS_INV_SUPPRESSED, testDesc & "Battery Header suppressed - no output for this investigation."
      seqNo = seqNo - 1
   Else
      
      msgEDI.SG2.SG6.SG18.Append
      With msgEDI.SG2.SG6.SG18(sg18Index)
         .GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
         .INV.Append
         .INV(0).InvestigationCharacteristicQualifier.Value = "MQ"
            
         If IsNull(invRS!EDI_RC_Index) Then
'         If IsNull(vData(10, rPosn)) Then
            rcIndex = -1
         Else
            rcIndex = invRS!EDI_RC_Index 'vData(10, rPosn)
         End If
         
         If rcIndex = 0 Then
            Status = IS_RC_NONE
            If blnReadCodeInfo Then
               msgControl.LogReportMessage Status, "Read Code suppressed by Lab system."
            End If
         Else
            If IsNull(invRS!Read_V2RC) Then
'            If IsNull(vData(4, rPosn)) Then
'               If Not IsNull(vData(0, rPosn)) Then
                  Status = IS_RC_NONE
                  If blnReadCodeInfo Then
                     msgControl.LogReportMessage IS_RC_NONE, testDesc & "not Read-Coded"
                  End If
'               End If
            Else
               If invRS!Read_Status = "D" Then
'               If vData(8, rPosn) = "D" Then
                  Status = IS_RC_DELETED
                  msgControl.LogReportMessage IS_RC_DELETED, testDesc & invRS!Read_Comments
               Else
                  If invRS!Read_Battery = "F" Then
'                  If vData(7, rPosn) = "F" Then
                     Status = IS_RC_NA
                     msgControl.LogReportMessage IS_RC_NA, testDesc & "Read code flagged as inapplicable to a battery headers in the bounded List."
                  End If
                  
                  If invRS!EDI_OP_Active = False Then
'                  If vData(3, rPosn) = False Then
                     Status = IS_INV_INACTIVE
                     msgControl.LogReportMessage Status, testDesc & "Read code has been set as inactive."
                  End If
               End If
            End If
         End If
         
         If (Status And TF_VALID_RC) = 0 Then
            .INV(0).InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = Replace(invRS!Read_V2Rubric, "'", "`") 'vData(5, rPosn)
            .INV(0).InvestigationCharacteristicDetails.InvestigationCharacteristicIdentification.Value = invRS!Read_V2RC 'vData(4, rPosn)
            .INV(0).InvestigationCharacteristicDetails.CodeListQualifier.Value = Replace(invRS!Read_V2EDI, "'", "`") 'vData(6, rPosn)
         Else
            If blnUpper Then
               strTemp = UCase(invRS!Investigation_Requested)
            Else
               strTemp = invRS!Investigation_Requested
            End If
            .INV(0).InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = Replace(Trim(Left(strTemp, 35)), "'", "`")
         End If
            
         .SEQ.SequenceInformation.SequenceNumber.Value = seqNo
            
         .RFF.Append
         .RFF(0).Reference.ReferenceQualifier.Value = "ASL"
         .RFF(0).Reference.ReferenceNumber.Value = seqId
         
         If invRS!Comment_Marker Then
            With iceCmd
               .ActiveConnection = iceCon
               .CommandType = adCmdStoredProc
               .CommandText = "ICELABCOMM_Report_InvestComments"
               .Parameters.Append .CreateParameter("InvId", adInteger, adParamInput, , invRS!Investigation_Index)
               Set RS = .Execute
            End With
         
            msgData.FreeText msgEDI.SG2.SG6.SG18.Item(sg18Index), _
                             RS, _
                             "SPC", _
                             , _
                             "NHS002"
         End If
      End With
      
      SG18_Test invRS, seqNo
   End If
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHS002.SG18_Battery"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG18_Discharge(ReportId As Long)
   Dim iceCmd As New ADODB.Command
   Dim resRS As New ADODB.Recordset
   Dim dischargeLines As New Collection
   Dim chunks() As String, splitParagraph() As String
   Dim paragraph As String
   Dim I, i2 As Integer
   Dim iFtxId As Integer
   Dim iFtxSeq As Integer

   On Error GoTo ErrorHandler

   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_Select_Letter_For_EDIFACT2"

      .Parameters.Append .CreateParameter("serviceReportIndex", adInteger, adParamInput, , ReportId)
      Set resRS = .Execute
   End With

   'This segment is mandatory, though we have nothing to put in it....
   msgEDI.SG2.SG6.SG16.Append
   With msgEDI.SG2.SG6.SG16.Item(0)
      .S16.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "16"
      .SEQ.SequenceInformation.SequenceNumber.Value = 1
      .SPC.Append

      With .SPC.Item(0)
         .SpecimenCharacteristicQualifier.Value = "TSP"
         .SpecimenCharacteristicsDetails.SpecimenCharacteristicIdentification.Value = "T027"
         .SpecimenCharacteristicsDetails.CodeListQualifier.Value = "920"
      End With
   End With

   'Harvest all the data now, as we need to regroup into lines.
   Do Until resRS.EOF

      If Trim(resRS!Paragraph_Title) <> "" Then
         dischargeLines.Add RTrim(CStr(resRS!Paragraph_Title))
      End If

      'Split the Paragraph_Text field by vbCrLf...
      chunks = VBA.Split(CStr(resRS!Paragraph_Text), vbCrLf)
      For I = 0 To UBound(chunks)
         'Remove these invalid chars from the vbCrLf chunk:
         chunks(I) = Replace(chunks(I), vbTab, "")
         chunks(I) = Replace(chunks(I), Chr(0), "")
         chunks(I) = Replace(chunks(I), Chr(160), "")

         If RTrim(chunks(I)) <> "" Then
            '...Then split the vbCrLf chunk up into segments 80 chars big
            splitParagraph = msgData.DataToArray(CStr(chunks(I)), 70)
            
            For i2 = 0 To UBound(splitParagraph)
               dischargeLines.Add RTrim(splitParagraph(i2))
            Next
         End If
      Next

      resRS.MoveNext
   Loop

   i2 = 1

   msgEDI.SG2.SG6.SG18.Append
   With msgEDI.SG2.SG6.SG18(sg18Index)
      Dim iSeq As Integer
      
         .GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
         .INV.Append
         .INV(0).InvestigationCharacteristicQualifier.Value = "MQ"
         .INV(0).InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = "Discharge" & iSeq
   
         .RFF.Append
         .RFF(0).Reference.ReferenceQualifier.Value = "ASL"
         .RFF(0).Reference.ReferenceNumber.Value = 1

         '.SEQ.SequenceInformation.SequenceNumber.Value = iSeq

         'Shall I bother with this?...
         If resRS.RecordCount = 0 Then
            eClass.CurrentProcedure = "IceMsg.Edifact_NHS002.MapSG18_Discharge"
            eClass.Add 88888, "Unable to retrieve discharge information for Service_Report_Index " & ReportId, ""
            Exit Sub
         End If
  
     ''      msgData.FreeText msgEDI.SG2.SG6.SG18.Item(sg18Index), _
     ''                       resRS, _
     ''                       "RIT", _
     ''                       0, _
     ''                       "NHS002"
      
      If dischargeLines.Count > 99 Then
        iFtxSeq = -1
        For iSeq = 1 To dischargeLines.Count
          iFtxId = (iSeq - 1) Mod 5
          
          If iFtxId = 0 Then
            .FTX.Append
            iFtxSeq = iFtxSeq + 1
            .FTX(iFtxSeq).TextSubjectQualifier.Value = "RIT"
          End If
            
          Select Case iFtxId
            Case 0
              .FTX(iFtxSeq).TextLiteral.FreeText1.Value = dischargeLines.Item(iSeq)
              
            Case 1
              .FTX(iFtxSeq).TextLiteral.FreeText2.Value = dischargeLines.Item(iSeq)
            
            Case 2
              .FTX(iFtxSeq).TextLiteral.FreeText3.Value = dischargeLines.Item(iSeq)
            
            Case 3
              .FTX(iFtxSeq).TextLiteral.FreeText4.Value = dischargeLines.Item(iSeq)
            
            Case 4
              .FTX(iFtxSeq).TextLiteral.FreeText5.Value = dischargeLines.Item(iSeq)
            
          End Select
        Next
      
      Else
        For iSeq = 1 To 99
           For I = 1 To 99
              If ((iSeq - 1) * 99) + I > dischargeLines.Count Then
                 Exit For
              End If
  
              .FTX.Append
              .FTX.Item(.FTX.Count - 1).TextSubjectQualifier.Value = "RIT"
              .FTX.Item(.FTX.Count - 1).TextLiteral.FreeText1.Value = dischargeLines.Item(((iSeq - 1) * 99) + I)
           Next
  
           If (iSeq * 99) + 1 > dischargeLines.Count Then
              Exit For
           End If
        Next
      End If
   End With

   Exit Sub


ErrorHandler:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS002.MapSG18_Discharge"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub


