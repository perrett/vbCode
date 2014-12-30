VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Edifact_NHS003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private blnHeader As Boolean
Private blnBody As Boolean
Private RepIndex As Long
Private sg18Index As Integer
Private srcRS As ADODB.Recordset
Private msgEDI As nhs003.MsgMEDRPT
Private msgErr As nhs003.ErrorCollection
Private strArray() As String
Private DownloadInd As Integer

Private Const TF_VALID_RC = &HA5

Private colSamp As New Collection

Public Property Get AllowApostrophe() As Boolean
   AllowApostrophe = True
End Property

Public Function CreateMessage(ReportId As Long) As Variant
   On Local Error GoTo procEH
   Dim msgBuf As String
   Dim I As Integer
   
   Set msgEDI = CreateObject("nhs003.MsgMEDRPT") ' New nhs003.MsgMEDRPT
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
   MapSG10 ReportId
   sg18Index = 0
   MapSG6_SG16 ReportId
   If msgControl.ReportStatus > 0 Then
      msgBuf = "Errors encountered when generating message'See comments for description'"
   Else
      With msgEDI.UNT
         .MessageReferenceNumber.Value = msgControl.MessageId
         .NumberOfSegmentsInAMessage.Value = 0
      End With
      
      Set msgErr = msgEDI.Validate
      
      If msgErr.Count > 0 Then
         msgControl.ReportStatus = MS_PARSE_FAIL
         For I = 0 To msgErr.Count - 1
            msgControl.LogReportMessage MS_PARSE_FAIL, msgControl.CurrentReport & ": " & _
                                                       msgErr(I).Path & " - " & _
                                                       msgErr(I).Description
         Next I
      End If
      
      msgBuf = msgEDI.parse
      
'      msgData.InvalidChars = ""
'      msgData.ReplaceChars = ""
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
   eClass.CurrentProcedure = "Edifact_NHS003.AddMessage"
   eClass.FurtherInfo = "Report: " & ReportId
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function FileHeader() As String
   Dim strHeader As String
   Dim flags As String
   
   strHeader = "UNA:+.? '" & _
                "UNB+UNOC:3+<SENDINGTRADER>:80+<RECEIVINGTRADER>:80+<DATE:TIME>+<INTERCHANGE>++MEDRPT"
   strHeader = Replace(strHeader, "<SENDINGTRADER>", msgControl.CurrentSender)
   strHeader = Replace(strHeader, "<RECEIVINGTRADER>", msgControl.CurrentReceiver)
   strHeader = Replace(strHeader, "<DATE:TIME>", Format(Now(), "YYMMDD:HHNN"))
   strHeader = Replace(strHeader, "<INTERCHANGE>", msgControl.CurrentInterchange)
   
   If msgControl.CurrentAcks Then
'     Acknowledgements required
      If msgControl.CurrentExtension = ".XEN" Then
'        File will be Encrypted
         If msgControl.TestData Then
'           Set the test flag
            flags = "++1+P1+1"
         Else
            flags = "++1+P1"
         End If
      Else
         If msgControl.TestData Then
'           Set the test flag
            flags = "++1++1"
         Else
            flags = "++1"
         End If
      End If
   Else
'     Acknowledgements not required
      If msgControl.CurrentExtension = ".XEN" Then
         If msgControl.TestData Then
            flags = "+++P1+1"
         Else
            flags = "+++P1"
         End If
      Else
         If msgControl.TestData Then
            flags = "++++1"
         Else
            flags = ""
         End If
      End If
   End If
   
   FileHeader = strHeader & flags & "'"
End Function

Public Property Get FileTrailer() As String
   FileTrailer = "UNZ+" & msgControl.MsgInBatch & "+" & msgControl.CurrentInterchange & "'"
End Property

Public Property Get InvalidChars() As Variant
   ReDim strArray(2)
   strArray(0) = Chr(10) & Chr(13)
   strArray(1) = Chr(0) & Chr(0)
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
   Dim pLTS As ADODB.Parameter
   Dim pMatch As ADODB.Parameter
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim nat902 As String
   Dim hpData As New clsHealthParties
   Dim gpCode As String
'   Dim gp902Name As String
'   Dim gp906Code As String
'   Dim gp906Name As String
'   Dim blnNat902 As Boolean
'   Dim blnNat906 As Boolean
   
'   With iceCmd
'      .ActiveConnection = iceCon
'      .CommandType = adCmdStoredProc
'      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportIndex)
'      .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , 0)
'   End With
   
   eClass.FurtherInfo = "EDI_HP_Type = 901"
   With msgEDI.SG1
   
'******************
'        901
'******************
      .Append
      With .Item(0).NAD
         .PartyQualifier.Value = "MR"
         .PartyIdentificationDetails.PartyIdIdentification.Value = repListRS!EDI_Loc_Nat_Code_To
         .PartyIdentificationDetails.CodeListQualifier.Value = "901"
      End With
      
      .Item(0).S01.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "01"
      .Item(0).SPR.ServiceProviderQualifier.Value = "ORG"
         
      DownloadInd = repListRS!GP_Download
      hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
      hpData.LTSIndex = repListRS!EDI_LTS_Index
      hpData.DownloadRequest = DownloadInd
      hpData.Read ReportIndex, True, True, blnClinNameUseAll
      
'******************
'        906
'******************
            
      .Append
      With .Item(1)
         .S01.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "01"
         .NAD.PartyQualifier.Value = "PO"
         .NAD.PartyName.PartyName1.Value = hpData.HP906Name
         
         gpCode = hpData.HP906Code
         If hpData.HP906Nat Then
            .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = gpCode
            
            If Left(gpCode, 1) = "C" Then
               .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "902"
            Else
               .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
            End If
         Else
            .RFF.Reference.ReferenceQualifier.Value = "AHI"
            If gpCode = "" Or gpCode = "~" Then
               gpCode = "UNK"
            End If
            
            .RFF.Reference.ReferenceNumber.Value = gpCode
         End If
         .SPR.ServiceProviderQualifier.Value = "PRO"
      End With
   
'******************
'        902
'******************
      
      .Append
      With .Item(2)
         .S01.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "01"
         .NAD.PartyQualifier.Value = "MR"
         .NAD.PartyName.PartyName1.Value = hpData.HP902Name
         
         gpCode = hpData.HP902Code
         If hpData.HP902Nat Then
            .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = gpCode
            
            If Left(gpCode, 1) = "C" Then
               .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "902"
            Else
               .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
            End If
         Else
            .RFF.Reference.ReferenceQualifier.Value = "AHI"
            If gpCode = "" Or gpCode = "~" Then
               gpCode = "UNK"
            End If
            
            .RFF.Reference.ReferenceNumber.Value = gpCode
         End If
         .SPR.ServiceProviderQualifier.Value = "PRO"
      End With
   
'******************
'        ORG
'******************

      .Append
      With .Item(3)
         .S01.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "01"
         .NAD.PartyQualifier.Value = "SLA"
         .SPR.ServiceProviderQualifier.Value = "ORG"
         .NAD.PartyName.PartyName1.Value = msgControl.Organisation
      End With
   
'******************
'        DEPT
'******************
                  
      .Append
      With .Item(4)
         .S01.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "01"
         .NAD.PartyQualifier.Value = "SLA"
         If repListRS!Specialty & "" = "" Then
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
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS003.MapSG1"
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
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG2"
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
   
   eClass.FurtherInfo = "Patient Demographics"
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
                     strArray(j) = Trim(strRepAdr(I)) & Chr(255)  '  Symphonia fails with escaped char as last on line
                  Else
                     strArray(j) = Trim(Left(strRepAdr(I), 34)) & Chr(255)
                  End If
                  
                  j = j + 1
               End If
            Next I
               
'            If j = 0 Then
'               If RS!Pat_PostCode & "" = "" Then
'                  msgControl.ReportStatus = RS_DATA_INTEGRITY
'                  msgControl.LogReportMessage MS_DATA_INTEGRITY, "No Address or postcode for " & _
'                                                                 RS!Surname
'               End If
'            Else
            If j > 0 Then
               I = UBound(strArray)
               With .ADR.AddressDetails
                  .AddressFormatCoded.Value = "US"
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
               End With
            End If
             
            If RS!Pat_PostCode & "" <> "" Then
               .ADR.PostcodeIdentification.Value = RS!Pat_PostCode
            End If
         End If
      End With
      
      With .SG6
         .S06.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "06"
         If msgControl.AnonymizeData Then
            .SG7.PNA.IdentificationNumber.IdentityNumber.Value = "9999999999"
            .SG7.PNA.IdentificationNumber.IdentityNumberQualifier.Value = "OPI"
         Else
            
   '        Do we have a valid NHS number?
            If RS!New_NHS_No <> "" Then
               .SG7.PNA.IdentificationNumber.IdentityNumber.Value = RS!New_NHS_No
               .SG7.PNA.IdentificationNumber.IdentityNumberQualifier.Value = "OPI"
            ElseIf RS!Hospital_Number <> "" And Left(RS!Hospital_Number, 3) <> "ICE" Then
   '           No NHS number but there is a valid hospital number
               .RFF.Reference.ReferenceQualifier.Value = "SSI"
               .RFF.Reference.ReferenceNumber.Value = RS!Hospital_Number
            End If
         End If
      End With
      
      .SG6.SG7.S07.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "07"
      With .SG6.SG7.PNA
         .PartyQualifier.Value = "PAT"
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
   
      With .SG6.SG7
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
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG6_SG7"
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
   
   eClass.FurtherInfo = "Report Level Clinical Details"
   RS.Filter = "Comment_Type = 'D'"
   If RS.RecordCount > 0 Then
      With msgEDI.SG2.SG6
         .SG10.S10.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "10"
         .SG10.CIN.ClinicalInformationQualifier.Value = "UN"
         msgData.FreeText .SG10, _
                          RS, _
                          "CID", _
                          0, _
                          "NHS003"
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
   eClass.CurrentProcedure = "AHSLMessaging.Edifact3.MapSG10"
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
   Dim cText As String
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Sample"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   sampSeq = 0
   battSeq = 1
   
   If blnUseRCIndex Then
      cText = "ICELABCOMM_Report_Invest_By_Index"
   Else
      cText = "ICELABCOMM_Report_Invest_By_Code"
   End If
   
   Do Until RS.EOF
      With iceCmd
         .CommandText = cText
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
         If RS2!Result_Recs > 0 Then
            If IsNull(RS2!Result_Index) Or RS2!Result_Recs > 1 Then
               SG18_Battery RS2, sampSeq + 1, battSeq
               battSeq = battSeq + 1
            Else
               SG18_StandAlone RS2, sampSeq + 1
            End If
         Else
            msgControl.LogReportMessage IS_TEST_SUPPRESSED, RS2!Investigation_Code & " (" & _
                                        RS2!Investigation_Requested & ") - All tests for this investigation have been suppressed"
         End If
         
         RS2.MoveNext
      Loop
      
      RS2.Close
      
      If sg18Index > lastIndex Then
         
         msgEDI.SG2.SG6.SG16.Append
         With msgEDI.SG2.SG6.SG16.Item(sampSeq)
            .S16.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "16"
            .SEQ.SequenceInformation.SequenceNumber.Value = sampSeq + 1
            .SPC.Append
            
            With .SPC.Item(0)
               .SpecimenCharacteristicsDetails.SpecimenCharacteristic.Value = Replace(RS!Sample_Text, "'", "`")
               .SpecimenCharacteristicQualifier.Value = "TSP"
               .SpecimenCharacteristicsDetails.CodeListQualifier.Value = RS!Sample_Type
            End With
            
            If RS!Fasting Then
               .SPC.Append
               With .SPC(1)
                  .SpecimenCharacteristicQualifier.Value = "FS"
                  .SpecimenCharacteristicsDetails.SpecimenCharacteristicIdentification.Value = "F"
'                  .SpecimenCharacteristicsDetails.SpecimenCharacteristic.Value = "F"
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
                                "NHS003", _
                                9
            End If
            RS2.Close
         End With
         
         sampSeq = sampSeq + 1
      Else
         msgControl.LogReportMessage IS_SAMPLE_SUPPRESSED, RS!Sample_Text & _
                                     "- All tests on this sample have their output suppressed"
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
   Set iceCmd = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS0033.MapSG6_SG16"
   eClass.FurtherInfo = "Sample_Index = " & RS!Sample_Index
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
'      .CommandText = "ICELABCOMM_Report_Results"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , invRS!Investigation_Index)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      Set resRS = .Execute
      .Parameters.Delete 1
   End With
   
   resRS.Filter = "EDI_OP_Suppress = False"
   testCount = resRS.RecordCount

   If testCount = 0 Then
      Status = IS_INV_SUPPRESSED
      msgControl.LogReportMessage Status, invRS!Investigation_Code & " (" & invRS!Investigation_Requested & _
                                  ") ALL the tests in this battery have their output suppressed"
   Else
      resRS.Filter = ""
      
      sg18Index = sg18Index + 1
      testCount = 1
      cmtCnt = 0
      Do Until resRS.EOF
         eClass.FurtherInfo = "Investigation Index = " & invRS!Investigation_Index & _
                              "Result Index = " & resRS!Result_Index
         If IsNull(resRS!EDI_RC_Index) Then
'         If IsNull(vData(12, rPosn)) Then
            rcIndex = -1
         Else
            rcIndex = resRS!EDI_RC_Index 'vData(12, rPosn)
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
                  .INV.InvestigationCharacteristicQualifier.Value = "MQ"
                  
                  .RFF.Reference.ReferenceQualifier.Value = "ARL"
                  .RFF.Reference.ReferenceNumber.Value = RelatedBattery
                  
                  strTemp = msgData.AnalyseResultString(resRS!Result)
                  strArray = Split(strTemp, "|")
                  strTemp = strArray(0)
                  
                  If UBound(strArray) > 0 Then
                     strSi = strArray(1)
                  Else
                     strSi = ""
                  End If
                  
                  If (resRS!Status = "PR" Or resRS!Status = "SR") Then
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
                     .RSL.MeasurementUnitDetails.MeasurementUnit.Value = IIf((resRS!EDI_Nat_UOM & "" = ""), resRS!UOM_Code, resRS!EDI_Nat_UOM)
                  ElseIf strTemp <> "" Then
                     .FTX.Append
                     .FTX.Item(.FTX.Count - 1).TextSubjectQualifier.Value = ftxType
                     .FTX.Item(.FTX.Count - 1).TextLiteral.FreeText1.Value = Replace(resRS!Result, "'", "`")
   '                  cmtCnt = cmtCnt + 1
                  End If
                           
                  If resRS!Abnormal_Flag Then
                     If Trim(resRS!Abnormal_Code & "") <> "" Then
                        If InStr(1, "HI|LO|OR|PA", Trim(resRS!Abnormal_Code & "")) = 0 Then
                           .RSL.ResultNormalcyIndicatorCoded.Value = "OR"
                        Else
                           .RSL.ResultNormalcyIndicatorCoded.Value = resRS!Abnormal_Code
                        End If
                     End If
                  End If
                                                
                  .INV.InvestigationCharacteristicQualifier.Value = "MQ"
                  
                  Status = 0
                  
                  With .INV.InvestigationCharacteristicDetails
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
                              msgControl.LogReportMessage Status, testDesc & "Not Read-Coded"
                           End If
                        Else
                           If resRS!Read_Status = "D" Then
'                           If vData(9, rPosn) = "D" Then
                              Status = IS_RC_DELETED
                              msgControl.LogReportMessage Status, testDesc & resRS!Read_Comments
                           Else
                              If resRS!Read_Test = "F" Then
'                              If vData(8, rPosn) = "F" Then
                                 Status = Status Or IS_RC_NA
                                 msgControl.LogReportMessage Status, testDesc & "Read code flagged as inapplicable t a test in the bounded list."
                              End If
                        
                              If resRS!EDI_OP_Active = False Then
'                              If vData(3, rPosn) = False Then
                                 Status = Status Or IS_TEST_INACTIVE
                                 msgControl.LogReportMessage Status, testDesc & "Test not read-coded - marked as inactive."
                              End If
                           
                              If resRS!Read_Ratio <> "T" Then
'                              If vData(10, rPosn) <> "T" Then
                                 If ftxType = "SPC" Then ' A numeric result IsNumeric(resRS!Result) Then
                                    If Trim(resRS!UOM_Code) = "" Then
                                       If Trim(resRS!EDI_Nat_UOM & "") = "" Then
                                          Status = Status Or IS_RC_REMOVED
                                          msgControl.LogReportMessage Status, testDesc & "The bounded list specifies a UOM is required with this read code - None present."
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
                           
'                 Now check for any result comments
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
                                      cmtCnt

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
'         rPosn = rPosn + 1
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
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS0033.SG18_Test"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub SG18_StandAlone(invRS As ADODB.Recordset, _
                            seqId As Integer)
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
'      .CommandText = "ICELABCOMM_Report_Results"
      .Parameters.Append .CreateParameter("InvId", adInteger, adParamInput, , invRS!Investigation_Index)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      Set resRS = .Execute
      .Parameters.Delete ("LTSIndex")
   End With
   
   Status = 0

   If IsNull(resRS!EDI_Local_Test_Code) Then
      testDesc = resRS!Result_Code & " (" & resRS!Result_Rubric & ") "
   Else
      testDesc = Replace(resRS!EDI_Local_Test_Code, Chr(0), "") & " (" & resRS!EDI_Local_Rubric & ") "
   End If
   
   testDesc = Replace(testDesc, "'", "`")
   eClass.FurtherInfo = "Investigation Index = " & invRS!Investigation_Index & _
                        "Result Index = " & resRS!Result_Index
   
   If resRS!EDI_OP_Suppress Then
'   If vData(2, rPosn) Then
      Status = IS_INV_SUPPRESSED
      msgControl.LogReportMessage Status, testDesc & "Standalone test has output suppressed"
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
            
            ftxType = "RIT"
   '         strRes = dbData.AbnormalCode(ahReport.Investigations(invRpt).Results(0).ResultID)
   '         strArray = Split(strRes, "|")
            
            If IsNumeric(strTemp) And resRS!Result_Type <> "T" Then
               ftxType = "SPC"
               .RSL.ResultTypeCoded.Value = "NV"
               .RSL.ResultDetails1.MeasurementValue.Value = strTemp
               .RSL.ResultDetails1.MeasurementSignificanceCoded.Value = strSi
               .RSL.MeasurementUnitDetails.MeasurementUnit.Value = IIf(resRS!EDI_Nat_UOM & "" = "", resRS!UOM_Code, resRS!EDI_Nat_UOM)
               
            ElseIf strTemp <> "" Then
               .FTX.Append
               .FTX.Item(.FTX.Count - 1).TextSubjectQualifier.Value = ftxType
               .FTX.Item(.FTX.Count - 1).TextLiteral.FreeText1.Value = Replace(resRS!Result, "'", "`")
            End If
            
            If resRS!Abnormal_Flag And Trim(resRS!Abnormal_Code & "") <> "" Then
               If InStr(1, "HI|LO|OR|PA", Trim(resRS!Abnormal_Code & "")) = 0 Then
                  .RSL.ResultNormalcyIndicatorCoded.Value = "OR"
               Else
                  .RSL.ResultNormalcyIndicatorCoded.Value = resRS!Abnormal_Code
               End If
            End If
            
            .INV.InvestigationCharacteristicQualifier.Value = "MQ"
            
            With .INV.InvestigationCharacteristicDetails
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
                        msgControl.LogReportMessage Status, testDesc & " Not Read-Coded)"
                     End If
                  Else
                     If resRS!Read_Status = "D" Then
'                     If vData(9, rPosn) = "D" Then
                        Status = IS_RC_DELETED
                        msgControl.LogReportMessage Status, testDesc & resRS!Read_Comments
                     Else
                        If resRS!Read_Test = "F" Then
'                        If vData(8, rPosn) = "F" Then
                           Status = Status Or IS_RC_NA
                           msgControl.LogReportMessage Status, testDesc & "Read code flagged as inapplicable to a test in the bounded list"
                        End If
                        
                        If resRS!EDI_OP_Active = False Then
'                        If vData(3, rPosn) = False Then
                           Status = Status Or IS_TEST_INACTIVE
                           msgControl.LogReportMessage Status, testDesc & "Test not read-coded - marked as inactive"
                        End If
                     
                        If resRS!Read_Ratio <> "T" Then
'                        If vData(10, rPosn) <> "T" Then
                           If ftxType = "SPC" Then ' A numeric result IsNumeric(.RSL.ResultTypeCoded.Value) Then
                              If Trim(resRS!UOM_Code) = "" Then
                                 If Trim(resRS!EDI_Nat_UOM & "") = "" Then
                                    Status = Status Or IS_RC_REMOVED
                                    msgControl.LogReportMessage Status, testDesc & "The bounded list specifies a UOM is required with this read code - None present."
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
            
            .RFF.Reference.ReferenceQualifier.Value = "ASL"
            .RFF.Reference.ReferenceNumber.Value = seqId
            
            If (resRS!Status = "PR" Or resRS!Status = "SR") Then
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
                                0
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
                                0
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
         
         If RS.State = adStateOpen Then
            RS.Close
         End If
         
      End With
      sg18Index = sg18Index + 1
   End If
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsg.Edifact_NHS003.SG18_Standalone"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub SG18_Battery(invRS As ADODB.Recordset, _
                         seqId As Integer, _
                         ByRef seqNo As Long)
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
   Dim RCStatus As enumInvResStatus
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
   eClass.FurtherInfo = "Investigation Index = " & invRS!Investigation_Index
   
   If invRS!EDI_OP_Suppress Then
'   If vData(2, rPosn) Then
      msgControl.LogReportMessage IS_INV_SUPPRESSED, testDesc & "Battery Header suppressed - no output for this investigation"
      seqNo = seqNo - 1
   Else
      
      msgEDI.SG2.SG6.SG18.Append
      With msgEDI.SG2.SG6.SG18(sg18Index)
         .GIS.ProcessingIndicator.ProcessingIndicatorCoded.Value = "N"
         .INV.InvestigationCharacteristicQualifier.Value = "MQ"
         
         If IsNull(invRS!EDI_RC_Index) Then
'         If IsNull(vData(10, rPosn)) Then
            rcIndex = -1
         Else
            rcIndex = invRS!EDI_RC_Index 'vData(10, rPosn)
         End If
         
         If rcIndex = 0 Then
            RCStatus = IS_RC_NONE
            If blnReadCodeInfo Then
               msgControl.LogReportMessage RCStatus, "Read Code suppressed by Lab system."
            End If
         Else
            If IsNull(invRS!Read_V2RC) Then
'            If IsNull(vData(4, rPosn)) Then
'               If Not IsNull(vData(0, rPosn)) Then
                  RCStatus = IS_RC_NONE
                  If blnReadCodeInfo Then
                     msgControl.LogReportMessage RCStatus, testDesc & "not Read-Coded"
                  End If
'               End If
            Else
               If invRS!Read_Status = "D" Then
'               If vData(8, rPosn) = "D" Then
                  RCStatus = IS_RC_DELETED
                  msgControl.LogReportMessage RCStatus, testDesc & invRS!Read_Comments
               Else
                  If invRS!Read_Battery = "F" Then
'                  If vData(7, rPosn) = "F" Then
                     RCStatus = IS_RC_NA
                     msgControl.LogReportMessage RCStatus, testDesc & "Read code flagged as inapplicable to a battery header in the bounded List."
                  End If
                  
                  If invRS!EDI_OP_Active = False Then
'                  If vData(3, rPosn) = False Then
                     RCStatus = IS_INV_INACTIVE
                     msgControl.LogReportMessage RCStatus, testDesc & "Read code has been set as inactive."
                  End If
               End If
            End If
         End If
         
         If (RCStatus And TF_VALID_RC) = 0 Then
            .INV.InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = Replace(invRS!Read_V2Rubric, "'", "`") 'vData(5, rPosn)
            .INV.InvestigationCharacteristicDetails.InvestigationCharacteristicIdentification.Value = invRS!Read_V2RC 'vData(4, rPosn)
            .INV.InvestigationCharacteristicDetails.CodeListQualifier.Value = Replace(invRS!Read_V2EDI, "'", "`") 'vData(6, rPosn)
         Else
            If blnUpper Then
               strTemp = UCase(invRS!Investigation_Requested)
            Else
               strTemp = invRS!Investigation_Requested
            End If
            .INV.InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = Replace(Trim(Left(strTemp, 35)), "'", "`")
         End If
            
   '         If SeqReqd Then
         .SEQ.SequenceInformation.SequenceNumber.Value = seqNo
   '         End If
         
            
         .RFF.Reference.ReferenceQualifier.Value = "ASL"
         .RFF.Reference.ReferenceNumber.Value = seqId
         
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
                             "SPC"
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
   eClass.CurrentProcedure = "Edifact_NHS003.SG18_Battery"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
