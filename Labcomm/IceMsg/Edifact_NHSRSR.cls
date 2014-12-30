VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Edifact_NHSRSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msgEDI As nhsrsr.MsgNHSRSR
Private msgErr As nhsrsr.ErrorCollection

Private PatientAge As Integer

Public Property Get AllowApostrophe() As Boolean
   AllowApostrophe = True
End Property

Public Function CreateMessage(ReportId As Long) As Variant
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim msgBuf As String
   Dim I As Integer
   
   Set msgEDI = New nhsrsr.MsgNHSRSR
   With msgEDI
      With .UNH
         .MessageReferenceNumber.Value = msgControl.MessageId
         With .MessageIdentifier
            .MessageTypeIdentifier.Value = "NHSRSR"
            .MessageTypeVersionNumber.Value = "D"
            .MessageTypeReleaseNumber.Value = "1"
            .ControllingAgency.Value = "NH"
            .AssociationAssignedCode.Value = "NHS002"
         End With
      End With

      .BGM.DocumentMessageName.DocumentMessageNameCoded.Value = "RRR"
   
      With .DTM.DateTimePeriod
         .DateTimePeriodQualifier.Value = "137"
         .DateTimePeriod.Value = Format(Now(), "yyyymmddHhMm")
         .DateTimePeriodFormatQualifier.Value = "203"
      End With
   End With
   
   MapSG1 ReportId
   MapSG2 ReportId
   mapSG2_SG5 ReportId
   MapSG2_SG6 ReportId
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Sample"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   Do Until RS.EOF
      MapSG2_SG10 ReportId, RS!Sample_Index
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
   
   If msgControl.ReportStatus = 0 Then
      msgBuf = "Errors encountered when generating message'See comments for description"
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
      
      msgBuf = msgEDI.Parse
   
'      msgData.InvalidChars = ""
'      msgData.ReplaceChars = ""
'      msgData.SegmentCount msgBuf
'      msgData.EscapeData msgBuf
   End If
   
   CreateMessage = msgBuf
   Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHSRSR.CreateMessage"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function FileHeader() As String
   Dim strHeader As String
   Dim flags As String
   
   strHeader = "UNA:+.? '" & _
               "UNB+UNOB:2+<SENDINGTRADER>:80+<RECEIVINGTRADER>:80+<DATE:TIME>+<INTERCHANGE>++NHSRSR'"
   strHeader = Replace(strHeader, "<SENDINGTRADER>", msgControl.CurrentSender)
   strHeader = Replace(strHeader, "<RECEIVINGTRADER>", msgControl.CurrentReceiver)
   strHeader = Replace(strHeader, "<DATE:TIME>", Format(Now(), "YYMMDD:HHNN"))
   strHeader = Replace(strHeader, "<INTERCHANGE>", msgControl.CurrentInterchange)
   
   FileHeader = strHeader
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
   RetainGrave = (True And blnRetainGrave)
End Property

Private Sub MapSG1(ReportIndex As Long)
                     
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim clinNatCode As String
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_HealthParties"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportIndex)
      Set RS = .Execute
   End With
   
   With msgEDI.SG1
      .Append
      RS.Filter = "EDI_HP_Type = 901"
      With .Item(0).NAD
         .PartyQualifier.Value = "MR"
         .PartyIdentificationDetails.PartyIdIdentification.Value = repListRS!EDI_Loc_Nat_Code_To
         .PartyIdentificationDetails.CodeListQualifier.Value = "901"
      End With
      
      .Item(0).SPR.ServiceProviderQualifier.Value = "ORG"
                  
      .Append
      RS.Filter = "EDI_HP_Type = 902"
      
      With .Item(1)
         .NAD.PartyQualifier.Value = "PO"
         If IsNull(repListRS!GP_Code) Then
            If Trim(RS!EDI_OP_Name & "") = "" Or _
               Trim(RS!Clinician_Local_Code & "") = "" Then
               msgControl.ReportStatus = RS_DATA_INTEGRITY
               msgControl.LogReportMessage RS_DATA_INTEGRITY, "Clinician Surname or Local Id not present"
            Else
'              Clinician goes through "as is"
               .NAD.PartyName.PartyName1.Value = Trim(RS!EDI_OP_Name)
               .RFF.Append
               .RFF(0).Reference.ReferenceQualifier.Value = "AHI"
               .RFF(0).Reference.ReferenceNumber.Value = RS!Clinician_Local_Code
            End If
         Else
            .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = repListRS!GP_Code
            .NAD.PartyName.PartyName1.Value = repListRS!EDI_OP_Name
            .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
         End If
         .SPR.ServiceProviderQualifier.Value = "PRO"
      End With
                  
      RS.Filter = ""
      RS.Filter = "EDI_HP_Type = 906"
      If RS.EOF Then
         RS.Filter = "EDI_HP_Type = 902"
         clinNatCode = Trim(repListRS!GP_Code & "")
      End If

'     Message Recipient - GP Individual
      .Append
      With .Item(2)
         .NAD.PartyQualifier.Value = "MR"
         If IsNull(RS!EDI_OP_Name) Then
            If Trim(RS!EDI_OP_Name & "") = "" Or _
               Trim(RS!Clinician_Local_Code & "") = "" Then
               msgControl.ReportStatus = RS_DATA_INTEGRITY
               msgControl.LogReportMessage RS_DATA_INTEGRITY, "Clinician Surname or Local Id not present"
            Else
               .NAD.PartyName.PartyName1.Value = Trim(RS!EDI_OP_Name & "")
               .RFF.Append
               .RFF(0).Reference.ReferenceQualifier.Value = "AHI"
               .RFF(0).Reference.ReferenceNumber.Value = RS!Clinician_Local_Code
            End If
         Else
            .NAD.PartyName.PartyName1.Value = RS!EDI_OP_Name
            .NAD.PartyIdentificationDetails.PartyIdIdentification.Value = RS!EDI_NatCode
            .NAD.PartyIdentificationDetails.CodeListQualifier.Value = "900"
         End If
         .SPR.ServiceProviderQualifier.Value = "PRO"
      End With
                  
      .Append
      With .Item(3)
         .NAD.PartyQualifier.Value = "SLA"
         .SPR.ServiceProviderQualifier.Value = "ORG"
         .NAD.PartyName.PartyName1.Value = msgControl.Organisation
      End With
                  
      .Append
      With .Item(4)
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
   eClass.CurrentProcedure = "Edifact_NHSRSR.MapSG1"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG2(ReportId As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim rTime As String
   
   With msgEDI.SG2
      .S02.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "02"
      .RFF.Reference.ReferenceQualifier.Value = "SRI"
      .RFF.Reference.ReferenceNumber.Value = repListRS!Report_Identifier
      With .DTM.DateTimePeriod
         .DateTimePeriodQualifier.Value = "ISR"
         
         rTime = Trim(repListRS!Report_Date)
         If Len(rTime) <= 10 Then
            .DateTimePeriodFormatQualifier.Value = "102"
            .DateTimePeriod.Value = Format(rTime, "yyyymmdd")
         Else
            .DateTimePeriodFormatQualifier.Value = "203"
            .DateTimePeriod.Value = Format(rTime, "yyyymmddhHnN")
         End If
      End With
      
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICELABCOMM_Report_Comments"
         .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
         Set RS = .Execute
      End With
      
      RS.Filter = "Comment_Type = 'R'"
      If RS.RecordCount > 0 Then
         msgData.FreeText msgEDI.SG2, _
                          RS, _
                          "SPC", _
                          0, _
                          "NHSRSR", _
                          9
      End If
      RS.Close
      
   End With
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHSRSR.MapSG2"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub mapSG2_SG5(RepId As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Patient"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RepId)
      Set RS = .Execute
   End With
   
   If Trim(RS!Date_Of_Birth & "") <> "" Then
      PatientAge = DateDiff("yyyy", RS!Date_Of_Birth, Now())
   Else
      PatientAge = 0
   End If
   
   With msgEDI.SG2.SG5
      .S05.SegmentGroupUsageDetails.SegmentGroupUsageCoded.Value = "05"
      .PNA.Append
      .PNA(0).PartyQualifier.Value = "PAT"
      If Trim(RS!New_NHS_No & "") <> "" Then
         .PNA(0).IdentificationNumber.IdentityNumberQualifier.Value = "OPI"
         .PNA(0).IdentificationNumber.IdentityNumber.Value = RS!New_NHS_No
      Else
         If Left(RS!Hospital_Number, 3) <> "ICE" Then
            .RFF.Append
            .RFF(0).Reference.ReferenceQualifier.Value = "RPI"
            .RFF(0).Reference.ReferenceNumber.Value = RS!Hospital_Number
         End If
      End If
      
      With .PNA(0)
         .NameTypeCoded.Value = "US"
         .NameComponentDetails1.NameComponentQualifier.Value = "SU"
         .NameComponentDetails1.NameComponent.Value = RS!Surname
         If Trim(RS!Forename & "") <> "" Then
            .NameComponentDetails2.NameComponentQualifier.Value = "FO"
            .NameComponentDetails2.NameComponent.Value = RS!Forename
         End If
         If Trim(RS!Midname & "") <> "" Then
            .NameComponentDetails2.NameComponentQualifier.Value = "MI"
            .NameComponentDetails2.NameComponent.Value = RS!Midname
         End If
         If Trim(RS!Title & "") <> "" Then
            .NameComponentDetails2.NameComponentQualifier.Value = "TI"
            .NameComponentDetails2.NameComponent.Value = RS!Title
         End If
      End With
      
'      If ahReport.Patient.TelNo <> "" Then
'         .COM.Append
'         .COM(0).CommunicationContact.CommunicationNumber.Value = ahReport.Patient.TelNo
'         .COM(0).CommunicationContact.CommunicationChannelQualifier.Value = "TE"
'      End If
      
      If Trim(RS!Date_Of_Birth & "") <> "" Then
         With .DTM.DateTimePeriod
            .DateTimePeriodQualifier.Value = "329"
            .DateTimePeriod.Value = Format(RS!Date_Of_Birth, "yyyymmdd")
            .DateTimePeriodFormatQualifier.Value = "102"
         End With
      End If
            
      .PDI.SexCoded.Value = RS!Sex
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
   eClass.CurrentProcedure = "Edifact_NHSRSR.MapSG2_SG5"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
   
Private Sub MapSG2_SG6(ReportId As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   
   With msgEDI.SG2.SG5
      .SG6.Append
      With .SG6(0)
         With iceCmd
            .ActiveConnection = iceCon
            .CommandType = adCmdStoredProc
            .CommandText = "ICELABCOMM_Report_Comments"
            .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
            Set RS = .Execute
         End With
         
         RS.Filter = "Comment_Type = 'R'"
         If RS.RecordCount > 0 Then
            .S06.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "06"
            msgData.FreeText msgEDI.SG2.SG5.SG6(0), _
                             RS, _
                             "CID", _
                             0, _
                             "NHSRSR", _
                             9
         End If
         
         RS.Close
      End With
   End With
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHSRSR.MapSG2_SG6"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapSG2_SG10(ReportId As Long, SampId As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim invRS As New ADODB.Recordset
   Dim resRS As New ADODB.Recordset
   Dim RS As New ADODB.Recordset
   Dim invTot As Integer
   Dim vData As Variant
   Dim vFields As Variant
   Dim rPosn As Integer
   Dim testDesc As String
   Dim rcIndex As Long
   
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Investigation"
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , SampId)
      .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
      Set invRS = .Execute
      .Parameters.Delete (2)
      .Parameters.Delete (1)
   End With
   
   Do Until invRS.EOF
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICELABCOMM_Report_Results"
         .Parameters(0).Value = invRS!Investigation_Index
         Set resRS = .Execute
      End With
      
      If IsNull(resRS!EDI_RC_Index) Then
         vFields = Array(16, 17, 19, 20, 21, 28, 29, 30, 31, 32, 33, 34, 5)
      Else
         vFields = Array(22, 23, 24, 25, 26, 35, 36, 37, 38, 39, 40, 41, 5)
      End If
      
      vData = resRS.GetRows(-1, adBookmarkFirst, vFields)
      resRS.MoveFirst
      rPosn = 0
      
      With msgEDI.SG2.SG5
         Do Until resRS.EOF
            .SG10.Append
            With .SG10(invTot)
               .S10.SegmentGroupDetails.SegmentGroupUsageCoded.Value = "10"
               With .INV
                  .InvestigationCharacteristicQualifier.Value = "RS"
                  
'                 If a read code exists, use it, otherwise use the rubric
                  If IsNull(vData(5, rPosn)) Then
                     .InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = resRS!Result_Rubric
                  Else
                     .InvestigationCharacteristicDetails.InvestigationCharacteristicIdentification.Value = vData(5, rPosn)
                     .InvestigationCharacteristicDetails.CodeListQualifier.Value = vData(7, rPosn)
                     .InvestigationCharacteristicDetails.InvestigationCharacteristic.Value = vData(6, rPosn)
                  End If
               End With
               
               If invRS!Comment_Marker Then
                  With iceCmd
                     .ActiveConnection = iceCon
                     .CommandType = adCmdStoredProc
                     .CommandText = "ICELABCOMM_Report_InvestComments"
                     .Parameters(0).Value = invRS!Investigation_Index
                     Set RS = .Execute
                  End With
               
                  msgData.FreeText msgEDI.SG2.SG5.SG10(invTot), _
                                   RS, _
                                   "RID", _
                                   , _
                                   "NHSRSR", _
                                   9
                  RS.Close
               End If
               
               With msgEDI.SG2.SG5.SG10(invTot).SG11
                  
                  .S11.SegmentGroupUsageDetails.SegmentGroupUsageCoded.Value = "11"
                  If resRS!Result <> "" Then
                     .FTX.Append
                     .FTX.Item(.FTX.Count - 1).TextSubjectQualifier.Value = "RIT"
                     .FTX.Item(.FTX.Count - 1).TextLiteral.FreeText1.Value = resRS!Result
                  End If
                  
'                 Now check for any result comments
                  If resRS!Comment_Marker Then
                     With iceCmd
                        .CommandText = "ICELABCOMM_Report_ResultComment"
                        .Parameters(0).Value = resRS!Result_Index
                        Set RS = .Execute
                     End With
                  
                     msgData.FreeText msgEDI.SG2.SG5.SG10(invTot).SG11, _
                                      RS, _
                                      "IVC", _
                                      0, _
                                      "NHSRSR"
                     RS.Close
                  End If
                  
                  If Trim(repListRS!Report_Date & "") <> "" Then
                     .DTM.Append
                     .DTM(0).DateTimePeriod.DateTimePeriodFormatQualifier.Value = "RIV"
                     .DTM(0).DateTimePeriod.DateTimePeriod.Value = Format(repListRS!Report_Date, "yyyymmdd")
                     .DTM(0).DateTimePeriod.DateTimePeriodQualifier.Value = "102"
                  End If
               End With
            End With
            invTot = invTot + 1
            resRS.MoveNext
            rPosn = resRS.AbsolutePosition - 1
         Loop
         resRS.Close
      End With
      
      invRS.MoveNext
   Loop
   
   invRS.Close
   Set invRS = Nothing
   Set iceCmd = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_NHSRSR.MapSG2_SG10"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get Notification() As Boolean
   Notification = False
End Property

