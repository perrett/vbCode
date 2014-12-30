VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASTM_1238"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msgAstm As astm1238.MsgASTM1238
Private msgErr As astm1238.ErrorCollection
Private LineCount As Long

Public Function CreateMessage(ReportId As Long) As String
   On Local Error GoTo procEH
   Dim I As Integer
   Dim msgBuf As String
   
   Set msgAstm = CreateObject("astm1238.MsgASTM1238") ' New astm1238.MsgASTM1238
   MapPatient ReportId
   MapOBR ReportId
   
   Set msgErr = msgAstm.Validate
   If msgErr.Count > 0 Then
      msgControl.ReportStatus = MS_PARSE_FAIL
      For I = 0 To msgErr.Count - 1
         msgControl.LogReportMessage MS_PARSE_FAIL, msgControl.CurrentReport & ": " & _
                                                    msgErr(I).Path & " - " & _
                                                    msgErr(I).Description
      Next I
   End If
   
   msgBuf = Replace(msgAstm.Parse, "\`", "'")
'   msgData.EscapeData msgBuf, True
   msgControl.LineCount = msgData.LineCount(msgBuf)
   
   Set msgAstm = Nothing
   CreateMessage = msgBuf
   
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "ASTM_1238.CreateMessage"
   eClass.Add Err.Number, Err.Description, Err.Source

End Function

Private Sub MapPatient(RepId As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim pLTS As ADODB.Parameter
   Dim gpName As String
   Dim gpCode As String
   Dim strRemain As String
   Dim lngPos As Long
   Dim hpData As New clsHealthParties
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Patient"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RepId)
      Set pLTS = .CreateParameter("Index", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      Set RS = .Execute
   End With
   
   With msgAstm.P
      .SetID.Value = msgControl.MessageId
      
      If msgControl.AnonymizeData Then
         .ServiceAssignedID.Value = "9999999999"
         .NHSNumber.Value = "9999999999"
         
         With .PatientName
            .FamilyName.Value = "Testpatient"
            .GivenName.Value = "ASTM"
            .MiddleInitialOrName.Value = "1238"
            .Prefix.Value = "Mr"
         End With
         
         With .DateOfBirth
            .Year.Value = "1900"
            .Month.Value = "01"
            .Day.Value = "01"
         End With
         
         .Sex.Value = "M"
         
         With .PatientAddress
            .StreetAddress.Value = "Street Address"
            .Number.Value = "1001"
            .City.Value = "City"
            .County.Value = "County"
            .Country.Value = "Country"
            .PostCode.Value = "PO5 COD"
         End With
      Else
         If Left(RS!Hospital_Number, 3) <> "ICE" Then
            .ServiceAssignedID.Value = RS!Hospital_Number
         End If
         
         .NHSNumber.Value = RS!New_NHS_No & ""
         
         With .PatientName
            .FamilyName.Value = RS!Surname & ""
            .GivenName.Value = RS!Forename & ""
            .MiddleInitialOrName.Value = RS!Midname & ""
            .Prefix.Value = RS!Title & ""
         End With
         
         If RS!Date_Of_Birth & "" <> "" Then
            With .DateOfBirth
               .Year.Value = Format(RS!Date_Of_Birth, "yyyy")
               .Month.Value = Format(RS!Date_Of_Birth, "mm")
               .Day.Value = Format(RS!Date_Of_Birth, "dd")
            End With
         End If
         
         If RS!Sex = 1 Then
            .Sex.Value = "M"
         ElseIf RS!Sex = 2 Then
            .Sex.Value = "F"
         Else
            .Sex.Value = "W"
         End If
         
         With .PatientAddress
            .StreetAddress.Value = RS!Pat_Addr_Line1 & ""
            .Number.Value = RS!Pat_Addr_Line2 & ""
            .City.Value = RS!Pat_Addr_Line3 & ""
            .County.Value = RS!Pat_Addr_Line4 & ""
            .Country.Value = RS!Pat_Addr_Line5 & ""
            .PostCode.Value = RS!Pat_PostCode & ""
         End With
      End If
      
      RS.Close
      
'     Set RS = iceCmd.Execute
      
      hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
      hpData.LTSIndex = repListRS!EDI_LTS_Index
      hpData.Read RepId, False
      
      gpName = Replace(hpData.HP902Name, "`", "'") '  We need to re-instate apostophes for ASTM
      gpCode = hpData.HP902Code
      
      With .AttendingPhysicianID.ClinicianName
         strRemain = Replace(gpName, ",", ".")
         strRemain = Replace(strRemain, " ", ".")
         strRemain = Replace(strRemain, "..", ".")
         lngPos = InStrRev(strRemain, ".")
         If lngPos > 0 Then
            .FamilyName.Value = Mid(strRemain, lngPos + 1)
            strRemain = Left(strRemain, lngPos - 1)
            lngPos = InStr(strRemain, ".")
            If lngPos > 0 Then
               If UCase(Left(strRemain, lngPos - 1)) = "DR" Or _
                  UCase(Left(strRemain, lngPos - 1)) = "MR" Then
                  .Prefix.Value = Left(strRemain, lngPos - 1)
                  strRemain = Mid(strRemain, lngPos + 1)
               End If
               lngPos = InStr(strRemain, ".")
               If lngPos > 0 Then
                  .GivenName.Value = Left(strRemain, lngPos - 1)
                  .MiddleInitialOrName.Value = Mid(strRemain, lngPos + 1)
               Else
                  .GivenName.Value = strRemain
               End If
            Else
               .GivenName.Value = strRemain
            End If
         Else
            .FamilyName.Value = strRemain
         End If
      End With
      
      .AttendingPhysicianID.ClinicianCode.Value = gpCode
            
'      RS.Close
   End With
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "ASTM_1238.MapPatient"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function FileHeader() As String
   Dim strHeader As String
   
   strHeader = "H`{|}\`<DATE>``<SENDERID>`<ORG>`BHI```<RECEIVERID>``P`ASTM4.1`<DATE>`"
   strHeader = Replace(strHeader, "<DATE>", Format(Now(), "yyyymmddhhnn"))
   strHeader = Replace(strHeader, "<SENDERID>", msgControl.CurrentSender)
   strHeader = Replace(strHeader, "<RECEIVERID>", msgControl.CurrentReceiver)
   strHeader = Replace(strHeader, "<ORG>", msgControl.Organisation)
   FileHeader = strHeader & vbCrLf
End Function

Public Property Get FileTrailer() As String
   FileTrailer = "L`1``" & msgControl.MsgInBatch & "`" & msgControl.LineCount & "`" & vbCrLf
End Property

Public Property Get InvalidChars() As Variant
   ReDim strArray(2)
   strArray(0) = ""
   strArray(1) = ""
   InvalidChars = strArray
End Property

Public Property Get Notification() As Boolean
   Notification = False
End Property

Public Property Get ReleaseChar() As String
   ReleaseChar = "\"
End Property

Public Property Get RetainGrave() As Boolean
   RetainGrave = True
End Property

Public Property Get LinesInMessage() As Long
   LinesInMessage = LineCount
End Property

Private Sub MapOBR(RepId As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim sampRS As New ADODB.Recordset
   Dim invRS As New ADODB.Recordset
   Dim RS As New ADODB.Recordset
   Dim obrCmt As Integer
   Dim sampSeq As Long
   Dim I As Integer
   Dim strTemp As String
   Dim pos As Integer
   Dim vData As Variant
   Dim vFields As Variant
   Dim iPosn As Long
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_ASTMComments"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RepId)
      Set RS = .Execute
   End With
   
   If RS.RecordCount > 0 Then
      msgAstm.OBR_Group.Append
      With msgAstm.OBR_Group(0).C
         obrCmt = 0
         Do Until RS.EOF
            If Len(RS!Service_Report_Comment) > 2 Then
               .Append
               With .Item(obrCmt)
                  .SetID.Value = obrCmt + 1
                  .Source.Value = "L"
                  .CommentText.Value = RS!Service_Report_Comment
               End With
               obrCmt = obrCmt + 1
            End If
            RS.MoveNext
         Loop
         RS.Close
      End With
   End If
   
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Sample"
      Set sampRS = .Execute
   End With
   
   sampSeq = 0
   
   Do Until sampRS.EOF
      With iceCmd
         If blnUseRCIndex Then
            .CommandText = "ICELABCOMM_Report_Invest_By_Index"
         Else
            .CommandText = "ICELABCOMM_Report_Invest_By_Code"
         End If
         
         .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
         .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , sampRS!Sample_Index)
         .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
         Set invRS = .Execute
         .Parameters.Delete ("Abnormal")
         .Parameters.Delete ("SampId")
         .Parameters.Delete ("LTSIndex")
      End With
      
'      If IsNull(invRS!EDI_RC_Index) Then
'         vFields = Array(9, 10, 12, 13, 18, 19, 20, 21, 22, 23, 8)
'      Else
'         vFields = Array(14, 15, 16, 17, 24, 25, 26, 27, 28, 29, 8)
'      End If
'
'      iPosn = 0
'      vData = invRS.GetRows(-1, adBookmarkFirst, vFields)
'      invRS.MoveFirst
      
      Do Until invRS.EOF
         If invRS!EDI_OP_Suppress Then
'         If vData(2, iPosn) Then
            msgControl.LogReportMessage IS_INV_SUPPRESSED, Replace(invRS!EDI_Local_Test_Code, Chr(0), "") & _
                                        " Investigation suppressed - no output produced."
         Else
            If sampSeq > 0 Or msgAstm.OBR_Group.Count = 0 Then
               msgAstm.OBR_Group.Append
               obrCmt = 0
            End If
            
            With msgAstm.OBR_Group(sampSeq).OBR
               .SetID.Value = sampSeq + 1
               pos = InStr(1, sampRS!Service_Report_Id, " ")
               If pos = 0 Then
                  strTemp = sampRS!Service_Report_Id
               Else
                  strTemp = Left(sampRS!Service_Report_Id, pos - 1)
               End If
               
               .FillerOrderNumber.EntityID.Value = strTemp
               .FillerOrderNumber.ApplicationID.Value = msgControl.OrganisationCode
               
               If Trim(sampRS!Collection_DateTime & "") = "" Then
                  If Trim(sampRS!Collection_DateTimeReceived & "") = "" Then
                     .ObservationDateTime.Time.Value = Format(sampRS!DateTime_Of_Report, "yyyymmddHhNn")
                  Else
                     .ObservationDateTime.Time.Value = Format(sampRS!Collection_DateTimeReceived, "yyyymmddHhNn")
                  End If
               Else
                  .ObservationDateTime.Time.Value = Format(sampRS!Collection_DateTime, "yyyymmddHhNn")
               End If
               
               If blnASTMSameDate Then
                  .SpecimenReceivedDateTime.Time.Value = Format(sampRS!Collection_DateTime, "yyyymmddHhNn")
               Else
                  .SpecimenReceivedDateTime.Time.Value = Format(sampRS!Collection_DateTimeReceived, "yyyymmddHhNn")
               End If
               
               .SpecimenSource.SpecimenSourceNameOrCode.Identifier.Value = sampRS!Sample_Text
               .QuantityTiming.Append
               .QuantityTiming.Item(0).Priority.Value = "R"
               .ScheduledDateTime.Time.Value = Format(Now(), "yyyymmddHhNn")
               
               With .UniversalServiceIdentifier
                  If blnASTMLocal Then
                     If invRS!Investigation_Code & "" <> "" Then
'                     If vData(0, iPosn) & "" <> "" Then
                        .Identifier.Value = invRS!Investigation_Code
                        .Text.Value = invRS!Investigation_Requested 'vData(1, iPosn)
                        .NameOfCodingSystem.Value = "LC"
                     End If
                  
                     If Not IsNull(invRS!Read_V2RC) Then
                        .AlternateIdentifier.Value = invRS!Read_V2RC 'vData(4, iPosn)
                        .AlternateText.Value = invRS!Read_V2Rubric 'vData(5, iPosn)
                        .AlternateCodingSystem.Value = "RC"
                     End If
                  Else
                     If IsNull(invRS!Read_V2RC) Then
'                     If IsNull(vData(4, iPosn)) Then
                        If blnUpper Then
                           strTemp = UCase(Left(invRS!Investigation_Requested, 35))
                        Else
                           strTemp = Left(invRS!Investigation_Requested, 35)
                        End If
                        .Text.Value = strTemp
                     Else
                        .Identifier.Value = invRS!Read_V2RC 'vData(4, iPosn)
                        .Text.Value = invRS!Read_V2Rubric 'vData(5, iPosn)
                        .NameOfCodingSystem.Value = "RC"
                     End If
                  End If
               End With
      
               With iceCmd
                  .CommandText = "ICELABCOMM_Report_InvestComments"
                  .Parameters(0).Value = invRS!Investigation_Index
                  Set RS = .Execute
                  .Parameters(0).Value = RepId
               End With
               
               With msgAstm.OBR_Group(sampSeq)
                  Do Until RS.EOF
                     If Len(RS!Service_Investigation_Comment) > 2 Then
                        .C.Append
                        With .C(obrCmt)
                           .SetID.Value = obrCmt + 1
                           .Source.Value = "L"
                           .CommentText.Value = RS!Service_Investigation_Comment
                        End With
                        obrCmt = obrCmt + 1
                     End If
                     RS.MoveNext
                  Loop
               End With
               
               MapOBX sampSeq, invRS!Investigation_Index
            End With
            
            If msgAstm.OBR_Group(sampSeq).OBX_Group.Count = 0 Then
               msgControl.LogReportMessage IS_INV_SUPPRESSED, invRS!Investigation_Code & " - " & invRS!Investigation_Requested & _
                                           " ALL test suppressed for this investigation"
            Else
               sampSeq = sampSeq + 1
            End If
         End If
         
         invRS.MoveNext
         iPosn = iPosn + 1
      Loop

      invRS.Close
      sampRS.MoveNext
      iPosn = 0
   Loop
   
   If sampRS.RecordCount = 0 Then
      msgControl.ReportStatus = RS_DATA_INTEGRITY
      msgControl.LogReportMessage RS_DATA_INTEGRITY, "No sample data for this report"
   ElseIf sampSeq = 0 Then
      msgControl.ReportStatus = RS_SUPPRESSION
      msgControl.LogReportMessage RS_SUPPRESSION, "All the tests for this message have the output suppressed"
   End If
   
   sampRS.Close
   Set RS = Nothing
   Set invRS = Nothing
   Set sampRS = Nothing
   Set iceCmd = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "ASTM_1238.MapOBR"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub MapOBX(OBRCnt As Long, _
                   InvID As Long)
   On Local Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim resRS As New ADODB.Recordset
   Dim RS As New ADODB.Recordset
   Dim OBXcnt As Integer
   Dim strTemp As String
   Dim obxCmt As Integer
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
      
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , InvID)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      Set resRS = .Execute
      .Parameters.Delete ("LTSIndex")
   End With
      
   OBXcnt = 0
      
'   If IsNull(resRS!EDI_RC_Index) Then
''      vFields = Array(18, 19, 21, 22, 23, 30, 31, 32, 33, 34, 35, 36, 5)
'      vFields = Array(16, 17, 19, 20, 21, 28, 29, 30, 31, 32, 33, 34, 5)
'   Else
''      vFields = Array(22, 23, 24, 25, 26, 35, 36, 37, 38, 39, 40, 41, 5)
'      vFields = Array(22, 23, 24, 25, 26, 35, 36, 37, 38, 39, 40, 41, 5)
'   End If
'
'   vData = resRS.GetRows(-1, adBookmarkFirst, vFields)
   resRS.MoveFirst
'   rPosn = 0
   
   Do Until resRS.EOF
      testDesc = Replace(resRS!EDI_Local_Test_Code & "", Chr(0), "") & " (" & resRS!EDI_Local_Rubric & "" & ") "
      
'     Has the output for this test been supressed?
      If resRS!EDI_OP_Suppress Then
'      If vData(2, rPosn) Then
         msgControl.LogReportMessage IS_TEST_SUPPRESSED, testDesc & "Test suppressed"
      Else
         msgAstm.OBR_Group(OBRCnt).OBX_Group.Append
         With msgAstm.OBR_Group(OBRCnt).OBX_Group(OBXcnt)
            With .OBX
               .SetID.Value = OBXcnt + 1
               .ValueType.Value = "ST"
               With .ObservationIdentifier
               
                  If blnASTMLocal Then
                     .Identifier.Value = Trim(resRS!Result_Code & "") 'Trim(vData(0, rPosn) & "")
                     .Text.Value = Trim(resRS!Result_Rubric & "") 'Trim(vData(1, rPosn) & "")
                     .NameOfCodingSystem.Value = "LC"
                     
                     If Not IsNull(resRS!Read_V2RC) Then
'                     If Not IsNull(vData(5, rPosn)) Then
                        .AlternateIdentifier.Value = resRS!Read_V2RC 'vData(5, rPosn)
                        .AlternateText.Value = resRS!Read_V2Rubric 'vData(6, rPosn)
                        .AlternateCodingSystem.Value = "RC"
                     End If
                  
                  Else
                     If Not IsNull(resRS!Read_V2RC) Then
'                     If Not IsNull(vData(5, rPosn)) Then
                        .Identifier.Value = resRS!Read_V2RC 'vData(5, rPosn)
                        .Text.Value = resRS!Read_V2Rubric 'vData(6, rPosn)
                        .NameOfCodingSystem.Value = "RC"
                     Else
                        If blnUpper Then
                           .Text.Value = Left(UCase(resRS!Result_Rubric), 35)
                        Else
                           .Text.Value = Left(resRS!Result_Rubric, 35)
                        End If
                     End If
                  End If
               End With
               
               .ObservationValue.Append
               .ObservationValue.Item(0).Value = resRS!Result
               .Units.Identifier.Value = resRS!UOM_Code
               
               If Not IsNull(resRS!Lower_Range) Then
                  If resRS!Lower_Range = "" Then
                     strTemp = "<" & resRS!Upper_Range
                  ElseIf resRS!Upper_Range = "" Then
                     strTemp = ">" & resRS!Lower_Range
                  Else
                     strTemp = resRS!Lower_Range & "-" & resRS!Upper_Range
                  End If
               End If
         
               .ReferencesRange.Value = strTemp
               strTemp = ""
               
               If resRS!Abnormal_Flag Then
                  msgData.AdjustRangeData Trim(resRS!Lower_Range & ""), Trim(resRS!Upper_Range & "")
                  If UCase(resRS!Abnormal_Code) = "H" Or UCase(resRS!Abnormal_Code) = "L" Then
                     .AbnormalFlags.Value = resRS!Abnormal_Code
                  Else
                     .AbnormalFlags.Value = "*"
                  End If
                  
'                  If IsNumeric(resRS!Result) Then
''                     .AbnormalFlags.Append
'                     If Val(resRS!Result) > Val(msgData.HighRange) Then
'                        .AbnormalFlags.Value = "H"
'                     ElseIf Val(resRS!Result) < Val(msgData.LowRange) Then
'                        .AbnormalFlags.Value = "L"
'                     End If
'                  End If
               End If
            End With
            
            With .C
               If resRS!Comment_Marker Then
                  With iceCmd
                     .CommandText = "ICELABCOMM_Report_ResultComment"
                     .Parameters(0).Value = resRS!Result_Index
                     Set RS = .Execute
                  End With
                  
                  Do Until RS.EOF
                     If Len(RS!Service_Result_Comment) > 2 Then
                        .Append
                        With .Item(obxCmt)
                           .SetID.Value = obxCmt + 1
                           .Source.Value = "L"
                           .CommentText.Value = RS!Service_Result_Comment
                        End With
                        obxCmt = obxCmt + 1
                     End If
                     RS.MoveNext
                  Loop
                  obxCmt = 0
                  RS.Close
               End If
            
               Set RS = Nothing
               resRS.MoveNext
               rPosn = resRS.AbsolutePosition - 1
               OBXcnt = OBXcnt + 1
            End With
         End With
      End If
   Loop
   
   resRS.Close
   Set resRS = Nothing
   Set RS = Nothing
   Set iceCmd = Nothing
   
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "ASTM_1238.MapOBX"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

