VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IceMsgControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  A class module to store details on AHSLMessaging objects. The ICEMsgData class
'  holds a reference to individual messages, the file name and number to be written
'  and other sundry details. This allows as many messages as required to be 'open'
'  at any time.

'  Each message may consist of several reports. The class IceMsgReps holds details
'  of these.

Private orgId As String
Private trustCode As String
'Private orgCode As String
Private PracticeName As String
Private natCode As String
Private ReceiverID As String
Private SenderID As String
Private blnLinkInterchange As Boolean
Private msgRefIndex As Long
Private Specialty As String
Private MsgFormat As String
Private Interchange As Long
'Private LTIndex As Long
Private ConfNatCode As String
Private confTotal As Long
Private ConfTarget As Long
Private AckRequired As Boolean
Private SendTestData As Boolean
Private MsgStatus As enumMessageStatus
Private errStatus As enumMessageStatus
Private msgId As Long
Private RepId As RepData
Private repStatus As Long
Private msgFormatIndex As Integer
Private msgCount As Long
Private errCount As Long
Private objMsg As Object
'Private edi2 As Edifact_NHS002
'Private edi3 As Edifact_NHS003
Private html As clsHTML
Private errPath As String
Private histPath As String
Private pendPath As String
Private fileBase As String
Private fExt As String
Private curFile As String
Private outFile As String
Private errFile As String

Private ackNatCode As String
Private ackMsgCtrl As RepData
Private ackIC As String

Private Anonymize As Boolean
Private confFile As String
Private histFile As String
Private errPend As String
Private strArray() As String
Private astmLines As Long
Private ackErrors As Long
Private copyReportTo As String

Private colReps As New Collection
'Private colRT As New Collection
'Private fileNo As Integer
'Private KeyId As String
'Private TraderRef As String
'Private ControlRef As String
'Private fPath As String
'Private blnWriteHeader As Boolean
'Private confTotal() As Integer
'Private confMax As Integer
'Private confData() As Conformance
'Private smtpAddress As String

Private colData As New Collection
Private colConf As New Collection

Public Sub AckAccepted()
   MsgStatus = (MsgStatus Or TF_ACCEPT_ACK)
'   ReportStatus = MsgStatus
End Sub

Public Sub AckRejected(Partial As Boolean)
   MsgStatus = MsgStatus And TF_SET_ERROR
   MsgStatus = (MsgStatus Or MS_ACK_RECEIVED)
   If Partial Then
      MsgStatus = (MsgStatus Or MS_ACK_REJECT_PART)
   Else
      MsgStatus = (MsgStatus Or MS_ACK_REJECT_ALL)
   End If
   ReportStatus = MsgStatus
End Sub

Public Property Let AckStatus(lngNewValue As enumMessageStatus)
'  For Acks, ackmsgctrl points to the current ack header. It is set up when calling SetReport
   ackMsgCtrl.MessageImpExp = lngNewValue
End Property

Public Function AddImpExpHeader(ByRef Fname As String, _
                                rCount As Long, _
                                Trader As String, _
                                InterchangeNo As Long, _
                                dateTimeStamp As String, _
                                vErr As Long, _
                                AckStatus As Integer, _
                                rcErrors As Boolean, _
                                HeaderFlag As Long, _
                                Optional natCode As String = "") As String

'  Add new Service_ImpExp_Header record & return the reference
   On Error GoTo procEH
   Dim statId As Integer
   Dim iceCmd As New ADODB.Command
   Dim impExpRef As Long
   
   If InterchangeNo > 0 Then
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
         .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , msgRefIndex)
         .Parameters.Append .CreateParameter("IVal", adInteger, adParamInput, , InterchangeNo)
         
         If blnLinkInterchange Then
            .CommandText = "ICEMSG_Update_Interchange"
         Else
            .CommandText = "ICEMSG_Update_Interchange_ByMsgType"
            .Parameters.Append .CreateParameter("MsgType", adVarChar, adParamInput, 16, MsgFormat)
         End If
         
         .Execute
         If .Parameters("Return") <> 0 Then
            Err.Raise 3265, "IceMsg", .CommandText
         End If
      End With
      
      Set iceCmd = Nothing
   End If
   
   If vErr > 0 Then
      statId = 1
   Else
      statId = AckStatus
   End If
   
   eClass.FurtherInfo = "ICEMSG_Ins_Service_ImpExp_Headers"
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_Ins_Service_ImpExp_Headers"
      .Parameters.Append .CreateParameter("recId", adInteger, adParamReturnValue)
      .Parameters.Append .CreateParameter("Organisation", adVarChar, adParamInput, 6, trustCode)
      .Parameters.Append .CreateParameter("Trader_Code", adVarChar, adParamInput, 15, Trader)
      .Parameters.Append .CreateParameter("ImpExp", adChar, adParamInput, 1, "O")
      .Parameters.Append .CreateParameter("ImpExp_File", adVarChar, adParamInput, 120, Fname)
      .Parameters.Append .CreateParameter("Provider_Id", adInteger, adParamInput, , 1)
      .Parameters.Append .CreateParameter("Service_Type", adInteger, adParamInput, , 2)
      .Parameters.Append .CreateParameter("Messages", adInteger, adParamInput, , rCount)
      .Parameters.Append .CreateParameter("Date_Added", adDate, adParamInput, , dateTimeStamp)
      .Parameters.Append .CreateParameter("Control_Ref", adVarChar, adParamInput, 14, Interchange)
      .Parameters.Append .CreateParameter("Errors_Found", adInteger, adParamInput, , vErr)
      .Parameters.Append .CreateParameter("Comment_Marker", adBoolean, adParamInput, , True)
      .Parameters.Append .CreateParameter("Status_Flag", adInteger, adParamInput, , statId)
      .Parameters.Append .CreateParameter("EDI_LTS_Index", adInteger, adParamInput, , IIf(ltIndex = -1, Null, ltIndex))
      .Parameters.Append .CreateParameter("Warning_Flag", adBoolean, adParamInput, , rcErrors)
      .Parameters.Append .CreateParameter("HeaderFlag", adInteger, adParamInput, , HeaderFlag)
      .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 10, natCode)
      .Parameters.Append .CreateParameter("FormatIndex", adInteger, adParamInput, , msgFormatIndex)
      .Execute
      
      Fname = Replace(Fname, "@@@@", Format(.Parameters("RecId").Value, "000000000#"))
      iceCon.Execute "UPDATE Service_ImpExp_Headers SET " & _
                        "ImpExp_File = '" & Fname & "' " & _
                     "WHERE Service_ImpExp_Id = " & .Parameters("RecId").Value
         
   End With
   impExpRef = iceCmd.Parameters("RecId")
   Set iceCmd = Nothing
   
   If InterchangeNo > 0 Then
      If ConfNatCode <> "" Then
         If ConfTarget <= 0 Then
            msgData.AddImpExpComment impExpRef, "Interchange selected for conformance testing.", dateTimeStamp, RS_CONFORMANCE
            confTotal = 0
         Else
            confTotal = confTotal + 1
         End If
         
         With iceCmd
            .ActiveConnection = iceCon
            .CommandType = adCmdStoredProc
            .CommandText = "ICEMSG_Update_Conformance"
            .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
            .Parameters.Append .CreateParameter("LTIndex", adInteger, adParamInput, , ltIndex)
            .Parameters.Append .CreateParameter("NewVal", adInteger, adParamInput, , confTotal)
            .Execute
            If .Parameters("Return") <> 0 Then
               Err.Raise 3265, "IceMsg", .CommandText
            End If
         End With
         Set iceCmd = Nothing
      End If
   End If
   AddImpExpHeader = impExpRef
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsgControl.AddImpExpHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

'Private Function AddImpExpMessage(ReportId As String, _
'                             HeaderRef As Long, _
'                             Discipline As String, _
'                             dateTimeStamp As String, _
'                             MsgStat As Long, _
'                             Optional MsgComment As String = "", _
'                             Optional reject As Long = 0, _
'                             Optional fileId As String = "", _
'                             Optional MsgPosn As Long)
'
''  Write new Service_ImpExp_Message for all affected messages
'
'   Dim strSQL As String
'   Dim RS As New ADODB.Recordset
'   Dim iceCmd As New ADODB.Command
'   Dim svcMessId As Long
'   Dim patLocId As String
'   Dim patName As String
'   Dim svcId As String
'   Dim patSex As Long
'   Dim patDOB As Variant
'   Dim svcDate As Variant
'   Dim dest As String
'   Dim svcRepInd As Long
'   Dim repID As String
'   Dim indTo As Long
'   Dim rejectCnt As Long
'   Dim Fname As String
'
''  Select the original message
'   strSQL = "SELECT * " & _
'            "FROM Service_ImpExp_Messages " & _
'            "WHERE Service_Report_Index = '" & ReportId & "'"
'   RS.Open strSQL, ICEcon, adOpenKeyset, adLockReadOnly
'
'   If LogLevel = 2 Then
'      If RS.BOF And RS.EOF Then
'         eClass.LogMessage "No records found"
'      Else
'         eClass.LogMessage RS.RecordCount & " records found"
'      End If
'   End If
'
'   eClass.FurtherInfo = "ICEMSG_Ins_Service_ImpExp_Message"
'
'   With iceCmd
'      .ActiveConnection = ICEcon
'      .CommandType = adCmdStoredProc
'      .CommandText = "ICEMSG_Ins_Service_ImpExp_Message"
'      .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
'      .Parameters.Append .CreateParameter("Service_ImpExp_Id", adInteger, adParamInput, , HeaderRef)
'      .Parameters.Append .CreateParameter("Service_Message_Id", adVarChar, adParamInput, 50, RS!Service_Message_Id)
'      .Parameters.Append .CreateParameter("Patient_Local_Id", adVarChar, adParamInput, 30, RS!Patient_Local_ID)
'      .Parameters.Append .CreateParameter("Service_Id", adVarChar, adParamInput, 40, RS!Service_Id)
'      .Parameters.Append .CreateParameter("Patient_Name", adVarChar, adParamInput, 70, RS!Patient_Name)
'      .Parameters.Append .CreateParameter("Patient_Sex", adInteger, adParamInput, , RS!Patient_Sex)
'      .Parameters.Append .CreateParameter("Patient_DOB", adDate, adParamInput, , RS!Patient_DOB)
'      .Parameters.Append .CreateParameter("Service_Date_Of_Service", adDate, adParamInput, , RS!Service_Date_Of_Service)
'      .Parameters.Append .CreateParameter("Service_Message_Comment", adVarChar, adParamInput, 120, IIf(MsgComment = "", RS!Service_Message_Comment, MsgComment))
'      .Parameters.Append .CreateParameter("Rejected_Count", adInteger, adParamInput, , IIf(reject = 0, RS!Rejected_Count, reject))
'      .Parameters.Append .CreateParameter("Rejected_File", adVarChar, adParamInput, 128, IIf(fileId = "", RS!Rejected_File, fileId))
'      .Parameters.Append .CreateParameter("Date_Added", adDate, adParamInput, , dateTimeStamp)
'      .Parameters.Append .CreateParameter("Destination", adVarChar, adParamInput, 30, RS!Destination)
'      .Parameters.Append .CreateParameter("Service_Report_Index", adInteger, adParamInput, , RS!Service_Report_Index)
'      .Parameters.Append .CreateParameter("Discipline", adInteger, adParamInput, , Discipline)
'      .Parameters.Append .CreateParameter("To_EDI", adBoolean, adParamInput, , True)
'      .Parameters.Append .CreateParameter("MsgPosn", adInteger, adParamInput, , MsgPosn)
'      .Parameters.Append .CreateParameter("Report_Identifier", adVarChar, adParamInput, 35, RS!Report_Identifier)
'      .Parameters.Append .CreateParameter("EDI_Individual_Index_To", adInteger, adParamInput, , RS!EDI_Individual_Index_To)
'      .Parameters.Append .CreateParameter("MsgStat", adInteger, adParamInput, , MsgStat)
'      .Execute
'      AddImpExpMessage = .Parameter("Return")
'   End With
'   RS.Close
'   Set RS = Nothing
'End Function

Public Property Get AnonymizeData() As Boolean
   AnonymizeData = Anonymize
End Property

Public Sub CreateMessage(ReportIndex As Long)
   On Error GoTo procEH
   Dim vData As String
   Dim fileBuf As String
   Dim buf As String
   Dim htmlCopy As clsHTML
   Dim blnStripCRLF As Boolean
   Dim I As Integer
   
   DoEvents
   If IsNull(repListRS!Specialty_Active) Then
      RepId.OrStatus RS_DATA_INTEGRITY
     LogReportMessage RS_DATA_INTEGRITY, "Unable to determine Specialty"
   End If
   
   If IsNull(repListRS!Msg_Active) Then
      RepId.OrStatus RS_DATA_INTEGRITY
      LogReportMessage RS_DATA_INTEGRITY, "Unable to determine Message format"
   End If
   
   If IsNull(repListRS!service_Report_Type) Then
      RepId.OrStatus RS_DATA_INTEGRITY
      LogReportMessage RS_DATA_INTEGRITY, "No corresponding Service Report. Has the original been tidied?"
   End If
   
   If RepId.Status = 0 Then
      If repListRS!Specialty_Active Then
         If repListRS!Msg_Active Then
            If blnOverride Then
'              Override message settings using ini file values
               If GlobalMsgFormat <> "" Then
                  MsgFormat = GlobalMsgFormat
               End If
               SendTestData = blnGlobalTest
               Anonymize = blnGlobalAnonymize
            End If
            
            MsgStatus = MsgStatus Or MS_MSGOK
            vData = objMsg.CreateMessage(ReportIndex)
            
         Else
            vData = "Message Type not active (" & repListRS!EDI_Msg_Format & ")'"
            RepId.OrStatus MS_DATA_INTEGRITY
            msgControl.LogReportMessage MS_MSGOK, "Message Type not active (" & repListRS!EDI_Msg_Format & ")"
            eClass.LogMessage repListRS!EDI_Msg_Format & " is inactive", , "Error"
         End If
      Else
         vData = "Specialty not active (" & repListRS!service_Report_Type & ")'"
         RepId.OrStatus MS_DATA_INTEGRITY
         msgControl.LogReportMessage MS_DATA_INTEGRITY, "Specialty not active (" & repListRS!service_Report_Type & ")"
         eClass.LogMessage repListRS!service_Report_Type & " is inactive", , "Error"
      End If
   Else
      vData = "Unable to determine Specialty'"
      RepId.OrStatus MS_DATA_INTEGRITY
      LogReportMessage MS_DATA_INTEGRITY, "Unable to determine Specialty"
      eClass.LogMessage "Unable to determine Specialty", , "Error"
   End If
   
VerifyMessage:
   DoEvents
   
   If RepId.Status > 0 Then
      vData = ">>> Report:" & ReportName & " (Report_Index: " & RepId.EDIIndex & ")'***'" & vData
      RepId.OrStatus MS_NO_OUTPUT
      iceErr.Add vData, RepId
      eClass.LogMessage vData, , "Error"
   Else

'     Any copies in HTML format?
      If repListRS!service_Report_Type <> "002" Then
         If Not (IsNull(repListRS!EDI_SMTP_Active)) Then
            If repListRS!EDI_SMTP_Active Then
               Set htmlCopy = New clsHTML
               strArray = Split(repListRS!GP_SMTP, ";")
               
               For I = 0 To UBound(strArray)
                  htmlCopy.SendHTMLCopy ReportIndex, strArray(I)
                  eClass.LogMessage "HTML copy sent to: " & strArray(I)
               Next I
               
               Set htmlCopy = Nothing
            End If
         End If
      End If
      
      If (MsgStatus And MS_CONFORMANCE) = MS_CONFORMANCE Then
         RepId.OrStatus RS_CONFORMANCE
         LogReportMessage RS_CONFORMANCE, "Report selected for conformance testing"
         eClass.LogMessage RepId.ReportName & ": Selected for conformance testing"
      End If
      
      msgCount = msgCount + 1
      RepId.MessageImpExp = msgCount
      colReps.Add RepId
      
'     Detect any invalid characters in the message and add those found to the invalid list
      msgData.InvalidChars = objMsg.InvalidChars
      msgData.RecordInvalid vData
      
'     The 'RetainGrave' property is either set or ANDed with the ini file setting. True AND True = true, any other
'     combination = False. ANDing with the inifile setting gives control over replacement.
      vData = msgData.EscapeData(vData, objMsg.ReleaseChar, objMsg.RetainGrave)
      
      Put #1, , vData
   End If
   DoEvents
   Set RepId = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
   If Err.Number = 3263 Then  '  Database inconsistencies
      vData = "Data Integrity Error - See Tracking and/or log for further information'"
      RepId.OrStatus MS_DATA_INTEGRITY
      RepId.LogData(MS_DATA_INTEGRITY) = eClass.LogError
      eClass.ClearErrors
      Resume VerifyMessage
      
   ElseIf Err.Number = 3234 Then '  db Timeout
      Set RepId = Nothing
      eClass.LogMessage "Timeout - Report left on replist"
      Exit Sub
      
   Else
      eClass.CurrentProcedure = "IceMsgControl.CreateMessage"
      eClass.Add Err.Number, Err.Description, Err.Source
   End If
End Sub

Public Property Get CurrentReport() As String
   CurrentReport = RepId.ReportName
End Property

Public Property Let ErrorReport(strNewValue As String)
   RepId.ErrorReport = strNewValue
End Property

Public Property Let errorStatus(lngNewValue As enumMessageStatus)
   errStatus = errStatus Or lngNewValue
End Property

Public Property Get CurrentAcks() As Boolean
   CurrentAcks = AckRequired
End Property

Public Property Get CurrentExtension() As String
   If Len(fExt) = 5 Then
      CurrentExtension = "." & UCase(Mid(fExt, 3))
   Else
      CurrentExtension = UCase(fExt)
   End If
End Property

Public Property Get CurrentInterchange() As Long
   CurrentInterchange = Interchange
End Property

Public Property Get CurrentPractice() As String
   CurrentPractice = PracticeName
End Property

Public Property Get CurrentReceiver() As String
   CurrentReceiver = ReceiverID
End Property

Public Property Get CurrentSender()
   CurrentSender = SenderID
End Property

Public Property Let FormatIndex(iNewVal As Integer)
   msgFormatIndex = iNewVal
End Property

Public Property Get FormatIndex() As Integer
   FormatIndex = msgFormatIndex
End Property

Public Property Let InvResStatus(sValue As enumInvResStatus)
   RepId.OrStatus sValue
End Property

Public Property Get InterchangeErrorName() As String
   InterchangeErrorName = ackNatCode & "_" & ackIC
End Property

Public Property Get InvResStatus() As enumInvResStatus
   InvResStatus = RepId.Status
End Property

Public Property Let LineCount(lngNewValue As Long)
   astmLines = astmLines + lngNewValue
End Property

Public Property Get LineCount() As Long
   LineCount = astmLines + 2
   astmLines = 0
End Property

Public Sub LogReportMessage(MsgStatus As Long, _
                            msgValue As String)
   RepId.LogData(MsgStatus) = msgValue
'   RepId.OrStatus MsgStatus
End Sub

Public Property Get MessageId() As Long
   MessageId = msgCount + 1
End Property

Public Property Let MessageStatus(sValue As enumMessageStatus)
   MsgStatus = MsgStatus Or sValue
'   If sValue < 0 Then
'      MsgStatus = MsgStatus Or sValue
'   Else
'      errStatus = errStatus Or sValue
'   End If
End Property

Public Property Get MessageStatus() As enumMessageStatus
   MessageStatus = MsgStatus
End Property

Public Property Get MsgInBatch() As Long
   MsgInBatch = msgCount
End Property

Public Sub NewFile(NationalCode As String, _
                   MessageFormat As String, _
                   Optional KeyId As String = "")
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim iRS As New ADODB.Recordset
   Dim baseName As String
   Dim fileHdr As String
   Dim buf As String
   Dim fileMsg As String
   Dim refIndex As Long
   
   DoEvents
   msgId = 1
   MsgStatus = MS_MSGOK
   errStatus = 0
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_MessageData"
      .Parameters.Append .CreateParameter("LTIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 6, repListRS!EDI_Loc_Nat_Code_To)
'      .Parameters.Append .CreateParameter("Format", adVarChar, adParamInput, 16, repListRS!EDI_Msg_Format)
      .Parameters.Append .CreateParameter("Spec", adVarChar, adParamInput, 6, repListRS!service_Report_Type)
      Set RS = .Execute
   End With
   
   If RS.EOF Then
      MsgStatus = MS_DATA_INTEGRITY
   Else
'     Is redirection in force for this message?
      If RS!EDI_Redirect_To <> "" Then
         If RS!EDI_Redirect_To <> NationalCode Then
'           Get the message details and trader code for the redirected account
            iceCmd.Parameters("NatCode").Value = RS!EDI_Redirect_To
            RS.Close
            Set RS = iceCmd.Execute
         End If
      End If
      
      blnUseRCIndex = RS!UseLabReadCodes
      
      msgRefIndex = RS!Ref_Index
      
      ' Check removed for new DTS enabled sites
      'If RS.Fields(14).Name = "File_Extension" Then
         fExt = RS!File_Extension & ""
      'End If
      
      If fExt = "" Then
      
'        Get the report details and set the file extension
         Select Case UCase(RS!EDI_Delivery_Method)
            Case "X"
'              Do we encrypt this message?
'               If RS!EDI_Encrypt_Enabled Then
''                 Yes, so is encryption enabled for this practice?
'                  If RS!EDI_Encryption > 0 Then
'                     fExt = "XEN"
'                  Else
'                     fExt = "XMS"
'                  End If
'               Else
                  fExt = "XMS"
'               End If
                              
            Case "H"
               fExt = "SND"
               
            Case "S"
               fExt = "SND"
            
            Case "L"
               fExt = "ICE"
         
         End Select
      End If
      
      If Not IsNull(RS!EDI_Hold_Output) Then
         If RS!EDI_Hold_Output Then
            fExt = "H" & fExt
         End If
      End If
      
      fExt = "." & fExt
      
'     Is this message to be copied to another recipient?
      If Left(Trim(RS!EDI_Copy_To & ""), 2) = "__" Then
         copyRecipient = RS!EDI_Copy_To
      Else
         copyRecipient = ""
      End If
      
      trustCode = Trim(RS!Organisation)
      orgId = Trim(RS!Organisation_Name)
      orgCode = Trim(RS!EDI_OrgCode)
      PracticeName = Trim(RS!EDI_Name)
      natCode = repListRS!EDI_Loc_Nat_Code_To
      refIndex = RS!Ref_Index
      AckRequired = RS!EDI_Acks_Active
      MsgFormat = RS!EDI_Msg_Format
      fileMsg = Replace(Mid(MsgFormat, InStr(1, MsgFormat, ",") + 1), ".", "")
      fileMsg = Replace(fileMsg, " ", "")
      
      SendTestData = RS!EDI_Msg_Test
      Anonymize = RS!EDI_Anonymize_Data
      
      ltIndex = RS!EDI_LTS_Index
      SenderID = RS!SenderID
      ReceiverID = RS!ReceiverID
      baseName = RS!EDI_Delivery_Method
      
      ConfNatCode = IIf(IsNull(RS!Conformance_Nat_Code), "", RS!Conformance_Nat_Code)
      ConfTarget = 0
      
      If ConfNatCode <> "" Then
         confTotal = IIf(IsNull(RS!Conformance_Total), 0, RS!Conformance_Total)
         ConfTarget = IIf(IsNull(RS!Conformance_Frequency), 500, RS!Conformance_Frequency) - confTotal
         If ConfTarget = 0 Then
            MsgStatus = MsgStatus Or MS_CONFORMANCE
         End If
      End If
      
      Set iceCmd = Nothing
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICEMSG_InterchangeData"
         .Parameters.Append .CreateParameter("RefIndex", adInteger, adParamInput, , refIndex)
         Set iRS = .Execute
      End With
      
      blnLinkInterchange = iRS!Link_Interchange_Nos
      ReceiverID = iRS!ReceiverID
      
      If blnLinkInterchange Then
         Interchange = Val(Trim(iRS!EDI_Last_Interchange & "")) + 1
      Else
         iRS.Filter = "EDI_Msg_Format = '" & MsgFormat & "'"
         
         If iRS.EOF = False Then
            Interchange = Val(Trim(iRS!EDI_Last_Interchange & "")) + 1
         Else
            Set iceCmd = Nothing
            With iceCmd
               .ActiveConnection = iceCon
               .CommandType = adCmdStoredProc
               .CommandText = "ICECONFIG_New_Interchangerecord"
               .Parameters.Append .CreateParameter("RefId", adInteger, adParamInput, , refIndex)
               .Parameters.Append .CreateParameter("ICVal", adInteger, adParamInput, , 0)
               .Parameters.Append .CreateParameter("MsgFormat", adVarChar, adParamInput, 16, MsgFormat)
               .Execute
            End With
            
            Interchange = 1
            
            eClass.LogMessage "New interchange record for Message Type: " & MsgFormat & " created"
         End If
      End If
      
      RS.Close
      iRS.Close
      
      If AckRequired Then
         MsgStatus = MS_AWAIT_ACK Or MsgStatus
      End If
      
'      Set objMsg = CreateObject(RS!ClassId)
      
      If blnOverride Then
'              Override message settings using ini file values
         If GlobalMsgFormat <> "" Then
            MsgFormat = GlobalMsgFormat
         End If
         SendTestData = blnGlobalTest
         Anonymize = blnGlobalAnonymize
      End If
      
      Select Case MsgFormat
         Case "ASTM1238,ASTM4.1"
            MsgStatus = MsgStatus Or MS_ASTM
            Set objMsg = New ASTM_1238
         
         Case "MEDRPT,NHS002"
            MsgStatus = MsgStatus Or MS_EDI2
            Set objMsg = New Edifact_NHS002
            
         Case "MEDRPT,NHS003"
            MsgStatus = MsgStatus Or MS_EDI3
            Set objMsg = New Edifact_NHS003
                     
         Case "MEDRPT,NHSRSR"
            MsgStatus = MsgStatus Or MS_RSR
            Set objMsg = New Edifact_NHSRSR
         
         Case "HTML,ICE001"
            MsgStatus = MsgStatus Or MS_HTML
            fExt = ".HTM"
            Set objMsg = New clsHTML ' New HTMLClass
'           Force sending via smtp
            baseName = "S"
            
         Case "XML,ICE002"
            MsgStatus = MS_XML
            Set objMsg = New XML_ReportMsg
            objMsg.SenderID = SenderID
            objMsg.ReceiverID = ReceiverID
         
         Case "HL7,HL7 v2.4"
            MsgStatus = MsgStatus Or MS_HL7
            'fExt = ".HL7"
            Set objMsg = New LabcommHL7_XML
            'Set objMsg = New HL7_XML
            'objMsg.ProduceXML = False
            
         Case "XML,XML v2.4"
            MsgStatus = MS_XML
            fExt = ".xml"
            Set objMsg = New clsHL7_XML
         
         'Case "XML,RPTMSG v1.2", "DOCMAN,KGH XML"
         Case "XML,RPTMSG v1.2", "DOCMAN,KGH XML", "DOCMAN,STH XML"
            MsgStatus = MsgStatus Or MS_DSCH
            Set objMsg = New XML_ReportMsg
            objMsg.SenderID = SenderID
            objMsg.ReceiverID = ReceiverID
            objMsg.MessageFormat = MsgFormat
            
         Case "HTML,CLINICAL LETTER"
            MsgStatus = 0
            Set objMsg = New EMail_ClinicalLetter
            objMsg.EMail_Type = FullLetter

'         Case "HTML,NOTIFY"
'            MsgStatus = 0
'            Set objMsg = New EMail_ClinicalLetter
'            objMsg.EMail_Type = Notify
    
         Case Else
            MsgBox "Invalid message format: " & MsgFormat, vbCritical, "System configuration Error"
            
      End Select
      
      histFile = ValidateFilepath(fs.BuildPath(HistoryPath, Format(Now(), "yymmdd")))
      baseName = baseName & Left(natCode & "_____", 6) & "_" & CStr(Interchange) & "_" & fileMsg & "_"
     
       If KeyId = "" Then
          fileBase = baseName & "@@@@"
          baseName = baseName & Format(Now(), "hhnnss")
       Else
          fileBase = baseName & KeyId & "_" & "@@@@"
          baseName = baseName & KeyId & "_" & Format(Now(), "hhnnss")
       End If

      If objMsg.Notification Then
         fExt = ".nfy"
         curFile = fs.BuildPath(PendingPath, natCode & "_Notifications_" & Format(Now(), "hh_nn_ss") & fExt)
         'histFile = fs.BuildPath(histFile, baseName & fExt)
         
         If fs.FileExists(histFile) Then
            Open histFile For Input As #1
            Open curFile For Output As #2
            Do Until EOF(1)
               Line Input #1, buf
               Print #2, buf
            Loop
            Close #1
            Close #2
         End If
         
      Else
                
'         outFile = fs.BuildPath(OutputPath, baseName & fExt)
         
         curFile = fs.BuildPath(PendingPath, baseName)
         
'         histFile = fs.BuildPath(histFile, baseName & IIf(Len(fExt) = 5, "." & Mid(fExt, 3), fExt))
         confFile = ValidateFilepath(fs.BuildPath(HistoryPath, "Conformance"))

      End If
      
'      eClass.LogMessage "Creating output file: " & curFile
      Open curFile For Binary As #1
      fileHdr = objMsg.FileHeader
      Put #1, , fileHdr
      
      Set RS = Nothing
      Set iceCmd = Nothing
   End If
   
   frmMain.lblFile.Caption = baseName
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsgControl.NewFile"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get NationalCode() As String
   NationalCode = natCode
End Property

Public Property Get Organisation() As String
   Organisation = orgId
End Property

Public Property Get OrganisationCode() As String
   OrganisationCode = orgCode
End Property

Public Property Get Trust() As String
   Trust = trustCode
End Property

Public Function ReadAck(AckFile As String) As Integer
   On Error GoTo procEH
   Dim rec As String
   Dim buf As String
   Dim objAck As New Edifact_ACKS
   Dim mPos As Integer
   
   Open AckFile For Input As #1
   buf = ""
   Do Until EOF(1)
      Line Input #1, rec
      buf = buf & rec
   Loop
   Close #1
   
   
   If Left(buf, 13) = "UNA:+.? 'UNB+" Then
      mPos = InStr(InStr(1, buf, "'") + 1, buf, "'")  '  Find the end of the UNB segment
      rec = Mid(buf, 1, mPos)
      
      If InStr(1, rec, "NHSACK", vbTextCompare) > 0 Then
         '  NHS Acknowledgement message
         objAck.Decode buf
         ReadAck = 1
         
      ElseIf InStr(1, rec, "MEDRPT", vbTextCompare) > 0 Then
         '  NHS002 or NHS003 Message
         ReadAck = -1
      End If
   
   Else
      '  Not an Edifact file
      ReadAck = 0
   End If
   
   If (MsgStatus And (MS_ACK_REJECT_PART Or MS_ACK_REJECT_ALL)) > 0 Then
      ackNatCode = objAck.ErroredNatCode
      ackIC = ackIC & "#" & objAck.ErroredInterchange
   End If
   
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceMsgControl.ReadAck"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Property Get ReportName()
   ReportName = RepId.ReportName
End Property

Public Property Let ReportImpExp(lngNewValue As Long)
   RepId.MessageImpExp = lngNewValue
End Property

Public Property Get ReportImpExp() As Long
   ReportImpExp = RepId.MessageImpExp
End Property

Public Property Let ReportStatus(sValue As enumReportStatus)
   RepId.OrStatus sValue
End Property

Public Property Get ReportStatus() As enumReportStatus
   ReportStatus = RepId.Status
End Property

'Public Sub RunTimeUpdate(natCode As String, _
'                         Specialty As String, _
'                         msgType As String, _
'                         lastRun As String)
'   On Error GoTo procEH
'   Dim rtIndex As String
'   Dim rtc As clsRunTimes
'
'   rtIndex = natCode & ":" & Specialty & ":" & msgType
'   Set rtc = colRT(rtIndex)
'   Set rtc = Nothing
'   Exit Sub
'
'procEH:
'   If Err.Number = 5 Then
'      Set rtc = New clsRunTimes
'      With rtc
'         .NationalCode = natCode
'         .Specialty = Specialty
'         .MessageType = msgType
'         .LastRunAt = lastRun
'      End With
'      colRT.Add rtc, rtIndex
'      Resume Next
'   Else
'      eClass.FurtherInfo = "IceMsgControl.RunTimeUpdate"
'      eClass.Add Err.Number, Err.Description, Err.Source
'   End If
'End Sub

Public Sub SetReport(EDIIndex As Long, _
                     repDisc As String, _
                     ReportId As String, _
                     Optional AckPosition As Long = 0, _
                     Optional AddToCollection As Boolean = False, _
                     Optional CopyTo As String = "", _
                     Optional LTSIndex As Long)
   Set RepId = New RepData
   RepId.Discipline = repDisc
   RepId.EDIIndex = EDIIndex
   RepId.LTSIndex = LTSIndex
   RepId.ReportName = ReportId
   RepId.ErrorReport = repStatus
   RepId.NationalCode = natCode
   RepId.MessageFormat = MsgFormat
   RepId.MessageImpExp = AckPosition
   RepId.CopyTo = CopyTo
   
   If repDisc = "HDR" Then
      Set ackMsgCtrl = RepId
   End If
   
   If AddToCollection Then
      colReps.Add RepId
   End If
End Sub

Public Property Get SpecialtyText() As String
   SpecialtyText = Specialty
End Property

Public Property Get TestData() As Boolean
   TestData = SendTestData
End Property

Public Property Let UnsetStatus(uValue As Long)
   RepId.AndStatus uValue
End Property

'Public Sub UpdateRunTimes()
'   Dim rtc As clsRunTimes
'
'   iceCon.BeginTrans
'
'   For Each rtc In colRT
'      strSQL = "UPDATE EDI_Loc_Specialties SET " & _
'                  "EDI_Last_Run = '" & Format(rtc.LastRunAt, "YYYYMMDD hh:nn:ss") & "' " & _
'               "WHERE EDI_Nat_Code = '" & rtc.NationalCode & "' " & _
'                  "AND EDI_Korner_Code = '" & rtc.Specialty & "' " & _
'                  "AND EDI_Msg_Format = '" & rtc.MessageType & "'"
'      iceCon.Execute strSQL
'
'      With rtc
'         SendMessage "IceMsg", .Specialty & ": Runtime updated to " & .LastRunAt & " for " & .NationalCode & " (" & .MessageType & ")"
'         eClass.LogMessage .NationalCode & "  " & .MessageType & " released and run-time updated for specialty " & .Specialty, , "Release"
'      End With
'   Next
'
'   If blnRollBack Then
'      iceCon.RollbackTrans
'   Else
'      iceCon.CommitTrans
'   End If
'
'   Set rtc = Nothing
'   Set colRT = Nothing
'End Sub

Public Function UpdateDocmanStatus(ReportIndex As String, DocmanRef As String, rejected As Boolean) As Integer
   '  Returns and integer
   '  0 = Success
   '  -1 = db Error
   '  1 = Failed to locate record
   
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim dmStatus As Integer
   
   On Error GoTo procEH
   strSQL = "SELECT sh.Service_ImpExp_Id FROM Service_ImpExp_Messages sm " & _
            "  JOIN Service_ImpExp_Headers sh " & _
            "  ON sm.Service_ImpExp_Id = sh.Service_ImpExp_Id " & _
            "WHERE Service_Type = 2 and Service_Report_Index = " & ReportIndex
   RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
   
   If RS.EOF = False Then
      Dim impExp As Long
      Dim iceCmd As ADODB.Command
      
      impExp = RS!Service_ImpExp_Id
      
      iceCon.BeginTrans
      
      Set iceCmd = New ADODB.Command
      
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("@MsgId", adInteger, adParamInput, , impExp)
         
         If Not rejected Then
            .CommandText = "ICEMSG_Ack_UpdateHeaders"
            .Parameters.Append .CreateParameter("@Status", adInteger, adParamInput, , (MS_ACK_RECEIVED Or MS_MSGOK))
         Else
            .CommandText = "ICEMSG_Ack_ErrorHeaders"
            .Parameters.Append .CreateParameter("@Status", adInteger, adParamInput, , MS_ACK_REJECT_ALL)
         End If
         
         .Execute
      End With
      
      msgData.AddImpExpComment impExp, "Docman Reference: " & DocmanRef, Format(Now(), "yyyy-mm-dd hh:MM;ss")
                 
      If blnRollBack Then
         iceCon.RollbackTrans
         dmStatus = -1
      Else
         iceCon.CommitTrans
         dmStatus = 0
      End If
   Else
      dmStatus = 1
   End If
   
   RS.Close
   
   UpdateDocmanStatus = dmStatus
   
   Exit Function
   
procEH:
   eClass.LogMessage "UpdateDocmanStatus: " & Err.Description
   iceCon.RollbackTrans
   UpdateDocmanStatus = -1
   
End Function

Public Sub WriteAcks()
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim repLogId As RepLogData
   Dim impExpId As Long
   Dim I As Integer
   Dim cText As String
   
   iceCon.BeginTrans
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("ImpExp", adInteger, adParamInput, , RepId.MessageImpExp)
      .Parameters.Append .CreateParameter("Status", adInteger, adParamInput, , RepId.Status)
   End With
   
   If (MsgStatus And MS_ACK_FAIL) = MS_ACK_FAIL Then
      ltIndex = -1
      RepId.MessageImpExp = AddImpExpHeader(RepId.ErrorReport, _
                                            0, _
                                            "", _
                                            0, _
                                            Format(Now(), "dd/mm/yyyy"), _
                                            1, _
                                            1, _
                                            False, _
                                            MsgStatus, _
                                            natCode)
      RepId.EDIIndex = RepId.MessageImpExp
   End If
   
   For I = 1 To colReps.Count
      Set RepId = colReps(I)
      If RepId.Discipline = "HDR" Then
         iceCmd.Parameters(0).Value = RepId.EDIIndex
         If (RepId.MessageImpExp And &H600) = 0 Then
            cText = "ICEMSG_ACK_UpdateHeaders"
         Else
            cText = "ICEMSG_ACK_ErrorHeaders"
            eClass.LogMessage "Acknowledgement Errors found with: " & RepId.ReportName
         End If
         RepId.MessageImpExp = 0
      Else
         cText = "ICEMSG_ACK_UpdateMessages"
         iceCmd.Parameters(0).Value = RepId.MessageImpExp
      End If
      
      With iceCmd
         .CommandText = cText
         .Parameters(1).Value = RepId.Status
         .Execute
      End With
                  
      For Each repLogId In RepId
         msgData.AddImpExpComment RepId.EDIIndex, repLogId.LogData, Format(Now(), "dd/mm/yyyy hh:nn"), repLogId.LogStatus, , RepId.MessageImpExp
      Next
                  
'      strArray = Split(RepId.LoggedMessages, "|")
'      For j = 0 To UBound(strArray)
'         msgData.AddImpExpComment RepId.EDIIndex, strArray(j), Now(), RepId.Status, "NHS_ACKS"
'      Next j
   Next I
   
   If blnRollBack = False Then
      iceCon.CommitTrans
   Else
      iceCon.RollbackTrans
      eClass.LogMessage "Rollback flag set - All updates rolled back"
   End If
   
   Set iceCmd = Nothing
   Exit Sub
   
procEH:
   iceCon.RollbackTrans
   
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
   eClass.CurrentProcedure = "IceMsgControl.WriteAcks"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub WriteFile()
   On Error GoTo procEH
   Dim rtc As clsRunTimes
   Dim pendFile As String
   Dim FileHeader As String
   Dim FileTrailer As String
   Dim fileBuf As String
   Dim buf As String
   Dim RepId As New RepData
   Dim eFile As File
   Dim flags As String
   Dim fileImpExp As Long
   Dim msgImpExp As Long
   Dim impExpRef As Long
   Dim dateAdded As String
   Dim I As Integer
   Dim j As Integer
   Dim confFlag As String
   Dim confCount As Long
   Dim sqlBuf As New StringBuffer
   Dim logBuf As New StringBuffer
   Dim repLogId As RepLogData
   Dim dbRetry As Integer
   Dim tFile As String
   Dim iceCmd As New ADODB.Command
   Dim strSQL As String
   
   DoEvents
   pendFile = curFile & fExt
   frmMain.Caption = fs.GetFileName(pendFile)
   
   If msgCount > 0 Then
      FileTrailer = objMsg.FileTrailer
      Put #1, , FileTrailer
      Close #1
   
      dateAdded = Format(Now(), "dd/mm/yyyy hh:nn:ss")
      
      If AckRequired Then
         MsgStatus = MsgStatus Or MS_AWAIT_ACK
      End If
   
      iceCon.BeginTrans
      
      impExpRef = 0
      If msgCount > 0 Then
         If impExpRef = 0 Then
            histFile = fs.BuildPath(histFile, fileBase & fExt)
            
            impExpRef = AddImpExpHeader(histFile, _
                                        msgCount, _
                                        ReceiverID, _
                                        CStr(Interchange), _
                                        dateAdded, _
                                        0, _
                                        IIf(AckRequired, -1, 0), _
                                        False, _
                                        MsgStatus, _
                                        natCode)
            fileBase = Replace(fileBase, "@@@@", Format(impExpRef, "000000000#"))
            outFile = fs.BuildPath(OutputPath, fileBase & fExt)
         End If
            
'         For Each rtc In colRT
'            strSQL = "UPDATE EDI_Loc_Specialties SET " & _
'                        "EDI_Last_Run = '" & Format(rtc.LastRunAt, "YYYYMMDD hh:nn:ss") & "' " & _
'                     "WHERE EDI_Nat_Code = '" & rtc.NationalCode & "' " & _
'                        "AND EDI_Korner_Code = '" & rtc.Specialty & "' " & _
'                        "AND EDI_Msg_Format = '" & rtc.MessageType & "'"
'            iceCon.Execute strSQL
'            eClass.LogMessage rtc.NationalCode & "  " & rtc.MessageType & " released for specialty " & rtc.Specialty, , "Release"
'
'         Next
'
''         For I = 1 To colRT.Count
''            Set rtc = colRT(I)
''            strSQL = "UPDATE EDI_Loc_Specialties SET " & _
''                        "EDI_Last_Run = '" & Format(rtc.LastRunAt, "YYYYMMDD hh:nn:ss") & "' " & _
''                     "WHERE EDI_Nat_Code = '" & rtc.NationalCode & "' " & _
''                        "AND EDI_Korner_Code = '" & rtc.Specialty & "' " & _
''                        "AND EDI_Msg_Format = '" & rtc.MessageType & "'"
''            iceCon.Execute strSQL
''            eClass.LogMessage rtc.NationalCode & "  " & rtc.MessageType & " released for specialty " & rtc.Specialty, , "Release"
''
''         Next I
'         Set rtc = Nothing
'         Set colRT = Nothing
         
         SendMessage "IceMsg", "Creating file: " & outFile & " (" & msgControl.MsgInBatch & " reports)"
         eClass.LogMessage "Creating output file: " & outFile & " (" & msgControl.MsgInBatch & " reports)"
         eClass.FurtherInfo = "Writing Service_ImpExp_Messages"
         For I = 1 To colReps.Count
            Set RepId = colReps(I)
   '         blnXMLRun = (RepId.NationalCode = "X_XML")
            If (RepId.Status And TF_IS_SUPPRESS) = TF_IS_SUPPRESS Then
               RepId.OrStatus RS_SUPPRESSION
            End If
            
            msgImpExp = msgData.AddImpExpMessage(RepId.EDIIndex, _
                                                 impExpRef, _
                                                 RepId.Discipline, _
                                                 dateAdded, _
                                                 RepId.Status, _
                                                 "", _
                                                 0, _
                                                 "", _
                                                 RepId.MessageImpExp, _
                                                 RepId.LTSIndex)
            
            LogStatus = &HFFFFFFDE
            
            For Each repLogId In RepId
               If (repLogId.LogStatus And LogStatus) <> 0 Then '= repLogId.LogStatus Then
                  msgData.AddImpExpComment impExpRef, repLogId.LogData, Format(Now(), "dd/mm/yyyy hh:nn"), repLogId.LogStatus, "IceMsg", msgImpExp
               End If
            Next
            
            If copyRecipient = "" Then
               If (RepId.Status And RS_CONFORMANCE) = 0 Then
                  If RepId.CopyTo = "" Then
   '                 Delete the record from EDI_Rep_List if the conformance target has not been reached
                     eClass.FurtherInfo = "Delete after success - report: " & RepId.ReportName & " - index " & RepId.EDIIndex
                  
                     sqlBuf.Append "DELETE FROM EDI_Rep_List WHERE EDI_Report_Index = " & RepId.EDIIndex
                  Else
   '                 The 'Extra_Copy_To' field was set on the replist during the previous run. Therefore we update
   '                 the national code to this while resetting the aforementioned field.
   '                 This is the 2nd time the report has been processed. We update the replist and process it
   '                 for the 3rd (and final) time
                     sqlBuf.Append "UPDATE EDI_Rep_List SET " & _
                                       "EDI_Loc_Nat_Code_To = '" & RepId.CopyTo & "', " & _
                                       "Extra_Copy_To = '' " & _
                                   "WHERE EDI_Report_Index = " & RepId.EDIIndex & "; "
                  End If
               Else  '  A Conformance report is required
   
   '              Preserve Rep_List entries but change the national code to to the conformance code.
   '              These will be picked up and sent on the next run.
                  confCount = confCount + 1
                  
                  sqlBuf.Append "UPDATE EDI_Rep_List SET " & _
                                    "EDI_Loc_Nat_Code_To = '" & Trim(ConfNatCode) & "' " & _
                                "WHERE EDI_Report_Index = " & RepId.EDIIndex & "; "
               
                  msgData.AddImpExpComment impExpRef, "Selected for conformance testing", Now(), RepId.Status, "IceMsg", msgImpExp
               End If
   
            Else
               
               If (RepId.Status And RS_CONFORMANCE) = 0 Then
   '              Conformance limit not breached - Simply update the replist with the copy-to recipient.
                  sqlBuf.Append "UPDATE EDI_Rep_List SET " & _
                                   "EDI_Loc_Nat_Code_To = '" & copyRecipient & "' " & _
                                "WHERE EDI_Report_Index = " & RepId.EDIIndex & "; "
                                
               Else
   '              To process the copyTo paramater we record it on the replist as the 'Extra_Copy_To field.
   '              This is picked up on the next run as the report is processed a second time
                  sqlBuf.Append "UPDATE EDI_Rep_List SET " & _
                                   "EDI_Loc_Nat_Code_To = '" & ConfNatCode & "', " & _
                                   "Extra_Copy_To = '" & copyRecipient & "' " & _
                                "WHERE EDI_Report_Index = " & RepId.EDIIndex & "; "
                                
               End If
            End If
            
         Next
         strSQL = sqlBuf.Value
         
         eClass.FurtherInfo = "Determine Conformance status"
         If confCount > 0 Then
   '        Force the conformance entry to be anonymized
            strSQL = strSQL & _
                     "UPDATE EDI_Msg_Types " & _
                     "SET EDI_Anonymize_Data = 1 " & _
                     "WHERE Organisation = '" & orgId & "' " & _
                        "AND EDI_Msg_Format = 'MEDRPT,NHS003' " & _
                        "AND EDI_Org_NatCode = '" & Trim(ConfNatCode) & "'"
         End If
         
         If strSQL <> "" Then
            iceCon.Execute strSQL
         End If
         
            strSQL = sqlBuf.Value
            
            If confCount > 0 Then
   '           Force the conformance entry to be anonymized
               strSQL = strSQL & _
                        "UPDATE EDI_Msg_Types " & _
                        "SET EDI_Anonymize_Data = 1 " & _
                        "WHERE Organisation = '" & orgId & "' " & _
                           "AND EDI_Msg_Format = 'MEDRPT,NHS003' " & _
                           "AND EDI_Org_NatCode = '" & Trim(ConfNatCode) & "'"
            End If
            
            If Left(natCode, 2) = "__" Then
               If CopyToDir = "" Then
                  CopyToDir = fs.BuildPath(fs.GetParentFolderName(outFile), "TEMPXML")
               End If
               
               tFile = fs.BuildPath(CopyToDir, "IC_" & Interchange & "_" & fs.GetBaseName(fs.GetTempName) & "." & Mid(natCode, 3))
               fs.MoveFile curFile, tFile
               
   '           The Header still points to the output file Run an amemndment script
               strSQL = strSQL & _
                        "UPDATE Service_ImpExp_Headers SET " & _
                           "ImpExp_File = '" & tFile & "' " & _
                        "WHERE Service_ImpExp_Id = " & impExpRef
            Else
               If Not objMsg.Notification Then
                   fs.CopyFile curFile, outFile
               End If
            End If
         
            If strSQL <> "" Then
               iceCon.Execute strSQL
            End If
            
   '        Set logstatus to ignore message type in message status when logging comment
            LogStatus = &HFFFFFF00
            
            If UCase(fExt) = ".HTML" Then
              msgData.AddImpExpComment impExpRef, MsgFormat & " file - prepared as HTML (" & colReps.Count & " reports.)", dateAdded, (MsgStatus And LogStatus)
            Else
               
               Select Case UCase(Right(fExt, 3))
                  Case "XEN"
                    msgData.AddImpExpComment impExpRef, MsgFormat & " file - ready for encryption (" & colReps.Count & " reports.)", dateAdded, (MsgStatus And LogStatus)
                     
                  Case "XML"
                    msgData.AddImpExpComment impExpRef, MsgFormat & " file - prepared as XML (" & colReps.Count & " reports.)", dateAdded, (MsgStatus And LogStatus)
                  
                  Case Else
                     msgData.AddImpExpComment impExpRef, MsgFormat & " file - not encrypted (" & colReps.Count & " reports.)", dateAdded, (MsgStatus And LogStatus)
                  
               End Select
               
               If Mid(UCase(fExt), 2, 1) = "H" Then
                  msgData.AddImpExpComment impExpRef, "Output file will be held in outray", Now(), (MsgStatus And LogStatus)
               End If
            End If
         
            If natCode = ConfNatCode Then
               fs.CopyFile curFile, fs.BuildPath(confFile, fs.GetFileName(curFile))
               msgData.AddImpExpComment impExpRef, "File copied to conformance directory", dateAdded, (MsgStatus And LogStatus)
            End If
            
            msgData.AddImpExpComment impExpRef, "File placed in outTray", dateAdded, (MsgStatus And LogStatus)
            
            If fs.FileExists(histFile) Then
               fs.DeleteFile (histFile)
            End If
            
            If Left(natCode, 2) <> "__" Then
               fs.MoveFile curFile, histFile
            End If
   '      End If
      End If
      
      If blnRollBack = False Then
         iceCon.CommitTrans
      Else
         iceCon.RollbackTrans
         eClass.LogMessage "Rollback flag set - All database updates rolled back", , "Diagnostic"
      End If
      
      eClass.LogMessage "File created"
      SendMessage "IceMsg", "Output file created"
   
   Else
      Close #1
      If fs.FileExists(curFile) Then
         Kill curFile
      End If
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   If Err.Number = 3234 Then
      If dbRetry < maxRetries Then
         eClass.LogMessage "Timeout attempting to update after processing. - Attempt: " & dbRetry
         dbRetry = dbRetry + 1
         Resume
      End If
   End If
   
   eClass.CurrentProcedure = "IceMsgControl.WriteFile"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
