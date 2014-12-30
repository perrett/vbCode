VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Edifact_ACKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msgAck As edifact_ack.MsgNHSACK
Private inpAck As edifact_ack.InputParser
Private InterchangeRef As String
Private ackNatCode As String
Private senderRef As String
Private receiverRef As String
Private Status As enumReportStatus
Private errCount As Long

'*****************************************************************************
'  Report Status:  An individual report within the acknowledgement file
'  Message Status: The overall state of the file.
'  AckStatus:      When a file acknowledges multiple files, this records the
'                  state of each source file
'*****************************************************************************

Public Property Get ackErrors() As Long
   ackErrors = errCount
End Property

Private Sub Analyse(MsgNo As Long)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim errStr As String
   Dim I As Integer
   Dim impExp As Long
   Dim segId As String
   Dim segPos As String
   Dim segErrCode As String
   Dim AckStatus As Long
   
   With msgAck.MsgGrp(MsgNo).SG1
'     Find the header record that matches this interchange and trader code
      With .NIR
         InterchangeRef = .Reference.ReferenceNumber.Value
         senderRef = .SubjectInterchangeSender.SubjectInterchangeSenderInformation.Value
         receiverRef = .SubjectInterchangeRecipient.SubjectInterchangeRecipoentIdentification.Value
         
         With iceCmd
            .ActiveConnection = iceCon
            .CommandType = adCmdStoredProc
            .CommandText = "ICEMSG_Ack_ReadHeaders"
            .Parameters.Append .CreateParameter("TradeRef", adVarChar, adParamInput, 15, receiverRef)
            .Parameters.Append .CreateParameter("iChange", adVarChar, adParamInput, 10, InterchangeRef)
            Set RS = .Execute
         End With
      End With
         
      strSQL = "SELECT EDI_NatCode " & _
               "FROM EDI_Recipients er " & _
                  "INNER JOIN EDI_Recipient_Ref err " & _
                  "ON er.Ref_Index = err.Ref_Index " & _
                  "INNER JOIN EDI_Interchange_No ei " & _
                  "ON er.Ref_Index = ei.Ref_Index " & _
               "WHERE EDI_Trader_Account + EDI_Free_Part = '" & receiverRef & "'"
      
      If RS.RecordCount = 0 Then
'        No matching records - attempt to find the national code
         RS.Close
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         
         If RS.EOF Then
            errStr = "(Trader code + free part " & receiverRef & " not found)"
         Else
            errStr = RS!EDI_NatCode & " (" & receiverRef & ")"
         End If
         
         msgControl.SetReport 0, "HDR", "Interchange", 0, True
         msgControl.AckStatus = MS_ACK_FAIL
'         msgControl.ReportStatus = MS_ACK_FAIL
'         msgControl.MessageStatus = MS_ACK_FAIL
         msgControl.LogReportMessage 0, "Interchange " & InterchangeRef & " for EDI_Recipient " & errStr & _
                                     " does not exist in the audit records"
      Else
      
         With .NIR
            impExp = RS!Service_ImpExp_Id
'           Set up the Header for this group of acks
            msgControl.SetReport impExp, "HDR", fs.GetFileName(RS!ImpExp_File), -1, True
            
'           Get the existing header status
            If IsNull(RS!Header_Status) Then
               msgControl.AckStatus = (MS_ACK_RECEIVED Or MS_MSGOK)
               msgControl.ReportStatus = (MS_MSGOK Or MS_ACK_RECEIVED)
               msgControl.MessageStatus = (MS_MSGOK Or MS_ACK_RECEIVED)
            ElseIf .ActionCoded.Value = "IAF" Then
               msgControl.AckStatus = RS!Header_Status
               msgControl.ReportStatus = RS!Header_Status
               msgControl.MessageStatus = RS!Header_Status
            Else
'              Remove the 'MSG_OK' flag
               msgControl.AckStatus = (RS!Header_Status And &H7FFFFFFF)
               msgControl.ReportStatus = (RS!Header_Status And &H7FFFFFFF)
               msgControl.MessageStatus = (RS!Header_Status And &H7FFFFFFF)
            End If
                        
            Select Case .ActionCoded.Value
               Case "IAF"  '  Interchange fully accepted
                  msgControl.LogReportMessage MS_ACK_RECEIVED, "Interchange fully accepted"
                  msgControl.AckAccepted
                  msgControl.AckStatus = MS_ACK_RECEIVED
                  msgControl.ReportStatus = MS_ACK_RECEIVED
                  
               Case "IAP"  '  Interchange partially accepted
                  msgControl.LogReportMessage MS_ACK_REJECT_PART, "Interchange partially accepted - One or more messages rejected"
                  msgControl.AckRejected True
                  msgControl.AckStatus = MS_ACK_REJECT_PART
                  msgControl.ReportStatus = MS_ACK_REJECT_PART
                  
               Case "IRI"  '  Interchange fully rejected - unreadable
                  If .ErrorCategoryCoded.Value = "CRE" Then
                     errStr = "Cryptography problem"
                     msgControl.MessageStatus = MS_ACK_CRYPTO
                     msgControl.AckStatus = MS_ACK_CRYPTO
                  ElseIf .ErrorCategoryCoded.Value = "ISE" Then
                     errStr = "Syntax Error"
                  Else
                     errStr = "Error not defined"
                  End If
                  
                  msgControl.LogReportMessage MS_ACK_REJECT_ALL, "Complete Interchange rejected - " & errStr
                  msgControl.AckRejected False
                  msgControl.AckStatus = MS_ACK_REJECT_ALL
                  msgControl.ReportStatus = MS_ACK_REJECT_ALL
                  
               Case "IRM"  '  Interchange fully rejected - Error in one message
                  msgControl.LogReportMessage MS_ACK_REJECT_ALL, "Complete Interchange rejected due to errors in one message"
                  msgControl.AckRejected False
                  msgControl.ReportStatus = MS_ACK_REJECT_ALL
                  msgControl.AckStatus = MS_ACK_REJECT_ALL
                  
               Case "IRA"  '  Interchange completely rejeceted - all messages in error
                  msgControl.LogReportMessage MS_ACK_REJECT_ALL, "Complete Interchange Rejected - All messages in error"
                  msgControl.AckRejected False
                  msgControl.ReportStatus = MS_ACK_REJECT_ALL
                  msgControl.AckStatus = MS_ACK_REJECT_ALL
                  
            End Select
         End With
         
         AckStatus = msgControl.ReportStatus
         
         FTXData .FTX
         
'        Read all the reports for this message
         With iceCmd
            .CommandText = "ICEMSG_Ack_ReadMessages"
            .Parameters.Delete (1)
            .Parameters(0).Value = impExp
            Set RS = .Execute
         End With
                  
         For I = 0 To .SG2.Count - 1
            With .SG2(I)
               With .NMR
                  segId = IIf(.SegmentTagCoded.Value = "", "NA", .SegmentTagCoded.Value)
                  segPos = IIf(.SegmentPositionInMessage.Value = "", "NA", .SegmentPositionInMessage.Value)
                  segErrCode = .ErrorCategoryCoded.Value
               End With
               
               RS.MoveFirst
               eClass.FurtherInfo = "Looking for Message No.: " & .NMR.Reference.ReferenceNumber.Value
               RS.Find "Service_Message_Id = " & .NMR.Reference.ReferenceNumber.Value
               
               If RS.EOF = False Then
                  msgControl.SetReport impExp, RS!Discipline, RS!ReportId, RS!Service_impExp_Message_Id, True
                  msgControl.ReportStatus = IIf(IsNull(RS!Message_Status), 0, RS!Message_Status)
                  msgControl.ReportImpExp = RS!Service_impExp_Message_Id
                  
                  If .NMR.ActionCoded.Value = "MR" Then
'                    Message rejected
                     msgControl.ReportStatus = AckStatus
                     errCount = errCount + 1
                     
                     If segErrCode = "MSE" Then
                        errStr = "Syntax error in message"
                     ElseIf segErrCode = "MNE" Then
                        errStr = "Data error in message"
                     Else
                        errStr = ""
                     End If
                     
                     msgControl.LogReportMessage AckStatus, errStr & " Rejection data - Report: " & _
                                                 RS!ReportId & " (Segment: " & segId & _
                                                 " Position: " & segPos & ")"
                  End If
                  
                  FTXData .FTX
               Else
                  msgControl.LogReportMessage MS_ACK_FAIL, "Message id " & .NMR.Reference.ReferenceNumber.Value & _
                                              " not present in file. "
               End If
            End With
'            RS.MoveNext
         Next I
      End If
   End With
   
   RS.Close
   
'  If there are any errors, use the national code as part of the file name
   If (msgControl.MessageStatus And (MS_ACK_REJECT_PART Or MS_ACK_REJECT_ALL)) > 0 Then
      RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      ackNatCode = RS!EDI_NatCode
   End If
   
   Set RS = Nothing
   Set iceCmd = Nothing
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_ACKS.Analyse"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Decode(mBuf As String)
   On Error GoTo procEH
   Dim I As Long
   
   Set msgAck = CreateObject("edifact_ack.MsgNHSACK") ' New edifact_ack.MsgNHSACK
   Set msgAck = inpAck.Parse(mBuf)
   
   For I = 0 To msgAck.MsgGrp.Count - 1
      Analyse I
   Next I
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   msgControl.MessageStatus = MS_ACK_FAIL
'   eClass.CurrentProcedure = "Edifact_ACKS.Decode"
'   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get ErroredNatCode() As String
   ErroredNatCode = ackNatCode
End Property
   
Public Property Get ErroredInterchange() As String
   ErroredInterchange = InterchangeRef
End Property

Private Sub FTXData(ftxSeg As Object)
   Dim strBuf As String
   
   With ftxSeg
      If .IsPresent Then
         
         strBuf = .TextLiteral.FreeText1.Value
      
         If .TextLiteral.FreeText2.IsPresent Then
            strBuf = strBuf & " " & .TextLiteral.FreeText2.Value
         End If
         
         If .TextLiteral.FreeText3.IsPresent Then
            strBuf = strBuf & " " & .TextLiteral.FreeText3.Value
         End If
         
         If .TextLiteral.FreeText4.IsPresent Then
            strBuf = strBuf & " " & .TextLiteral.FreeText4.Value
         End If
         
         If .TextLiteral.FreeText5.IsPresent Then
            strBuf = strBuf & " " & .TextLiteral.FreeText5.Value
         End If
         
         msgControl.LogReportMessage msgControl.ReportStatus, strBuf
      End If
   End With
End Sub

Private Sub UpdateForTesting(ImpId As Long)
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim mPos As Integer
   
   strSQL = "SELECT Service_ImpExp_Message_Id " & _
            "FROM Service_ImpExp_Messages " & _
            "WHERE Service_ImpExp_Id = " & ImpId
   RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
   
   mPos = 1
   Do Until RS.EOF
      strSQL = "UPDATE Service_ImpExp_Messages SET " & _
               "Position_In_Message = " & mPos & _
               " WHERE Service_ImpExp_Message_Id = " & RS!Service_impExp_Message_Id
      iceCon.Execute strSQL
      mPos = mPos + 1
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
End Sub

Private Sub Class_Initialize()
   On Error GoTo procEH:
   Set inpAck = CreateObject("edifact_ack.InputParser")
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Edifact_ACKS.Class_Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
