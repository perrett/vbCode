VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDtsResponseData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************
'  Responses from the DTS Server
'***********************************
Private dtsStatus As String
Private fromSMTP As String
Private toSMTP As String
Private sentTo As String
Private DTSMsgId As String
Private DTSProcess As String
Private DTSWorkflow As String
Private transferDate As String
Private transferStatus As String
Private sentDate As String
Private FileId As String
Private DTSLocalId As String
Private timeStamp As Date

Public Property Let CurrentStatus(strNewValue As String)
   transferStatus = strNewValue
End Property

Public Property Get CurrentStatus() As String
   CurrentStatus = transferStatus
End Property

Public Property Get DTSChecked() As Date
   DTSChecked = timeStamp
End Property

Public Property Let DTSResponse(strNewValue As String)
   dtsStatus = strNewValue
   timeStamp = Now()
End Property

Public Property Get DTSResponse() As String
   DTSResponse = dtsStatus
End Property

Public Property Let FileData(strNewValue As String)
   FileId = Mid(strNewValue, 10)
   FileId = Left(FileId, InStr(1, FileId, ".") - 1)
End Property

Public Property Get FileData() As String
   FileData = FileId
End Property

Public Property Let FromSMTPAddress(strNewValue As String)
   fromSMTP = strNewValue
End Property

Public Property Get FromSMTPAddress() As String
   FromSMTPAddress = fromSMTP
End Property

Public Property Let localId(strNewValue As String)
   DTSLocalId = strNewValue
End Property

Public Property Get localId() As String
   localId = DTSLocalId
End Property

Public Property Let MsgId(strNewValue As String)
   DTSMsgId = strNewValue
End Property

Public Property Get MsgId() As String
   MsgId = DTSMsgId
End Property

Public Property Let ProcessId(strNewValue As String)
   DTSProcess = strNewValue
End Property

Public Property Get ProcessId() As String
   ProcessId = DTSProcess
End Property

Public Property Let Recipient(strNewValue As String)
   sentTo = strNewValue
End Property

Public Property Get Recipient() As String
   Recipient = sentTo
End Property

Public Property Let SentOn(strNewValue As String)
   sentDate = Replace(strNewValue, "T", " @ ")
End Property

Public Property Get SentOn() As String
   SentOn = sentDate
End Property

Public Property Let ToSMTPAddress(strNewValue As String)
   toSMTP = strNewValue
End Property

Public Property Get ToSMTPAddress() As String
   ToSMTPAddress = toSMTP
End Property

Public Property Let TransferredOn(strNewValue As String)
   transferDate = Replace(strNewValue, "T", " @ ")
End Property

Public Property Get TransferredOn() As String
   TransferredOn = transferDate
End Property

Public Property Let WorkFlowId(strNewValue As String)
   DTSWorkflow = strNewValue
End Property

Public Property Get WorkFlowId() As String
   WorkFlowId = DTSWorkflow
End Property

