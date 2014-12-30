VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSMQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oQTrans As MSMQTransaction
Private oQTD As MSMQCoordinatedTransactionDispenser
Private oQInfo As MSMQ.MSMQQueueInfo
Private oQueue As MSMQ.MSMQQueue
Private sQueueName As String
Private iFailCount As Integer

Public Property Get QueueId() As String
    QueueId = sQueueName
End Property

Public Property Let QueueId(QueueName As String)
    sQueueName = QueueName
End Property

Public Sub IncrementFail(Optional Reset As Boolean = False)
   If Reset Then
      iFailCount = 0
   Else
      iFailCount = iFailCount + 1
   End If
End Sub

Public Function SendQueueMessage(Label As String, Data As String) As String
   On Error GoTo procEH
   
   Dim qData As String
   'Dim oQInfo As MSMQ.MSMQQueueInfo
   'Dim oQueue As MSMQ.MSMQQueue
   Dim oMessage As MSMQ.MSMQMessage
   Dim oQTrans As MSMQTransaction
   'Dim oQTD As New MSMQCoordinatedTransactionDispenser
   Dim rData As String
   
   Set oQInfo = New MSMQQueueInfo
   
   oQInfo.FormatName = sQueueName
   
   'Open the queue for send access.
   Set oQueue = oQInfo.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
      
      
   If oQueue.IsOpen Then
      Set oQTrans = oQTD.BeginTransaction
      
      Set oMessage = New MSMQMessage
      
      With oMessage
        .Label = Label
        .Body = Data
        .Send oQueue, oQTrans
      End With
       
      oQTrans.Commit
      
      rData = ""
      oQueue.Close
   
   Else
       rData = "Unable to open " & sQueueName
   End If
   
   SendQueueMessage = rData
   
  Exit Function

procEH:
    SendQueueMessage = Err.Description
    
End Function

Public Function ReadQueue() As String
    
    On Error GoTo procEH
    
   If iFailCount <= 5 Then
      Dim qData As String
      Dim oMessage As MSMQ.MSMQMessage
      
      'Initialize the MSMQQueueInfo object.
      Set oQInfo = New MSMQQueueInfo
      
      oQInfo.FormatName = sQueueName
      
      Set oQTD = New MSMQCoordinatedTransactionDispenser
      
      'Open the queue for receive access.
      Set oQueue = oQInfo.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
      
      If oQueue.IsOpen Then
        Set oQTrans = oQTD.BeginTransaction
        Set oMessage = oQueue.ReceiveCurrent(Transaction:=oQTrans, WantBody:=True, ReceiveTimeout:=1000)
        
        If oMessage Is Nothing Then ' Nothing in the queue
          qData = ""
        Else
          qData = oMessage.Body
        End If
        
      Else
        qData = "Error|Unable to open Queue"
      End If
      
   Else
      qData = ""
   End If
   
   ReadQueue = qData
   
   Exit Function

procEH:
  ReadQueue = "Error|" & Err.Description

End Function

Public Sub CommitRead()

    oQTrans.Commit
    oQueue.Close

End Sub

Public Sub AbortRead()
    
    oQTrans.Abort
    oQueue.Close

End Sub

