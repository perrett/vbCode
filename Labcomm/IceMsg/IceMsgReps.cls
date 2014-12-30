VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IceMsgReps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  Store individual report indices for a message. Objects held in colrep
'  collection object in ICEMsgData

Private repName As String
Private RepId As String
Private EDIRepId As Long
Private SuppressionWarning(4) As String
Private OutputProduced As Boolean

Private colMsg As New Collection

Friend Property Let EDIIndex(lngNewValue As Long)
   EDIRepId = lngNewValue
End Property

Friend Property Get EDIIndex() As Long
   EDIIndex = EDIRepId
End Property

Friend Property Let DataProduced(blnNewValue As Boolean)
   OutputProduced = blnNewValue
End Property

Friend Property Get DataProduced() As Boolean
   DataProduced = OutputProduced
End Property

Friend Property Let ReportId(vNewValue As String)
   RepId = vNewValue
End Property

Public Property Get ReportId() As String
   ReportId = RepId
End Property

Friend Property Let ReportName(vNewValue As String)
   repName = vNewValue
End Property

Public Property Get ReportName() As String
   ReportName = repName
End Property

Friend Property Let SuppressData(Level As Integer, _
                                 dValue As String)
   SuppressionWarning(Level) = dValue
End Property

Friend Property Get SuppressData(Level As Integer) As String
   SuppressData = SuppressionWarning(Level)
End Property

Friend Sub AddLogMessage(MsgStr As String)
   Dim rMsg As New RepMsg
   
   rMsg.LogData = MsgStr
   colMsg.Add rMsg
   Set rMsg = Nothing
End Sub

Friend Function ReadLogMessage() As Variant
   Dim i As Integer
   Dim rMsg As RepMsg
   Dim strTemp As String
   
   For i = 1 To colMsg.Count
      Set rMsg = colMsg(i)
      strTemp = rMsg.LogData & "|"
   Next i
   ReadLogMessage = Left(strTemp, Len(strTemp) - 1)
End Function
