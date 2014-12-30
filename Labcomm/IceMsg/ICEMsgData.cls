VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICEMsgData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  Class object held in the coldata collection object in IceMsgControl. Stores
'  Details for individual AHSLMessaging objects

'  The collection 'colReps' allows all the report identities which make up a
'  message to be stored.

Private fExt As String
Private dMethod As String
Private repCount As Integer
Private validCount As Integer
Private validReps As Integer
Private msgRef As Object
Private msgKey As String
Private dType As String
Private ControlRef As String
Private TraderRef As String
Private expectAcks As Integer
Private fileEncrypt As Boolean
Private repDets As IceMsgReps
Private confData As String
Private confLTI As String
Private ReadCodeData() As String
Private rCnt As Integer
Private trCnt As Integer

Private colReps As New Collection

Private Sub Class_Initialize()
   confData = ""
   repCount = 0
   rCnt = 0
   trCnt = 0
End Sub

Friend Property Let AcksRequired(intNewValue As Integer)
   expectAcks = intNewValue
End Property

Friend Property Get AcksRequired() As Integer
   AcksRequired = expectAcks
End Property

Friend Sub AddReport(EDIRepListIndex As Long, _
                     ReportName As String, _
                     ReportID As String, _
                     OutputProduced As Boolean)
   
   Dim repDets As New IceMsgReps
   Dim i As Integer
   
   repCount = repCount + 1
   repDets.DataProduced = OutputProduced
   If OutputProduced Then
      validCount = validCount + 1
   End If
   repDets.EDIIndex = EDIRepListIndex
   repDets.ReportID = ReportID
   repDets.ReportName = ReportName
   If TypeName(objMsg) <> "Nothing" Then
      For i = 0 To 3
         repDets.SuppressData(i) = objMsg.SuppressedCodes(i)
      Next i
   End If
   colReps.Add repDets, CStr(repCount)
   Set repDets = Nothing

End Sub

Friend Property Let ConformanceCode(strNewValue As String)
   confData = strNewValue
End Property

Friend Property Get ConformanceCode() As String
   ConformanceCode = confData
End Property

Friend Property Let ConformanceTraderIndex(strNewValue As String)
   confLTI = strNewValue
End Property

Friend Property Get ConformanceTraderIndex() As String
   ConformanceTraderIndex = confLTI
End Property

Friend Property Let DeliveryMethod(ByVal strNewValue As String)
   dMethod = strNewValue
End Property

Public Property Get DeliveryMethod() As String
   DeliveryMethod = dMethod
End Property

Friend Property Let Discipline(ByRef vNewValue As String)
   dType = vNewValue
End Property

Public Property Get Discipline() As String
   Discipline = dType
End Property

Friend Property Let Encrypted(ByRef vNewValue As Boolean)
   fileEncrypt = vNewValue
End Property

Public Property Get Encrypted() As Boolean
   Encrypted = fileEncrypt
End Property

Friend Property Let FailedReadCode(strNewValue As String)
   
End Property

Friend Property Let fileExt(ByRef strNewValue As String)
   fExt = strNewValue
End Property

Public Property Get fileExt() As String
   fileExt = fExt
End Property

Friend Property Let InterchangeId(strNewValue As String)
   ControlRef = strNewValue
End Property

Public Property Get InterchangeId() As String
   InterchangeId = ControlRef
End Property

Friend Property Let MessageKey(ByRef strNewValue As String)
   msgKey = strNewValue
End Property

Public Property Get MessageKey() As String
   MessageKey = msgKey
End Property

Friend Property Let MessageRef(ByRef vNewValue As Object)
   Set msgRef = vNewValue
End Property

Public Property Get MessageRef() As Object
   Set MessageRef = msgRef
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = colReps.[_NewEnum]
End Function

Public Property Let ReportCount(intNewValue As Integer)
   repCount = repCount + intNewValue
End Property

Public Property Get ReportCount() As Integer
   ReportCount = validCount   '  repCount
End Property

Friend Property Get ReportEDIIndex() As Long
   ReportEDIIndex = colReps.Item(0).EDIIndex
End Property

Public Property Get ReportIndex(Index) As String
   ReportIndex = colReps.Item(Index).ReportID
End Property

Public Property Get ReportName(Index) As String
   ReportName = colReps.Item(Index).ReportName
End Property

Friend Property Let TraderId(strNewValue As String)
   TraderRef = strNewValue
End Property

Public Property Get TraderId() As String
   TraderId = TraderRef
End Property

Friend Property Get TotalReadCodeErrors() As Integer
   TotalReadCodeErrors = rCnt
End Property

Friend Property Let ReadCodeError(strNewValue As String)
   ReDim Preserve ReadCodeData(rCnt)
   ReadCodeData(rCnt) = strNewValue
   rCnt = rCnt + 1
End Property

Friend Property Get ReadCodeError() As String
   If trCnt < rCnt Then
      ReadCodeError = ReadCodeData(trCnt)
   Else
      ReadCodeError = ""
   End If
   trCnt = trCnt + 1
End Property

Friend Property Let MarkInvalid(Index As String)
   Set repDets = colReps(Index)
   repDets.DataProduced = False
   Set repDets = Nothing
End Property
