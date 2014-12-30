VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConformanceData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private LTIndex As Long
Private TraderCode As String
Private RunningTotal As Long
Private TriggerValue As Long
Private NationalCode As String

Friend Property Get AddToTotal() As Boolean
   RunningTotal = RunningTotal + 1
   If RunningTotal >= TriggerValue And TriggerValue > 0 Then
      RunningTotal = 0
      AddToTotal = True
   Else
      AddToTotal = False
   End If
End Property

Friend Property Let CurrentTotal(lngNewValue As Long)
   RunningTotal = lngNewValue
End Property


Friend Property Get CurrentTotal() As Long
   CurrentTotal = RunningTotal
End Property

Friend Property Let Index(lngNewValue As Long)
   LTIndex = lngNewValue
End Property

Friend Property Get Index() As Long
   Index = LTIndex
End Property

Friend Property Let NatCode(strNewValue As String)
   NationalCode = strNewValue
End Property

Friend Property Get NatCode() As String
   NatCode = NationalCode
End Property

Friend Property Get SendConformance() As Boolean
   SendConformance = (RunningTotal >= TriggerValue)
End Property

Friend Property Let TraderId(strNewValue As String)
   TraderCode = strNewValue
End Property

Friend Property Get TraderId() As String
   TraderId = TraderCode
End Property

Friend Property Let Trigger(lngNewValue As Long)
   TriggerValue = lngNewValue
End Property
