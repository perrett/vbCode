VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sampleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private strIdent As String
Private seqNo As Integer

Friend Property Let SampleId(strNewValue As String)
   strIdent = strNewValue
End Property

Friend Property Get SampleId() As String
   SampleId = strIdent
End Property

Friend Property Let SequenceId(intNewValue As Integer)
   seqNo = intNewValue
End Property

Friend Property Get SequenceId() As Integer
   SequenceId = seqNo
End Property
