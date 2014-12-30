VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "errorData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private currentProc As String
Private extraInfo As String
Private errNum As String
Private errDesc As String
Private errSrc As String

Friend Sub Add(ProcId As String, _
               Number As Long, _
               Description As String, _
               Source As String, _
               Optional Info As String = "")

   currentProc = ProcId
   errNum = Number
   errDesc = Description
   errSrc = Source
   extraInfo = Info

End Sub

Friend Function Retrieve() As Variant
   Retrieve = Array(currentProc, extraInfo, errNum, errDesc, errSrc)
End Function
