VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepLogData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StatusFlag As Long
Private LoggedData As String

Friend Property Get LogData() As String
   LogData = LoggedData
End Property

Friend Property Get LogStatus() As Long
   LogStatus = StatusFlag
End Property

Friend Property Let MessageData(Status As Long, _
                                msgValue As String)
   StatusFlag = Status
   LoggedData = msgValue
End Property

Friend Property Get MessageData(Status As Long) As String
   If (StatusFlag And Status) = StatusFlag Then
      MessageData = LoggedData
   Else
      MessageData = ""
   End If
End Property
