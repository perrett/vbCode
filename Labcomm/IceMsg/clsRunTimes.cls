VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRunTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private natCode As String
Private specCode As String
Private msgType As String
Private lastRun As String

Public Property Let NationalCode(strNewValue As String)
   natCode = strNewValue
End Property

Public Property Get NationalCode() As String
   NationalCode = natCode
End Property

Public Property Let Specialty(strNewValue As String)
   specCode = strNewValue
End Property

Public Property Get Specialty() As String
   Specialty = specCode
End Property

Public Property Let MessageType(strNewValue As String)
   msgType = strNewValue
End Property

Public Property Get MessageType() As String
   MessageType = msgType
End Property

Public Property Let LastRunAt(strNewValue As String)
   lastRun = strNewValue
End Property

Public Property Get LastRunAt() As String
   LastRunAt = lastRun
End Property

