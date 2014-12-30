VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RequeueData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rqReport As Boolean
Private rqIdent As String
Private outputDir As String

Public Property Let Identifier(strNewValue As String)
   rqIdent = strNewValue
End Property

Friend Property Get Identifier() As String
   Identifier = rqIdent
End Property

Public Property Let OutputDirectory(strNewValue As String)
   outputDir = strNewValue
End Property

Public Property Let RequeueReport(blnNewValue As Boolean)
   rqReport = blnNewValue
End Property

