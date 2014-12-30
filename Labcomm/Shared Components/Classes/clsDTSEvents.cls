VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDTSEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*************************************************************************************
'  A wrapper class for the XMLHTTP object event handler
'
'  We need to provide a link back to the calling class in order to check the status
'  of the outstanding request and process the response.
'*************************************************************************************

Private rData As clsDTSRequest
Private LocalId As String

Public Property Let CallingClass(oClass As clsDTSRequest)
   Set rData = oClass
   LocalId = rData.LocalId
End Property

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
   If rData.HTTPStatus = 4 Then
      rData.ProcessResponse
   End If
End Sub
