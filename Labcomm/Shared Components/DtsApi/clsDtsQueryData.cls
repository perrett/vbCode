VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDtsQueryData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************
'  Details of outstanding requests to the DTS Server
'*******************************************************

Private domFragment As MSXML2.IXMLDOMDocumentFragment
Private reqFile As String
Private failCount As Integer
Private locID As String

Public Property Let DTSFailure(intNewValue As Integer)
   failCount = failCount + 1
End Property

Public Property Get DTSFailure() As Integer
   DTSFailure = failCount
End Property

Public Property Let DTSQuery(oNewValue As MSXML2.IXMLDOMDocumentFragment)
   Set domFragment = oNewValue
End Property

Public Property Get DTSQuery() As MSXML2.IXMLDOMDocumentFragment
   Set DTSQuery = domFragment
End Property

Public Property Let Filename(strNewValue As String)
   reqFile = strNewValue
End Property

Public Property Get Filename() As String
   Filename = reqFile
End Property

Public Property Let localId(strNewValue As String)
   locID = strNewValue
End Property

Public Property Get localId() As String
   localId = locID
End Property
