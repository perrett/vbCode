VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Root As clsRootNode
Public Element As clsElement
Private sXMLDeclaration As String
Private sDOCTYPEDeclaration As String

Private Sub Class_Initialize()

    Set oXMLMessage = New DOMDocument
    Set Root = New clsRootNode
    Set Element = New clsElement

    oXMLMessage.async = False
    oXMLMessage.preserveWhiteSpace = False

End Sub

Private Sub Class_Terminate()

    Set oXMLMessage = Nothing

End Sub

Public Property Get MessageContents() As String

    MessageContents = sXMLDeclaration & sDOCTYPEDeclaration & oXMLMessage.xml

End Property

Public Property Let XMLDeclaration(ByVal vNewValue As String)

    sXMLDeclaration = vNewValue

End Property

Public Property Let DOCTYPEDeclaration(ByVal vNewValue As String)

    sDOCTYPEDeclaration = vNewValue

End Property


