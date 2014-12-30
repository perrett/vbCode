VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rtfSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private txtPos As Long
Private txtLen As Long
Private txtCol As Long
Private txtStyle As String
Private txtType As String

Public Property Let TextPosition(lngNewValue As Long)
   txtPos = lngNewValue
End Property

Public Property Get TextPosition() As Long
   TextPosition = txtPos
End Property

Public Property Let TextLength(lngNewValue As Long)
   txtLen = lngNewValue
End Property

Public Property Get TextLength() As Long
   TextLength = txtLen
End Property

Public Property Let TextColour(lngNewValue As Long)
   txtCol = lngNewValue
End Property

Public Property Get TextColour() As Long
   TextColour = txtCol
End Property

Public Property Let TextStyle(strNewValue As String)
   txtStyle = strNewValue
End Property

Public Property Get TextStyle() As String
   TextStyle = txtStyle
End Property

Public Property Let TextType(strNewValue As String)
   txtType = strNewValue
End Property

Public Property Get TextType() As String
   TextType = txtType
End Property
