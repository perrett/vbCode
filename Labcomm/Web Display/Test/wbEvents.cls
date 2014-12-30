VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "wbEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub onClick(hDoc As MSHTML.HTMLDocument)
   Dim src As MSHTML.IHTMLElement
   Dim fnId As String
   Dim fPos As Integer
   Dim fLen As Integer
   
   Set src = hDoc.activeElement
   
   If Not IsNull(src.onClick) Then
      fPos = InStr(1, src.onClick, "{") + 2
      fLen = (InStr(1, src.onClick, "}") - fPos) - 1
      fnId = Mid(src.onClick, fPos, fLen)
      MsgBox fnId & "(" & Mid(src.innerText, 3) & ")"
   End If
End Sub

