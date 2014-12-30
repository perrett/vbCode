VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrowserData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oHwnd As Long
Private oWindProc As Long
Private BrowserId As String

Friend Property Let BrowserIdentity(ByVal strNewValue As String)
   BrowserId = strNewValue
End Property

Friend Property Get BrowserIdentity() As String
   BrowserIdentity = BrowserId
End Property

Friend Property Let OldHWnd(ByVal lngNewValue As Long)
   oHwnd = lngNewValue
End Property

Friend Property Get OldHWnd() As Long
   OldHWnd = oHwnd
End Property

Friend Property Let OriginalWndProc(ByVal lngNewValue As Long)
   oWindProc = lngNewValue
End Property

Friend Property Get OriginalWndProc() As Long
   OriginalWndProc = oWindProc
End Property

