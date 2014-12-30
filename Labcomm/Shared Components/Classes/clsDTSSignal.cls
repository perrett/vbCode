VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDTSSignal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************************************
'
'  This is a global event signaller. The Request classes call this subroutine to notify
'  the parent (clsDtsApi) that data hase been returned from the xmlHTTP object.
'
'  Using an event in the request class will confuse vb, if a second event is fired before
'  the code has dealt with the first event.
'
'*****************************************************************************************

Public Event ResponseReceived(LocalId As String)

Public Sub Notify(LocalId As String)
   RaiseEvent ResponseReceived(LocalId)
End Sub
