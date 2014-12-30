VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRootNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Root As String

Public Property Get Load() As String

    Load = Root

End Property

Public Property Let Load(ByVal vNewValue As String)

    Root = vNewValue
    oXMLMessage.loadXML Root
    RootNodeLoaded = True

End Property
