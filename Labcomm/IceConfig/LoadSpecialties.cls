VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadSpecialties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private tv1 As TreeView
Private nd(3) As node

Private Sub Class_Initialize()
   Set tv1 = frmMain.TreeView1
   With frmMain.ediPr
      .Pages.Clear
      .Pages.Add "Main", "Service_Discipline_Map"
      .PropertyItems.Clear
      .PropertyItems.Add "CODE", "Korner Code", plpsCustom, , "The National Korner Code"
      .PropertyItems.Add ""
   End With
End Sub

Public Function Refresh() As String
'
End Function
