VERSION 5.00
Begin VB.Form frmPanels 
   Caption         =   "Test Panel Management"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDetails 
      Height          =   1230
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame fraObject 
      Caption         =   "Panel Details"
      Height          =   2235
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.ComboBox comboPType 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label lblObject 
      Caption         =   "Current Panels"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private itemType As String

Public Property Let DisplayItem(strNewValue As String)
   itemType = strNewValue
End Property

Public Property Get DisplayItem() As String
   DisplayItem = itemType
End Property

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim panelType As Integer
   
   If comboPType.ListIndex = 0 Then
      panelType = "0"
   ElseIf comboPType.ListIndex = 1 Then
      panelType = 1
   End If
      
   If itemType = "Panels" Then
      strSQL = "SELECT * " & _
               "FROM Request_Panels " & _
               "WHERE PanelName = '" & txtName.Text & "'"
      RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount = 0 Then
         strSQL = "INSERT INTO Request_Panels " & _
                  "(PanelName, PanelType) " & _
                  "VALUES ('" & txtName.Text & "', " & panelType & ")"
         ICECon.Execute strSQL
         lstDetails.AddItem txtName.Text
      Else
         strSQL = "UPDATE Request_Panels " & _
                  "SET Panel_Name = '" & txtName.Text & "' " & _
                  "    Panel_Type = " & panelType & _
                  "WHERE PanelId = " & lstDetails.ItemData(lstDetails.ListIndex)
         ICECon.Execute strSQL
      End If
      RS.Close
      
   Else
      strSQL = "SELECT * " & _
               "FROM Requests_Panels_Pages " & _
               "WHERE PageName = '" & txtName.Text & "'"
      RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount = 0 Then
         strSQL = "INSERT INTO Request_Panels_Pages " & _
                  "(PanelID, PageName) " & _
                  "VALUES (" & frmPanels.Tag & "'" & txtName.Text & "')"
         ICECon.Execute strSQL
         lstDetails.AddItem txtName.Text
      Else
         MsgBox "Page already exixts for this panel", vbOKOnly, "Update error"
      End If
   End If
End Sub

Private Sub comboPType_Click()
   DisplayPanels
End Sub

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   
   If itemType = "Panels" Then
      comboPType.Clear
      comboPType.AddItem "Clinical Sciences"
      comboPType.AddItem "Radiology"
      comboPType.ListIndex = 0
      comboPType_Click
   Else
      strSQL = "SELECT * " & _
               "FROM Request_Panels_Pages " & _
               "WHERE PanelID = '" & itemType & "'"
      RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
      Do Until RS.EOF
         lstDetails.AddItem RS!PageName
         RS.MoveNext
      Loop
   End If
'   RS.Close
'   Set RS = Nothing
End Sub

Private Sub lstDetails_Click()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   txtName.Text = lstDetails.List(lstDetails.ListIndex)
   strSQL = "SELECT * " & _
            "FROM Request_Panels_Pages " & _
            "WHERE PanelId = " & lstDetails.ItemData(lstDetails.ListIndex)
End Sub

Private Sub DisplayPanels()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim pType As Long
   Dim i As Integer
   
   If comboPType.ListIndex = 0 Then
      pType = 1
   Else
      pType = 2
   End If
   strSQL = "SELECT * " & _
            "FROM Request_Panels " & _
            "WHERE PanelType = " & pType & _
            " ORDER BY PanelId"
   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
   lstDetails.Clear
   Do Until RS.EOF
      lstDetails.AddItem RS!panelName
      lstDetails.ItemData(i) = RS!PanelId
      RS.MoveNext
      i = i + 1
   Loop

End Sub
