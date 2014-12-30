VERSION 5.00
Begin VB.Form frmDeleteData 
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1965
      TabIndex        =   2
      Top             =   675
      Width           =   1395
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Default         =   -1  'True
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Top             =   690
      Width           =   1395
   End
   Begin VB.ComboBox comboData 
      Height          =   315
      Left            =   690
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "frmDeleteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private tableId As String

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   Dim RS As New ADODB.Recordset
   
   If tableId = "Request_Panels" Then
      If MsgBox("This will delete ALL the pages associated with this panel. Are you sure?", vbYesNo, "IceConfig Deletion Warning") = vbYes Then
         strSQL = "SELECT PanelID " & _
                  "FROM " & tableId & _
                  " WHERE PanelName = '" & comboData.Text & "'"
         RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
         strSQL = "DELETE FROM Request_Panels_PAges " & _
                  "WHERE PanelID = " & RS!PanelId
         ICECon.Execute strSQL
         strSQL = "DELETE FROM " & tableId & _
                  " WHERE PanelName = '" & comboData.Text & "'"
         ICECon.Execute strSQL
         RS.Close
      End If
   Else
      strSQL = "DELETE FROM " & tableId & _
               " WHERE Page_Name = '" & comboData.Text & "'"
      ICECon.Execute strSQL
   End If
   Set RS = Nothing
   Unload Me
End Sub

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   
   comboData.Clear
   strSQL = "SELECT * " & _
            "FROM " & tableId
   
   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      If tableId = "Request_Panels" Then
         comboData.AddItem RS!panelName
      Else
         comboData.AddItem RS!PageName
      End If
      RS.MoveNext
   Loop
   comboData.ListIndex = 0
   RS.Close
   Set RS = Nothing
End Sub

Public Property Let DbTable(strNewValue As String)
   tableId = strNewValue
End Property
