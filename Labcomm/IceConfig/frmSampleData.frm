VERSION 5.00
Begin VB.Form frmSampleData 
   Caption         =   "Sample Mapping data"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDesc 
      Caption         =   "Amending Code Description"
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5895
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Description"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Description"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   2130
         Left            =   1200
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtNatCode 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Existing Descriptions"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Text"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   4320
      Width           =   1695
   End
End
Attribute VB_Name = "frmSampleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
   Combo1.AddItem Combo1.Text
   Combo1.Text = ""
   cmdAdd.Enabled = False
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   Combo1.RemoveItem Combo1.ListIndex
   cmdDelete.Enabled = False
End Sub

Private Sub CmdOk_Click()
'   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim strSQL As String

'  First delete all the existing records
   strSQL = "DELETE FROM  " & frmSampleData.Tag & _
                   " WHERE National_Code = '" & txtNatCode.Text & "'"
   iceCon.Execute strSQL
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockOptimistic
'  Now add the items in the combo box
   For i = 0 To Combo1.ListCount - 1
      strSQL = "INSERT INTO " & frmSampleData.Tag & " (National_Code, Local_Text) " & _
                 "VALUES ('" & txtNatCode.Text & "','" & Combo1.List(i) & "')"
      iceCon.Execute strSQL
'      RS.Open strSQL, ICECon, adOpenKeyset, adLockOptimistic
   Next i
'   RS.Close
'   Set RS = Nothing
   Unload Me
End Sub

Private Sub Combo1_Change()
   cmdAdd.Enabled = True
   cmdDelete.Enabled = False
End Sub

Private Sub Combo1_Click()
   cmdDelete.Enabled = True
   cmdAdd.Enabled = False
End Sub

Private Sub Form_Activate()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   frmSampleData.Caption = frmMain.edipr("SD").value
   Combo1.Clear
   strSQL = "SELECT * FROM " & frmSampleData.Tag & " WHERE National_Code = '" & _
              txtNatCode.Text & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockOptimistic
   Do Until RS.EOF
      Combo1.AddItem RS!Local_Text
      RS.MoveNext
   Loop
   If Combo1.ListCount > 0 Then
      Combo1.ListIndex = 0
   End If
   RS.Close
   Set RS = Nothing
End Sub

