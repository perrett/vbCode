VERSION 5.00
Begin VB.Form frmCustomStats 
   Caption         =   "Custom Statistics"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Values"
      Height          =   1575
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   1815
      Begin VB.CheckBox chkPNull 
         Alignment       =   1  'Right Justify
         Caption         =   "Set to Null"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "txtParam"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox lstPValue 
      Height          =   60
      Left            =   2040
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstParams 
      Height          =   645
      ItemData        =   "frmCustomStats.frx":0000
      Left            =   240
      List            =   "frmCustomStats.frx":0002
      TabIndex        =   0
      ToolTipText     =   "Click to select"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Parameters"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblProcName 
      Alignment       =   2  'Center
      Caption         =   "Stored Procedure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmCustomStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private spName As String

Public Property Let StoredProc(strNewValue As String)
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   lstParams.Clear
   lstPValue.Clear
   txtParam.Text = ""
   
   spName = strNewValue
   strSQL = "SELECT name, colorder, xType, Length " & _
            "FROM syscolumns " & _
            "WHERE id = " & spName & _
            " ORDER BY colorder"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Do Until RS.EOF
      lstParams.AddItem RS!Name
      lstParams.ItemData(lstParams.ListCount - 1) = RS!xType
      lstPValue.AddItem "-1"
      lstPValue.ItemData(lstParams.ListCount - 1) = RS!Length
      
      RS.MoveNext
   Loop
   
   RS.Close
   Set RS = Nothing
   Me.Show 1
End Property

Private Sub chkPNull_Click()
   If chkPNull.value > 0 Then
      txtParam.Text = ""
   End If
End Sub

Private Sub cmdCancel_Click()
   Me.Tag = "Cancel"
   Me.Hide
End Sub

Private Sub CmdOk_Click()
   Dim blnError As Boolean
   Dim i As Integer
   
   For i = 0 To lstPValue.ListCount - 1
      If lstPValue.List(i) = "-1" Then
         MsgBox lstParams.List(i) & " - Value not set", vbExclamation, "Parameter Error"
         blnError = True
      End If
   Next i
   
   If blnError = False Then
      Me.Tag = ""
      Me.Hide
   End If
End Sub

Private Sub lstParams_Click()
   If lstPValue.List(lstParams.ListIndex) = "-1" Then
      txtParam.Text = ""
   Else
      txtParam.Text = lstPValue.List(lstParams.ListIndex)
   End If
End Sub

Private Sub txtParam_Change()
   lstPValue.List(lstParams.ListIndex) = txtParam.Text
End Sub
