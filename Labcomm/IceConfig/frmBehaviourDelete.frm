VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalgrid6.ocx"
Begin VB.Form frmBehaviourDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Behaviour"
   ClientHeight    =   4920
   ClientLeft      =   1635
   ClientTop       =   1935
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6588
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      DisableIcons    =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select the rule you wish to delete from the selected test and click the delete button."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmBehaviourDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If vbalGrid1.SelectedRow < 1 Then Exit Sub
    If MsgBox("Are you sure you wish to delete the rule '" & vbalGrid1.Cell(vbalGrid1.SelectedRow, 1).Text & "' from the test '" & frmMain.TreeView1.SelectedItem.Parent.Text & "'?", vbYesNo + vbQuestion, "Delete Rule") = vbYes Then frmMain.TreeView1.Nodes.Remove vbalGrid1.Cell(vbalGrid1.SelectedRow, 2).Text
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    vbalGrid1.AutoWidthColumn "COL1"
End Sub

