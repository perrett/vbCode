VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmShowRequeue 
   Caption         =   "Requeue Confirmation"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid fgReq 
      Height          =   4935
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      HighLight       =   0
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.TextBox txtNote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txtDuplicate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtRequeued 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCommit 
      BackColor       =   &H0080FF80&
      Caption         =   "Commit"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lblNote 
      Caption         =   "Already on Rep List"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Duplicated"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Requeued"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frmShowRequeue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ButtonPushed As String

Private Sub cmdCancel_Click()
   ButtonPushed = "Cancel"
   Me.Hide
End Sub

Private Sub cmdCommit_Click()
   ButtonPushed = "Commit"
   Me.Hide
End Sub

Public Property Get FormAction() As String
   FormAction = ButtonPushed
End Property

Public Sub Resize(cWidth1 As Long, _
                  cWidth2 As Long, _
                  cWidth3 As Long)
   Dim exWidth As Long
   Dim flexWidth As Long
   Dim frmwidth As Long
   Dim ratio As Long
   
   exWidth = 150
   If cWidth1 < 3500 Then
      cWidth1 = 3500
   End If
   
   If cWidth2 < 1500 Then
      cWidth2 = 1500
   End If
   
   If cWidth3 < 2000 Then
      cWidth3 = 2000
   End If
   
   With fgReq
      .ColWidth(0) = cWidth1 + exWidth
      .ColWidth(1) = cWidth2 + exWidth
      .ColWidth(2) = cWidth3 + exWidth
      flexWidth = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 500
   End With
   
   frmwidth = flexWidth + 240
   If frmwidth < 7400 Then
      frmwidth = 7400
   End If
   
   With cmdCommit
      .Left = (.Left / frmShowRequeue.Width) * frmwidth
   End With
   
   With cmdCancel
      .Left = (.Left / frmShowRequeue.Width) * frmwidth
   End With
   
   With txtRequeued
      .Left = (.Left / frmShowRequeue.Width) * frmwidth
   End With
   
   With txtDuplicate
      .Left = (.Left / frmShowRequeue.Width) * frmwidth
   End With
   
   With txtNote
      .Left = (.Left / frmShowRequeue.Width) * frmwidth
   End With
   
   Me.Width = frmwidth
   fgReq.Width = flexWidth
   
End Sub

Private Sub Form_Load()
   fgReq.ColAlignment(0) = flexAlignLeftCenter
   fgReq.FillStyle = flexFillRepeat
End Sub
