VERSION 5.00
Begin VB.Form frmEDIRequeue 
   Caption         =   "Reprocess or Resend"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmEDIRequeue.frx":0000
      Top             =   1125
      Width           =   4575
   End
   Begin VB.Frame fraRequeue 
      Caption         =   "Select"
      Height          =   915
      Left            =   600
      TabIndex        =   0
      Top             =   105
      Width           =   3675
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2340
         TabIndex        =   5
         Top             =   540
         Width           =   1155
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   255
         Left            =   2340
         TabIndex        =   3
         Top             =   210
         Width           =   1155
      End
      Begin VB.OptionButton optRequeue 
         Caption         =   "Reprocess all reports"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1875
      End
      Begin VB.OptionButton optRequeue 
         Caption         =   "Resend File"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmEDIRequeue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RequeueType As String

Public Property Get RequeueValue() As String
   RequeueValue = RequeueType
   Me.Hide
End Property

Private Sub cmdCancel_Click()
   RequeueType = "Cancel"
   Me.Hide
End Sub

Private Sub cmdGo_Click()
   If optRequeue(0).value = True Then
      RequeueType = "Resend"
   Else
      RequeueType = "Reprocess"
   End If
   Me.Hide
End Sub
