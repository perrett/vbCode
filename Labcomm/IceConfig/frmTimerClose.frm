VERSION 5.00
Begin VB.Form frmTimerClose 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Abandon in:"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbandon 
      Cancel          =   -1  'True
      Caption         =   "Abandon"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   840
   End
   Begin VB.Label labDesc 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTimerClose.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmTimerClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TimeLeft As Integer
Private TimeoutAfter As Integer
Private TimeRemaining As Integer

Public Property Let TimeoutValue(intNewValue As Integer)
   TimeoutAfter = intNewValue
   TimeRemaining = TimeoutAfter
End Property

Private Sub cmdAbandon_Click()
   transId.TimeOutExpired
End Sub

Private Sub cmdContinue_Click()
   transId.ResetTimer
End Sub

Private Sub Form_Load()
   Timer1.Interval = 1000
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   If TimeRemaining > 0 Then
      If TimeRemaining < 6 Then
         If Me.Visible = False Then
            transId.TimeoutWarning
         End If
      End If
      TimeRemaining = TimeRemaining - 1
      frmTimerClose.Caption = "Timeout in " & TimeRemaining & " Seconds"
   Else
      frmMain.edipr.Refresh True
      transId.TimeOutExpired
   End If
End Sub
