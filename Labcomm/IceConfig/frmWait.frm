VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1020
   ClientLeft      =   6135
   ClientTop       =   5220
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrBusy 
      Enabled         =   0   'False
      Left            =   75
      Top             =   45
   End
   Begin VB.Frame fraBusy 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   750
      TabIndex        =   2
      Top             =   390
      Visible         =   0   'False
      Width           =   2310
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   120
         Picture         =   "frmWait.frx":0000
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   12
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   750
         Picture         =   "frmWait.frx":003D
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   11
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   330
         Picture         =   "frmWait.frx":007A
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   10
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   540
         Picture         =   "frmWait.frx":00B7
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   9
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   9
         Left            =   2010
         Picture         =   "frmWait.frx":00F4
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   8
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   1800
         Picture         =   "frmWait.frx":0131
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   7
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   7
         Left            =   1590
         Picture         =   "frmWait.frx":016E
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   6
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   960
         Picture         =   "frmWait.frx":01AB
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   5
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   1170
         Picture         =   "frmWait.frx":01E8
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   4
         Top             =   60
         Width           =   150
      End
      Begin VB.PictureBox pb1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   1380
         Picture         =   "frmWait.frx":0225
         ScaleHeight     =   180
         ScaleWidth      =   150
         TabIndex        =   3
         Top             =   60
         Width           =   150
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   330
      TabIndex        =   0
      Top             =   405
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Interrogating Database..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   3600
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit

Private lastSet As Integer

Private Sub Form_Activate()
    Me.Refresh
End Sub

Private Sub Form_Load()
    Me.Refresh
    tmrBusy.Interval = 100
End Sub

Private Sub Switch(id As Integer)
   Dim i As Integer
   
   For i = 0 To 9
      pb1(i).Visible = False
   Next i
   
   pb1(id).Visible = True
   
   Select Case id
      Case 8
         pb1(9).Visible = True
         pb1(0).Visible = True
      
      Case 9
         pb1(0).Visible = True
         pb1(1).Visible = True
         
         
      Case Else
         pb1(id + 1).Visible = True
         pb1(id + 2).Visible = True
      
   End Select
   
End Sub

Private Sub tmrBusy_Timer()
   If lastSet > 9 Then
      lastSet = 0
   End If
   
   Switch (lastSet)
   lastSet = lastSet + 1
End Sub
