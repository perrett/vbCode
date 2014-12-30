VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStartTime 
   Caption         =   "Start Time"
   ClientHeight    =   1515
   ClientLeft      =   6510
   ClientTop       =   2640
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   3840
   Begin VB.CheckBox chkEnable 
      Caption         =   "No start time"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker tPick 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "HH:mm"
      Format          =   45219843
      UpDown          =   -1  'True
      CurrentDate     =   37263
   End
End
Attribute VB_Name = "frmStartTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callOrigin As String

Private Sub chkEnable_Click()
   If chkEnable.Value = 0 Then
      tPick.Enabled = True
   Else
      tPick.Enabled = False
   End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub CmdOk_Click()
   If chkEnable.Value = 0 Then
      frmMain.edipr(callOrigin).Value = Format(tPick.Value, "hh:nn")
   Else
      frmMain.edipr(callOrigin).Value = ""
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   callOrigin = frmMain.edipr.Tag
End Sub
