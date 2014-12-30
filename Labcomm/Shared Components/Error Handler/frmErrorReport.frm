VERSION 5.00
Begin VB.Form frmErrorReport 
   Caption         =   "Ice Error Report"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdInfo 
      Caption         =   "<< Hide"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtMore 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmErrorReport.frx":0000
      Top             =   3720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmErrorReport.frx":0008
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00C0FFFF&
      Height          =   840
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Error Number"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblProcPath 
      Caption         =   "Procedure"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblRunTime 
      Caption         =   "Runtime Information"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblProc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Procedure Identifier"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblErr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Error number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "An error has occurred. Please notify Anglia Healthcare on 01603 819600 when convenient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmErrorReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private eClass As errorControl

Private Sub cmdContinue_Click()
   errorEvaluated = False
   Unload Me
End Sub

Private Sub cmdInfo_Click()
   If cmdInfo.Caption = "Show >>" Then
      frmErrorReport.Caption = "IceConfig Error Report"
      frmErrorReport.Height = 5580
      cmdContinue.Top = 4560
      cmdInfo.Top = 4560
      cmdInfo.Caption = "<< Hide"
      lblInfo.Caption = "The full error path and details"
      lblProcPath.Caption = "Procedure - Click to view relevant details"
      lblProc.Visible = False
      lstErr.Visible = True
      lblDesc.Visible = True
      txtPath.Visible = True
      lblRunTime.Visible = True
      txtMore.Visible = True
      lstErr.ListIndex = 0
   Else
      frmErrorReport.Caption = "IceConfig Error Summary"
      frmErrorReport.Height = 2700
      cmdContinue.Top = 1800
      cmdInfo.Top = 1800
      cmdInfo.Caption = "Show >>"
      lblInfo.Caption = "An error has occurred. Please notify Anglia Healthcare on 01603 819600 when convenient"
      lblProcPath.Caption = "Procedure"
      lblProc.Visible = True
      lstErr.Visible = False
      lblDesc.Visible = False
      txtPath.Visible = False
      lblRunTime.Visible = False
      txtMore.Visible = False
   End If
End Sub

Public Property Let ErrorClass(objNewValue As errorControl)
   Set eClass = objNewValue
End Property

Private Sub Form_Load()
   Dim e As errorData
   Dim vData As Variant
   
   cmdInfo_Click
'   frmErrorReport.Height = 2640
'   cmdContinue.Top = 1800
'   cmdDetails.Top = 1800
'   cmdDetails.Caption = "Show >>"
'   lblInfo.Caption = "Unable to comply with your request. Click 'Continue' to carry on or 'Show' to see further details"
'   lstErr.Visible = False
'   txtMore.Visible = False
'   txtPath.Visible = False
   vData = eClass.ErrorDetails(1)
   lblProc.Caption = vData(0)
   lblErr.Caption = vData(2)
   For Each e In eClass
      vData = e.Retrieve
      lstErr.AddItem vData(0)
   Next
   
   frmErrorReport.Caption = eClass.FormCaption
   Set e = Nothing
End Sub

Private Sub lstErr_Click()
   Dim vData As Variant
   
   vData = eClass.ErrorDetails(lstErr.ListIndex + 1)
   lblErr.Caption = vData(2)
   txtMore.Text = vData(1)
   txtPath.Text = vData(3)
End Sub
