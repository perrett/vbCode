VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   4200
   ClientTop       =   3555
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6000
      Top             =   120
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   131073
      ForeColor       =   16777215
      BackColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSplash.frx":0000
      Caption         =   "Member of the ICE family of products           "
      BevelOuter      =   0
      Alignment       =   5
      PictureAlignment=   4
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   5160
      Picture         =   "frmSplash.frx":031A
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   4080
      Picture         =   "frmSplash.frx":0BE4
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   3000
      Picture         =   "frmSplash.frx":14AE
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   1920
      Picture         =   "frmSplash.frx":1D78
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   840
      Picture         =   "frmSplash.frx":2082
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "See Help About for further information."
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   4440
      Width           =   2700
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3360
      Width           =   4005
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Sunquest Systems Ltd, 2000-2009."
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   4965
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ICE..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1950
   End
   Begin VB.Image Image6 
      Height          =   825
      Left            =   360
      Picture         =   "frmSplash.frx":294C
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private First As Boolean

Private Sub Form_Load()
   Label6.Caption = Label6.Caption + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
   First = True
End Sub

Private Sub Timer1_Timer()
   If First Then
'      If sqlServer.VerifyConnection(SQLDMOConn_CurrentState) = False Then
'         sqlServer.ApplicationName = "IceConfig"
'         sqlServer.EnableBcp = True
'         sqlServer.Connect UDLServer, dbUser, dbPass
'
'         Set sqlDb = sqlServer.Databases(UDLDatabase)
'      End If
      
      Set rCtrl = New requeueControl
      
      frmMain.Show
      First = False
   Else
       Unload Me
   End If
End Sub
