VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPanel 
      BackColor       =   &H00808080&
      Height          =   6945
      Index           =   4
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   8355
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1335
         Left            =   840
         TabIndex        =   7
         Top             =   2520
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2355
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   1
         TextRTF         =   $"incexrules.frx":0000
      End
      Begin VB.TextBox txtTestDesc 
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Index           =   0
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "incexrules.frx":008B
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtTestDesc 
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Index           =   2
         Left            =   4560
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "incexrules.frx":0225
         Top             =   2160
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtTestDesc 
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Index           =   1
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "incexrules.frx":03BE
         Top             =   2160
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtTestDisplay 
         BackColor       =   &H00C0FFFF&
         Height          =   1935
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   6585
      End
      Begin VB.TextBox txtTestDesc 
         BackColor       =   &H00C0FFFF&
         Height          =   1935
         Index           =   3
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "incexrules.frx":0460
         Top             =   3960
         Visible         =   0   'False
         Width           =   6585
      End
      Begin VB.Label AddMode 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5220
         TabIndex        =   6
         Top             =   3420
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
