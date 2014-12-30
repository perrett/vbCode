VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChart 
   Caption         =   "Graph"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   12705
   StartUpPosition =   1  'CenterOwner
   Begin MSChart20Lib.MSChart chrtStats 
      Height          =   8070
      Left            =   390
      OleObjectBlob   =   "frmChart.frx":0000
      TabIndex        =   1
      Top             =   105
      Width           =   11595
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   5610
      TabIndex        =   0
      Top             =   8415
      Width           =   1335
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
End Sub
