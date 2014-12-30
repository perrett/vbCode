VERSION 5.00
Begin VB.Form frmHistology 
   Caption         =   "Histology"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHisto 
      Caption         =   "Maintainence >>"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CheckBox chkHisto 
      Caption         =   "Input as text"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmHistology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

