VERSION 5.00
Object = "{3F118CA4-97A8-4926-AA0A-6FD80DFB3DCE}#2.6#0"; "PrpList2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form dummyFrmProf 
   Caption         =   "Form2"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   LinkTopic       =   "Form2"
   ScaleHeight     =   7215
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraProfile 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7005
      Left            =   -30
      TabIndex        =   0
      Top             =   30
      Width           =   8025
      Begin VB.CommandButton Command13 
         Caption         =   "Add New Test"
         Height          =   375
         Left            =   1950
         TabIndex        =   2
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Delete Test"
         Height          =   375
         Left            =   3990
         TabIndex        =   1
         Top             =   6240
         Width           =   1935
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   5415
         Left            =   390
         TabIndex        =   3
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9551
         _Version        =   131073
         ForeColor       =   16777215
         BackColor       =   8421504
         Caption         =   "Profile Properties"
         Begin VB.CommandButton Command15 
            Caption         =   "Delete This Profile"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   5040
            Width           =   2535
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Left            =   5640
            TabIndex        =   5
            Top             =   5040
            Width           =   1335
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4200
            TabIndex        =   4
            Top             =   5040
            Width           =   1335
         End
         Begin PropertiesListCtl.PropertiesList PR1 
            Height          =   4455
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   7858
            LicenceData     =   "00205E203743243B54231E58232611133D5F24205E3C727335205F39371100374322374524"
            DescriptionHeight=   35
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Ownerdraw       =   1
         End
      End
   End
End
Attribute VB_Name = "dummyFrmProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

