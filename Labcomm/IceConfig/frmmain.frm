VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{15138B51-7EB6-11D0-9BB7-0000C0F04C96}#1.0#0"; "SSLstBar.ocx"
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "SPLITTER.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{2D47C3AF-9C7A-44E4-9FCC-CDE675667A2D}#1.0#0"; "Web_Browser_Control.ocx"
Object = "{3F118CA4-97A8-4926-AA0A-6FD80DFB3DCE}#2.6#0"; "PrpList2.ocx"
Begin VB.Form frmMain 
   Caption         =   "ICE...Configuration"
   ClientHeight    =   10812
   ClientLeft      =   60
   ClientTop       =   816
   ClientWidth     =   16152
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10812
   ScaleWidth      =   16152
   Begin VB.Timer Timer1 
      Left            =   15030
      Top             =   630
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15000
      Top             =   1170
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":030A
            Key             =   "Tubes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0BE4
            Key             =   "Cogs"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":14BE
            Key             =   "No"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":17D8
            Key             =   "User"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":20B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":298C
            Key             =   "Connect"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2CA6
            Key             =   "Book"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2FC0
            Key             =   "Chart"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":32DA
            Key             =   "Syringe"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3BB4
            Key             =   "Mask"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3ECE
            Key             =   "Toilet"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":41E8
            Key             =   "Compass"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4502
            Key             =   "Scroll"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":481C
            Key             =   "SmallReport"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4B36
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4E50
            Key             =   "Calendar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":516A
            Key             =   "World"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5484
            Key             =   "BlueIce"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":579E
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5AB8
            Key             =   "PhoneFolder"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6392
            Key             =   "Skeleton"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":646B
            Key             =   "UK"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6785
            Key             =   "LabelList"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6A9F
            Key             =   "Clipboard"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6DB9
            Key             =   "Query"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":70D3
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":73ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7CC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7FE1
            Key             =   "Envelope"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":84DA
            Key             =   "eyeFolder"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":87F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":90CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A7FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":AB14
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":AE2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B338
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B9CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10812
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16152
      _ExtentX        =   28490
      _ExtentY        =   19071
      _Version        =   131073
      AutoSize        =   1
      SplitterBarJoinStyle=   0
      ClipControls    =   -1  'True
      PaneTree        =   "frmmain.frx":BF51
      Begin Threed.SSPanel mainPanel 
         Height          =   10212
         Left            =   5772
         TabIndex        =   9
         Top             =   468
         Width           =   10356
         _ExtentX        =   18267
         _ExtentY        =   18013
         _Version        =   131073
         ForeColor       =   16777215
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin Web_Browser_Control_V2.wbCtrl wb 
            Height          =   7812
            Left            =   1440
            TabIndex        =   107
            Top             =   480
            Width           =   7692
            _ExtentX        =   13568
            _ExtentY        =   13780
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H8000000C&
            Caption         =   "Read Code Errors"
            ForeColor       =   &H8000000E&
            Height          =   1545
            Index           =   12
            Left            =   6120
            TabIndex        =   78
            Top             =   7680
            Visible         =   0   'False
            Width           =   2715
         End
         Begin MSComctlLib.ImageList ImgLarge 
            Left            =   9390
            Top             =   1380
            _ExtentX        =   995
            _ExtentY        =   995
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmmain.frx":C023
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmmain.frx":C33D
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Height          =   1005
            Index           =   0
            Left            =   6720
            TabIndex        =   47
            Top             =   5160
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H8000000C&
            Caption         =   "Statistics"
            ForeColor       =   &H8000000E&
            Height          =   675
            Index           =   9
            Left            =   240
            TabIndex        =   15
            Top             =   150
            Visible         =   0   'False
            Width           =   1500
            Begin VB.CommandButton cmdStats 
               Caption         =   "Show Statistics"
               Height          =   330
               Left            =   2805
               TabIndex        =   100
               Top             =   1320
               Width           =   2415
            End
            Begin VB.ComboBox cboStatsPractice 
               Height          =   315
               Left            =   1110
               TabIndex        =   104
               Text            =   "Combo1"
               Top             =   780
               Width           =   1800
            End
            Begin MSDataGridLib.DataGrid xdgStats 
               Height          =   375
               Left            =   7635
               TabIndex        =   98
               Top             =   195
               Visible         =   0   'False
               Width           =   390
               _ExtentX        =   699
               _ExtentY        =   656
               _Version        =   393216
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Data Here"
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.CommandButton cmdGraph 
               BackColor       =   &H8000000B&
               Caption         =   "Graph"
               Default         =   -1  'True
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   210
               TabIndex        =   97
               Top             =   4020
               Width           =   1620
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H8000000C&
               Caption         =   "Date Range"
               ForeColor       =   &H8000000E&
               Height          =   1035
               Left            =   3120
               TabIndex        =   91
               Top             =   120
               Width           =   4335
               Begin VB.OptionButton optStatPeriod 
                  BackColor       =   &H8000000C&
                  Caption         =   "Back to..."
                  ForeColor       =   &H8000000E&
                  Height          =   225
                  Index           =   3
                  Left            =   840
                  TabIndex        =   95
                  Top             =   645
                  Width           =   1080
               End
               Begin VB.OptionButton optStatPeriod 
                  BackColor       =   &H8000000C&
                  Caption         =   "Last 28 Days"
                  ForeColor       =   &H8000000E&
                  Height          =   225
                  Index           =   2
                  Left            =   2805
                  TabIndex        =   94
                  Top             =   285
                  Width           =   1305
               End
               Begin VB.OptionButton optStatPeriod 
                  BackColor       =   &H8000000C&
                  Caption         =   "Last 7 Days"
                  ForeColor       =   &H8000000E&
                  Height          =   225
                  Index           =   1
                  Left            =   1545
                  TabIndex        =   93
                  Top             =   285
                  Value           =   -1  'True
                  Width           =   1305
               End
               Begin VB.OptionButton optStatPeriod 
                  BackColor       =   &H8000000C&
                  Caption         =   "This day only"
                  ForeColor       =   &H8000000E&
                  Height          =   225
                  Index           =   0
                  Left            =   150
                  TabIndex        =   92
                  Top             =   285
                  Width           =   1305
               End
               Begin MSComCtl2.DTPicker DTStatEnd 
                  Height          =   345
                  Left            =   2160
                  TabIndex        =   99
                  Top             =   570
                  Width           =   1515
                  _ExtentX        =   2667
                  _ExtentY        =   614
                  _Version        =   393216
                  CustomFormat    =   "ddd dd MMM yyyy"
                  Format          =   94437377
                  CurrentDate     =   37505
               End
            End
            Begin MSComCtl2.DTPicker DTStats 
               Height          =   345
               Left            =   1110
               TabIndex        =   90
               Top             =   270
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   614
               _Version        =   393216
               CustomFormat    =   "ddd dd MMM yyyy"
               Format          =   94437377
               CurrentDate     =   37504
            End
            Begin VB.Label Label5 
               BackColor       =   &H8000000C&
               Caption         =   "Practice"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   195
               TabIndex        =   103
               Top             =   795
               Width           =   690
            End
            Begin VB.Label Label3 
               BackColor       =   &H8000000C&
               Caption         =   "Start From"
               ForeColor       =   &H8000000E&
               Height          =   210
               Left            =   210
               TabIndex        =   96
               Top             =   330
               Width           =   855
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H8000000C&
            Caption         =   "Sample Colour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1080
            Index           =   13
            Left            =   525
            TabIndex        =   16
            Top             =   7950
            Width           =   2370
            Begin VB.Label labShowCol 
               BackColor       =   &H8000000C&
               BorderStyle     =   1  'Fixed Single
               Height          =   495
               Left            =   330
               TabIndex        =   24
               Top             =   375
               Width           =   3180
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Caption         =   "Search Criteria"
            ForeColor       =   &H8000000E&
            Height          =   1125
            Index           =   6
            Left            =   405
            TabIndex        =   48
            Top             =   2310
            Visible         =   0   'False
            Width           =   3000
            Begin VB.Frame fraShowErr 
               BackColor       =   &H8000000C&
               Caption         =   "Display"
               ForeColor       =   &H8000000E&
               Height          =   885
               Left            =   2295
               TabIndex        =   85
               Top             =   1890
               Width           =   2520
               Begin VB.OptionButton optShowErr 
                  BackColor       =   &H8000000C&
                  Caption         =   "OK Only"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   3
                  Left            =   1410
                  TabIndex        =   89
                  Top             =   600
                  Width           =   960
               End
               Begin VB.OptionButton optShowErr 
                  BackColor       =   &H8000000C&
                  Caption         =   "Warnings"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   88
                  Top             =   600
                  Width           =   1005
               End
               Begin VB.OptionButton optShowErr 
                  BackColor       =   &H8000000C&
                  Caption         =   "Errors"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   87
                  Top             =   255
                  Width           =   810
               End
               Begin VB.OptionButton optShowErr 
                  BackColor       =   &H8000000C&
                  Caption         =   "All"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Index           =   0
                  Left            =   135
                  TabIndex        =   86
                  Top             =   255
                  Width           =   975
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00808080&
               Caption         =   "Options"
               ForeColor       =   &H8000000E&
               Height          =   2415
               Left            =   300
               TabIndex        =   71
               Top             =   285
               Width           =   1785
               Begin VB.OptionButton optLogSearch 
                  BackColor       =   &H00808080&
                  Caption         =   "Date Only"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   76
                  Top             =   360
                  Width           =   1365
               End
               Begin VB.OptionButton optLogSearch 
                  BackColor       =   &H00808080&
                  Caption         =   "Lab Report Id"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   75
                  Top             =   2010
                  Width           =   1365
               End
               Begin VB.OptionButton optLogSearch 
                  BackColor       =   &H00808080&
                  Caption         =   "By NHS/Hosp No."
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   74
                  Top             =   1596
                  Width           =   1650
               End
               Begin VB.OptionButton optLogSearch 
                  BackColor       =   &H00808080&
                  Caption         =   "By Patient"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   73
                  Top             =   1184
                  Width           =   1365
               End
               Begin VB.OptionButton optLogSearch 
                  BackColor       =   &H00808080&
                  Caption         =   "By Practice"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   72
                  Top             =   772
                  Width           =   1365
               End
            End
            Begin VB.CommandButton cmdSrchOk 
               Caption         =   "OK"
               Height          =   315
               Left            =   390
               TabIndex        =   70
               Top             =   3045
               Width           =   1785
            End
            Begin VB.CommandButton cmdSrchCancel 
               Caption         =   "Cancel"
               Height          =   315
               Left            =   2670
               TabIndex        =   69
               Top             =   3045
               Width           =   1695
            End
            Begin VB.Frame fraSCrit 
               BackColor       =   &H00808080&
               Caption         =   "Select the practice"
               ForeColor       =   &H8000000E&
               Height          =   495
               Index           =   1
               Left            =   5040
               TabIndex        =   66
               Top             =   1680
               Width           =   1665
               Begin VB.ComboBox ComboSrchPractice 
                  Height          =   315
                  ItemData        =   "frmmain.frx":C657
                  Left            =   1320
                  List            =   "frmmain.frx":C659
                  Style           =   2  'Dropdown List
                  TabIndex        =   67
                  Top             =   360
                  Width           =   2475
               End
               Begin VB.Label Label2 
                  BackColor       =   &H8000000C&
                  Caption         =   "Practice"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Left            =   600
                  TabIndex        =   68
                  Top             =   360
                  Width           =   855
               End
            End
            Begin VB.Frame fraSCrit 
               BackColor       =   &H00808080&
               Caption         =   "Enter Patient Name"
               ForeColor       =   &H8000000E&
               Height          =   375
               Index           =   2
               Left            =   5160
               TabIndex        =   61
               Top             =   1080
               Width           =   1785
               Begin VB.TextBox txtSrchSurname 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.TextBox txtSrchForename 
                  Height          =   315
                  Left            =   3000
                  TabIndex        =   63
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Surname"
                  ForeColor       =   &H8000000E&
                  Height          =   225
                  Left            =   120
                  TabIndex        =   65
                  Top             =   360
                  Width           =   705
               End
               Begin VB.Label Label11 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Forename"
                  ForeColor       =   &H8000000E&
                  Height          =   195
                  Left            =   2160
                  TabIndex        =   64
                  Top             =   360
                  Width           =   795
               End
            End
            Begin VB.Frame fraSCrit 
               BackColor       =   &H00808080&
               Caption         =   "Enter Lab Report ID"
               ForeColor       =   &H8000000E&
               Height          =   375
               Index           =   4
               Left            =   5040
               TabIndex        =   58
               Top             =   480
               Width           =   1905
               Begin VB.TextBox txtLogSearchLab 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   59
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lab Number"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Left            =   600
                  TabIndex        =   60
                  Top             =   360
                  Width           =   885
               End
            End
            Begin VB.Frame fraSCrit 
               BackColor       =   &H00808080&
               Caption         =   "Enter NHS or Hospital No."
               ForeColor       =   &H8000000E&
               Height          =   375
               Index           =   3
               Left            =   4560
               TabIndex        =   55
               Top             =   2640
               Width           =   2265
               Begin VB.TextBox txtSrchNHS 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   56
                  Top             =   360
                  Width           =   1500
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  'Transparent
                  Caption         =   "NHS or Hospital No."
                  ForeColor       =   &H8000000E&
                  Height          =   435
                  Left            =   720
                  TabIndex        =   57
                  Top             =   360
                  Width           =   885
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H8000000C&
               Caption         =   "Dates"
               ForeColor       =   &H8000000E&
               Height          =   1470
               Left            =   2295
               TabIndex        =   49
               Top             =   285
               Width           =   2535
               Begin VB.CheckBox chkRestrict 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000C&
                  Caption         =   "Restrict Search by date"
                  ForeColor       =   &H8000000E&
                  Height          =   225
                  Left            =   225
                  TabIndex        =   50
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   2055
               End
               Begin MSComCtl2.DTPicker dtPTo 
                  Height          =   375
                  Left            =   810
                  TabIndex        =   51
                  Top             =   870
                  Width           =   1455
                  _ExtentX        =   2561
                  _ExtentY        =   656
                  _Version        =   393216
                  Format          =   94437377
                  CurrentDate     =   37075
               End
               Begin MSComCtl2.DTPicker dtPFrom 
                  Height          =   375
                  Left            =   810
                  TabIndex        =   52
                  Top             =   315
                  Width           =   1455
                  _ExtentX        =   2561
                  _ExtentY        =   656
                  _Version        =   393216
                  Format          =   94437377
                  CurrentDate     =   37075
               End
               Begin VB.Label Label8 
                  BackStyle       =   0  'Transparent
                  Caption         =   "From"
                  ForeColor       =   &H8000000E&
                  Height          =   255
                  Left            =   210
                  TabIndex        =   54
                  Top             =   435
                  Width           =   495
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "To"
                  ForeColor       =   &H8000000E&
                  Height          =   225
                  Left            =   210
                  TabIndex        =   53
                  Top             =   990
                  Width           =   315
               End
            End
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   5000
            Left            =   9870
            Top             =   150
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Height          =   1005
            Index           =   1
            Left            =   2520
            TabIndex        =   42
            Top             =   240
            Visible         =   0   'False
            Width           =   1785
            Begin VB.CommandButton Command4 
               Caption         =   "Delete Rule"
               Height          =   375
               Left            =   4020
               TabIndex        =   45
               Top             =   3750
               Width           =   1935
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Add New Rule"
               Height          =   375
               Left            =   1050
               TabIndex        =   44
               Top             =   3750
               Width           =   1935
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Height          =   3105
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   43
               Text            =   "frmmain.frx":C65B
               Top             =   330
               Width           =   7185
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Caption         =   "Congfiguration Properties"
            ForeColor       =   &H8000000E&
            Height          =   885
            Index           =   2
            Left            =   5520
            TabIndex        =   36
            Top             =   120
            Visible         =   0   'False
            Width           =   2205
            Begin VB.TextBox Text2 
               BackColor       =   &H00C0FFFF&
               Height          =   1185
               Left            =   450
               MultiLine       =   -1  'True
               TabIndex        =   39
               Text            =   "frmmain.frx":CB05
               Top             =   405
               Width           =   7215
            End
            Begin VB.CommandButton Command19 
               Caption         =   "Cancel"
               Height          =   255
               Left            =   5430
               TabIndex        =   38
               Top             =   3840
               Width           =   1335
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Save"
               Height          =   255
               Left            =   3210
               TabIndex        =   37
               Top             =   3840
               Width           =   1335
            End
            Begin VB.Label CfgName 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   480
               TabIndex        =   41
               Top             =   3840
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Caption         =   "Amend as required"
            ForeColor       =   &H8000000E&
            Height          =   6645
            Index           =   3
            Left            =   960
            TabIndex        =   34
            Top             =   1320
            Visible         =   0   'False
            Width           =   5880
            Begin PropertiesListCtl.PropertiesList EdiPr 
               Height          =   972
               Left            =   360
               TabIndex        =   108
               Top             =   360
               Width           =   1332
               _ExtentX        =   2350
               _ExtentY        =   1715
               LicenceData     =   "15160620173710003501090D1A2452060B0724002A0849001330084903022416023501"
               DescriptionHeight=   44
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.CommandButton ssCmdEdiCancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   4200
               TabIndex        =   105
               Top             =   5280
               Width           =   1455
            End
            Begin VB.CommandButton ssCmdEDIok 
               Caption         =   "Ok"
               Height          =   375
               Left            =   1800
               TabIndex        =   106
               Top             =   5280
               Width           =   1335
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Caption         =   "File Details"
            ForeColor       =   &H8000000E&
            Height          =   1320
            Index           =   5
            Left            =   225
            TabIndex        =   31
            Top             =   1095
            Visible         =   0   'False
            Width           =   2175
            Begin VB.CommandButton OpenFileBtn 
               Caption         =   "Open File in Notepad"
               Height          =   405
               Left            =   2880
               TabIndex        =   77
               Top             =   7950
               Width           =   2175
            End
            Begin VB.CommandButton MailBtn 
               BackColor       =   &H0080FF80&
               Caption         =   "Requeue"
               Height          =   405
               Left            =   570
               MaskColor       =   &H0080FF80&
               TabIndex        =   32
               Top             =   7950
               Visible         =   0   'False
               Width           =   2175
            End
            Begin RichTextLib.RichTextBox LogText 
               Height          =   7290
               Left            =   360
               TabIndex        =   33
               Top             =   510
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   12869
               _Version        =   393217
               BackColor       =   12648447
               Enabled         =   -1  'True
               ReadOnly        =   -1  'True
               ScrollBars      =   3
               Appearance      =   0
               RightMargin     =   1.00000e5
               AutoVerbMenu    =   -1  'True
               OLEDragMode     =   0
               OLEDropMode     =   0
               TextRTF         =   $"frmmain.frx":CCE9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            ForeColor       =   &H8000000E&
            Height          =   885
            Index           =   7
            Left            =   8640
            TabIndex        =   25
            Top             =   2220
            Visible         =   0   'False
            Width           =   1545
            Begin VB.TextBox Text4 
               BackColor       =   &H00C0FFFF&
               Height          =   1365
               Left            =   360
               MultiLine       =   -1  'True
               TabIndex        =   30
               Text            =   "frmmain.frx":CD69
               Top             =   480
               Width           =   7125
            End
            Begin VB.Frame fraPickOpt 
               BackColor       =   &H00808080&
               Caption         =   "Picklist Options"
               ForeColor       =   &H00FFFFFF&
               Height          =   1695
               Left            =   480
               TabIndex        =   26
               Top             =   2070
               Width           =   7095
               Begin VB.CommandButton cmdPickCancel 
                  Cancel          =   -1  'True
                  Caption         =   "Cancel"
                  Height          =   375
                  Left            =   4080
                  TabIndex        =   80
                  Top             =   960
                  Width           =   2295
               End
               Begin VB.CommandButton cmdPickOK 
                  Caption         =   "Add Picklist Value"
                  Height          =   375
                  Left            =   960
                  TabIndex        =   79
                  Top             =   960
                  Width           =   2175
               End
               Begin VB.CheckBox Check1 
                  BackColor       =   &H00808080&
                  Caption         =   "Multichoice:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   28
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.TextBox Text5 
                  Height          =   285
                  Left            =   960
                  MaxLength       =   25
                  TabIndex        =   27
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   240
                  TabIndex        =   29
                  Top             =   300
                  Width           =   735
               End
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Height          =   1185
            Index           =   8
            Left            =   360
            TabIndex        =   17
            Top             =   5520
            Visible         =   0   'False
            Width           =   2070
            Begin VB.CommandButton Command13 
               Caption         =   "Add New Test"
               Height          =   375
               Left            =   1950
               TabIndex        =   19
               Top             =   6240
               Width           =   1935
            End
            Begin VB.CommandButton Command14 
               Caption         =   "Delete Test"
               Height          =   375
               Left            =   3990
               TabIndex        =   18
               Top             =   6240
               Width           =   1935
            End
            Begin Threed.SSFrame SSFrame1 
               Height          =   5415
               Left            =   480
               TabIndex        =   20
               Top             =   480
               Width           =   7215
               _ExtentX        =   12721
               _ExtentY        =   9546
               _Version        =   131073
               ForeColor       =   16777215
               BackColor       =   8421504
               Caption         =   "Profile Properties"
               Begin VB.CommandButton Command15 
                  Caption         =   "Delete This Profile"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   23
                  Top             =   5040
                  Width           =   2535
               End
               Begin VB.CommandButton Command12 
                  Caption         =   "Cancel"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   5640
                  TabIndex        =   22
                  Top             =   5040
                  Width           =   1335
               End
               Begin VB.CommandButton Command11 
                  Caption         =   "Save"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   4200
                  TabIndex        =   21
                  Top             =   5040
                  Width           =   1335
               End
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Height          =   1095
            Index           =   10
            Left            =   3240
            TabIndex        =   13
            Top             =   3870
            Visible         =   0   'False
            Width           =   1755
            Begin VB.TextBox txtHelp 
               BackColor       =   &H00C0FFFF&
               Height          =   2805
               Index           =   1
               Left            =   600
               MultiLine       =   -1  'True
               TabIndex        =   46
               Text            =   "frmmain.frx":CF8A
               Top             =   360
               Width           =   7095
            End
            Begin VB.TextBox txtHelp 
               BackColor       =   &H0080FFFF&
               Height          =   1365
               Index           =   0
               Left            =   780
               MultiLine       =   -1  'True
               TabIndex        =   14
               Text            =   "frmmain.frx":D19C
               Top             =   375
               Width           =   7095
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            Height          =   1185
            Index           =   4
            Left            =   3000
            TabIndex        =   12
            Top             =   1680
            Width           =   2235
            Begin RichTextLib.RichTextBox txtTestDisplay 
               Height          =   1935
               Left            =   480
               TabIndex        =   35
               Top             =   360
               Width           =   6255
               _ExtentX        =   11028
               _ExtentY        =   3408
               _Version        =   393217
               BackColor       =   12648447
               Enabled         =   -1  'True
               ScrollBars      =   2
               TextRTF         =   $"frmmain.frx":D380
            End
            Begin VB.TextBox txtTestDesc 
               BackColor       =   &H00C0FFFF&
               Height          =   405
               Index           =   0
               Left            =   1440
               MultiLine       =   -1  'True
               TabIndex        =   84
               Text            =   "frmmain.frx":D40B
               Top             =   2400
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txtTestDesc 
               BackColor       =   &H00C0FFFF&
               Height          =   405
               Index           =   2
               Left            =   4080
               MultiLine       =   -1  'True
               TabIndex        =   83
               Text            =   "frmmain.frx":D567
               Top             =   2400
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.TextBox txtTestDesc 
               BackColor       =   &H00C0FFFF&
               Height          =   405
               Index           =   1
               Left            =   2760
               MultiLine       =   -1  'True
               TabIndex        =   82
               Text            =   "frmmain.frx":D700
               Top             =   2400
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.TextBox txtTestDesc 
               BackColor       =   &H00C0FFFF&
               Height          =   1935
               Index           =   3
               Left            =   360
               MultiLine       =   -1  'True
               TabIndex        =   81
               Text            =   "frmmain.frx":D7A2
               Top             =   3840
               Visible         =   0   'False
               Width           =   6585
            End
            Begin VB.Label AddMode 
               BackColor       =   &H00808080&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   4740
               TabIndex        =   40
               Top             =   3300
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.Frame frapanel 
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   1035
            Index           =   11
            Left            =   7515
            TabIndex        =   10
            Top             =   3315
            Width           =   2115
            Begin VB.Label lblComment 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Please wait while the Logs for AHSL1 are loaded"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   945
               Left            =   360
               TabIndex        =   11
               Top             =   600
               Width           =   3195
            End
         End
      End
      Begin Listbar.SSListBar SSListBar1 
         Height          =   10212
         Left            =   24
         TabIndex        =   8
         Top             =   468
         Width           =   1524
         _ExtentX        =   2688
         _ExtentY        =   18013
         _Version        =   65536
         BorderStyle     =   0
         OLEDragMode     =   1
         OLEDropMode     =   2
         IconsLargeCount =   26
         Image(1).Index  =   1
         Image(1).Picture=   "frmmain.frx":DCB1
         Image(1).Key    =   "Tests"
         Image(2).Index  =   2
         Image(2).Picture=   "frmmain.frx":E58B
         Image(2).Key    =   "User"
         Image(3).Index  =   3
         Image(3).Picture=   "frmmain.frx":EE65
         Image(3).Key    =   "Group"
         Image(4).Index  =   4
         Image(4).Picture=   "frmmain.frx":F73F
         Image(4).Key    =   "Config"
         Image(5).Index  =   5
         Image(5).Picture=   "frmmain.frx":10019
         Image(5).Key    =   "Connections"
         Image(6).Index  =   6
         Image(6).Picture=   "frmmain.frx":10333
         Image(6).Key    =   "SPEC"
         Image(7).Index  =   7
         Image(7).Picture=   "frmmain.frx":10C0D
         Image(7).Key    =   "READ"
         Image(8).Index  =   8
         Image(8).Picture=   "frmmain.frx":10F27
         Image(8).Key    =   "HiLo"
         Image(9).Index  =   9
         Image(9).Picture=   "frmmain.frx":11241
         Image(9).Key    =   "UOM"
         Image(10).Index =   10
         Image(10).Picture=   "frmmain.frx":1155B
         Image(10).Key   =   "Specialty"
         Image(11).Index =   11
         Image(11).Picture=   "frmmain.frx":11E35
         Image(11).Key   =   "collect"
         Image(12).Index =   12
         Image(12).Picture=   "frmmain.frx":1214F
         Image(12).Key   =   "anatomy"
         Image(13).Index =   13
         Image(13).Picture=   "frmmain.frx":12469
         Image(13).Key   =   "Logs"
         Image(14).Index =   14
         Image(14).Picture=   "frmmain.frx":12D43
         Image(14).Key   =   "ICE"
         Image(15).Index =   15
         Image(15).Picture=   "frmmain.frx":1305D
         Image(15).Key   =   "Map"
         Image(16).Index =   16
         Image(16).Picture=   "frmmain.frx":13377
         Image(16).Key   =   "Monitor"
         Image(17).Index =   17
         Image(17).Picture=   "frmmain.frx":13691
         Image(18).Index =   18
         Image(18).Picture=   "frmmain.frx":13DE3
         Image(19).Index =   19
         Image(19).Picture=   "frmmain.frx":14235
         Image(20).Index =   20
         Image(20).Picture=   "frmmain.frx":1430E
         Image(21).Index =   21
         Image(21).Picture=   "frmmain.frx":14628
         Image(22).Index =   22
         Image(22).Picture=   "frmmain.frx":14F02
         Image(23).Index =   23
         Image(23).Picture=   "frmmain.frx":157DC
         Image(23).Key   =   "LOCATION"
         Image(24).Index =   24
         Image(24).Picture=   "frmmain.frx":15AF6
         Image(25).Index =   25
         Image(25).Picture=   "frmmain.frx":15E10
         Image(26).Index =   26
         Image(26).Picture=   "frmmain.frx":16A62
         Groups(1).CurrentGroup=   -1  'True
         Groups(1).Caption=   "Help"
         Groups(1).Key   =   "Help"
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   372
         Left            =   24
         TabIndex        =   5
         Top             =   24
         Width           =   16104
         _ExtentX        =   28406
         _ExtentY        =   656
         _Version        =   131073
         BackColor       =   8421504
         BevelOuter      =   1
         AutoSize        =   3
         Begin VB.ComboBox cboTrust 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Left            =   5475
            TabIndex        =   101
            Text            =   "cboTrust"
            Top             =   0
            Width           =   1455
         End
         Begin VB.ComboBox OrgList 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9885
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   -15
            Width           =   4035
         End
         Begin VB.Label OrgName 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Please select an organisation ---->"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   45
            TabIndex        =   7
            Top             =   15
            UseMnemonic     =   0   'False
            Width           =   5310
         End
         Begin VB.Label labOrgList 
            BackColor       =   &H8000000C&
            Caption         =   "Data Stream"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   8430
            TabIndex        =   102
            Top             =   30
            Width           =   1365
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   36
         Left            =   24
         TabIndex        =   4
         Top             =   10752
         Width           =   1512
         _ExtentX        =   2667
         _ExtentY        =   64
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   36
         Left            =   1608
         TabIndex        =   3
         Top             =   10752
         Width           =   14520
         _ExtentX        =   25612
         _ExtentY        =   64
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1764
               MinWidth        =   1764
               TextSave        =   "12/08/2011"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1764
               MinWidth        =   1764
               TextSave        =   "17:02"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   22013
            EndProperty
         EndProperty
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   10212
         Left            =   1620
         TabIndex        =   1
         Top             =   468
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   18013
         _Version        =   131073
         Caption         =   "ICE Config"
         AutoSize        =   3
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   10188
            Left            =   12
            TabIndex        =   2
            Top             =   12
            Width           =   4056
            _ExtentX        =   7154
            _ExtentY        =   17971
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   1
            OLEDragMode     =   1
            OLEDropMode     =   1
         End
      End
   End
   Begin VB.Menu item 
      Caption         =   "Item"
      Begin VB.Menu itemAdd 
         Caption         =   "Add..."
      End
      Begin VB.Menu itemAddAll 
         Caption         =   "Add All..."
         Visible         =   0   'False
      End
      Begin VB.Menu itemDelete 
         Caption         =   "Delete..."
      End
      Begin VB.Menu itemDeleteAll 
         Caption         =   "Delete All..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Requeue 
      Caption         =   "Requeue"
      Visible         =   0   'False
      Begin VB.Menu mnuRequeue 
         Caption         =   "Requeue"
      End
      Begin VB.Menu mnuResend 
         Caption         =   "Resend"
      End
      Begin VB.Menu mnuReprocess 
         Caption         =   "Reprocess"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DragNode As Object
Dim InDrag As Boolean
Public CfgPanelHelp As String
Public itemId As String
Public tvNode As String
Public newNode As MSComctlLib.Node
'Private sqlField As String
'Private sqlValue As String
Private maxLen As Integer
Private INI As String

Private LTS_Index As Long
Private LTS_OrgCode As String
Private LTS_DataStream As String

'Private tPList As PropertiesList
Dim Activity() As String
Dim ActivityDirs() As String
Dim ActivityDirPattern() As String
Private Type RECT
 Left   As Long
 Top    As Long
 Right  As Long
 Bottom As Long
End Type
Private nOrigin As String
Private errRet As Long
Private LogSearchOpt As Integer
Private LastLVItem As MSComctlLib.ListItem
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function DrawText& Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long)
Private Declare Function FillRect& Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

'Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim mfx As Single
Dim mfy As Single
Dim moNode As MSComctlLib.Node
Dim m_iScrollDir As Integer 'Which way to scroll
Dim mbFlag As Boolean
Dim CfgProgID As String
Dim CfgCfgID As String
Dim cfgWardId As String
Dim CfgUserID As String
Dim ShowRequesting As Boolean
Dim ShowUsers As Boolean
Dim hideEDI As Boolean
Dim hideAudit As Boolean
Dim hideConnections As Boolean
Dim DefaultItem As String

Public Property Get LogDisplayType() As Integer
   LogDisplayType = LogSearchOpt
End Property

Public Property Get CurrentLTSDataStream() As String
   CurrentLTSDataStream = LTS_DataStream
End Property

Public Property Let CurrentLTSIndex(lngNewValue As Long)
   LTS_Index = lngNewValue
End Property

Public Property Get CurrentLTSIndex() As Long
   CurrentLTSIndex = LTS_Index
End Property

Public Property Get CurrentLTSOrg() As String
   CurrentLTSOrg = LTS_OrgCode
End Property

Private Sub chkRestrict_Click()
   If chkRestrict.value = 0 Then
      dtPFrom.Enabled = False
      dtPTo.Enabled = False
   Else
      dtPFrom.Enabled = True
      dtPTo.Enabled = True
   End If
End Sub

Private Sub cmdGraph_Click()
   frmChart.Show 1
End Sub

Private Sub cmdPickCancel_Click()
   loadCtrl.TidyUp
End Sub

Private Sub cmdPickOK_Click()
   On Error GoTo procEH
   
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim thisNode As Node
   Dim i As Integer
   Dim vData As Variant
   Dim PickIndex As String
   Dim mStatus As ENUM_MenuStatus
   
   vData = objTV.ReadNodeData(objTV.ActiveNode)
'   pickIndex = objTView.NodeLevel(TreeView1.SelectedItem.Key)
'   Set thisNode = frmMain.TreeView1.SelectedItem
   If cmdPickOK.Caption = "Update" Then
      PickIndex = vData(0)
      If fraPickOpt.Tag = "Entries" Then
         mStatus = ms_DELETE
         strSQL = "UPDATE Request_Picklist_Data SET " & _
                  "Picklist_Value = '" & Text5.Text & "' " & _
                  "WHERE Picklist_Index = " & vData(0) & _
                     " AND Picklist_Value = '" & objTV.nodeKey(objTV.ActiveNode) & "'"
      Else
         mStatus = ms_BOTH
         strSQL = "UPDATE Request_Picklist SET " & _
                     "Picklist_Name = '" & Text5.Text & "', " & _
                     "Multichoice = " & Check1.value & _
                  " WHERE Picklist_Index = " & vData(0)
      End If
      iceCon.Execute strSQL
      PickIndex = vData(0)
   Else
      If fraPickOpt.Tag = "Entries" Then
         mStatus = ms_DELETE
         strSQL = "INSERT INTO Request_Picklist_Data " & _
                     "(Picklist_Index, Picklist_Value) VALUES (" & _
                     objTV.NodeLevel(objTV.ActiveNode) & ", '" & _
                     Text5.Text & "')"
      Else
         mStatus = ms_BOTH
         strSQL = "INSERT INTO Request_Picklist " & _
                     "(Picklist_Name, Multichoice)" & _
                  "VALUES ('" & _
                     Text5.Text & "', " & _
                     Check1.value & ")"
      End If
      iceCon.Execute strSQL
      strSQL = "SELECT Max(Picklist_Index) FROM Request_Picklist"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      PickIndex = RS(0)
      RS.Close
   End If
   
   cmdPickOK.Caption = "Update"
   Set RS = Nothing
   loadCtrl.Refresh PickIndex, mStatus
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.cmdPickOk_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub cmdSrchCancel_Click()
  fView.Show Fra_HELP, ""
End Sub

Private Sub CmdSrchOk_Click()

   On Local Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim defLang As String
   Dim param1 As String
   Dim GPCode As String
   Dim fromDate As Date
   Dim toDate As Date
   Dim strTemp As String
   
'  Read the default date fromat for the database language, store, then set dateformat to dmy
'  This ensures the date restriction constructs work as intended
'   strSQL = "SELECT dateformat " & _
'            "FROM master.dbo.syslanguages " & _
'            "WHERE langid = @@default_langid"
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'   defLang = RS!dateformat
'   RS.Close
'   ICECon.Execute "SET DateFormat dmy"
         
   If chkRestrict.value = 1 Then
      fromDate = dtPFrom.value
      toDate = dtPTo.value
   Else
      fromDate = 0
      toDate = 0
   End If
   
   frmMain.MousePointer = vbHourglass
   loadCtrl.SearchStartDate = fromDate
   loadCtrl.SearchEndDate = toDate
   loadCtrl.practice = ""
   
   If optLogSearch(0).value = True Then
'     Search by date
      loadCtrl.FirstView
'      loadCtrl.Dates objTV.ActiveNode, dtPFrom.value, dtPTo.value
      
   ElseIf optLogSearch(1).value Then
      strTemp = Left(ComboSrchPractice.Text, InStr(ComboSrchPractice.Text, " -") - 1)
'     Search by practice
      strSQL = "SELECT DISTINCT EDI_Local_Key1 " & _
               "FROM EDI_Recipient_Individuals ei " & _
                  "INNER JOIN EDI_Matching em " & _
                  "ON ei.Individual_Index = em.Individual_Index " & _
               "WHERE EDI_Org_NatCode = '" & strTemp & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      GPCode = RS("EDI_Local_Key1")
      RS.Close
      Set RS = Nothing
      loadCtrl.practice = GPCode
      loadCtrl.FirstView
'      loadCtrl.Dates OrgList.Text, fromDate, toDate, gpCode
   
   ElseIf optLogSearch(2).value Then
'     Search by Patient name
      loadCtrl.PatientSearch txtSrchSurname.Text, txtSrchForename.Text ', fromDate, toDate
   
   ElseIf optLogSearch(3).value Then
'     Search by Hospital/NHS Number
      loadCtrl.PatientSearch txtSrchNHS.Text ', , fromDate, toDate
   
   ElseIf optLogSearch(4).value Then
'     Search by report id
      loadCtrl.LabSearch txtLogSearchLab.Text ', fromDate, toDate
   End If
   
   TreeView1.Visible = True
'   ICECon.Execute "SET dateformat " & defLang
   frmMain.LogText.Visible = frmMain.TreeView1.Visible
   fView.Show Fra_LOGVIEW, ""
   frmMain.MousePointer = vbNormal
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   Else
      HandleError "LoadEDIRecipents.SpecificIN"
   End If
End Sub

Private Sub cmdStats_Click()
   loadCtrl.Statistics objTV.ActiveNode
End Sub

Private Sub Command12_Click()
   Command11.Enabled = False
   Command12.Enabled = False
   Command13.Enabled = True
   Command14.Enabled = True
   SSListBar1.Enabled = True
   TreeView1.Enabled = True
   OrgList.Enabled = True
End Sub

Public Sub CheckVersion()
   Dim dbVer As Integer
   Dim iniVer As Integer
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim iceCmd As New ADODB.Command
   'Dim iniFile As String
   
   strSQL = "SELECT Version FROM dbVersion"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      RS.MoveLast
      dbVer = RS!Version
   Else
      Err.Raise 3041, "IceConfig.ModIceConfig.GetConnection", "No version control held on database"
   End If
   RS.Close
   Set RS = Nothing
   
   hideEDI = (Read_Ini_Var("General", "hideEDI", iniFile) = 1)
   hideAudit = (Read_Ini_Var("General", "hideAudit", iniFile) = 1)
   hideConnections = (Read_Ini_Var("General", "hideConnections", iniFile) = 1)
   
   DefaultItem = Read_Ini_Var("General", "DefaultItem", iniFile)
   If DefaultItem = "" Then
      DefaultItem = "EDI Recipients"
   End If
End Sub

Private Sub Command16_Click()
    'Delete Test
    Load frmWait
    frmWait.Label1.Caption = "Please wait whilst test dependencies are calculated..."
    frmWait.Show
    frmWait.Refresh
    
End Sub

Private Sub Command19_Click()
    'Cancel Changes to Configuration Option
    Frame2.Visible = False
End Sub

Private Sub DTStats_Change()
   On Error GoTo procEH
   If DTStats.Tag = "" Then
      DTStatEnd.MaxDate = DTStats.value
      If optStatPeriod(0) Then
         DTStatEnd.value = DTStats.value
      ElseIf optStatPeriod(1).value Then
         DTStatEnd.value = DateAdd("d", -7, DTStats.value)
      ElseIf optStatPeriod(2).value Then
         DTStatEnd.value = DateAdd("d", -28, DTStats.value)
      Else
         
      End If
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmMain.DTStats_Change"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub DTStats_CloseUp()
   DTStatEnd.MaxDate = DTStats.value
   If optStatPeriod(0) Then
      DTStatEnd.value = DTStats.value
   ElseIf optStatPeriod(1).value Then
      DTStatEnd.value = DateAdd("d", -7, DTStats.value)
   ElseIf optStatPeriod(2).value Then
      DTStatEnd.value = DateAdd("d", -28, DTStats.value)
   Else
      DTStatEnd.value = DTStats.value + DTStats.Tag
   End If
End Sub

Private Sub DTStats_DropDown()
   DTStats.Tag = DateDiff("d", DTStats.value, DTStatEnd.value)
End Sub

Private Sub ediPr_AfterEdit(PropertyItem As PropertiesListCtl.PropertyItem, newValue As Variant, Cancel As Boolean)
   On Error GoTo procEH
   Dim blnUnlimited As Boolean
   Dim strDrive As String
   Dim strPath As String
   Dim strUNC As String
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim intType As Integer
   
   newValue = Replace(newValue, "'", "`")
   If PropertyItem.Style = plpsString Then
      PropertyItem = Replace(PropertyItem, "'", "`")
   End If
   
   With PropertyItem
      If .Style = plpsFolder Then
         If filepathToUNC Then
            strDrive = fs.GetDriveName(newValue)
            strPath = Mid(newValue, Len(strDrive) + 1)
            If GetUNCPath(strDrive, strUNC) = 0 Then
               newValue = fs.BuildPath(strUNC, strPath)
            Else
               strUNC = GetMachineName
               If InStr(1, UCase(strDrive), UCase(strUNC)) = 0 Then
                  newValue = "\\" & strUNC & Mid(newValue, 3)
               End If
            End If
         End If
      End If
      
      If .Style = plpsNumber Then
         If newValue = "" Then
            newValue = 0
         End If
         
         blnUnlimited = (.Min = 0 And .max = 0)
         
         If (newValue > .max Or newValue < .Min) And blnUnlimited = False Then
            MsgBox "Either the maximum or minimum value for " & .Caption & " has been breached. " & _
                   vbCrLf & "Please select a value between " & .Min & " and " & .max, vbInformation, _
                   .Caption & " Validation"
            Cancel = True
         Else
            Select Case .Key
               Case "NUMMAX"
                  With EdiPr("NUMMIN")
                     If .value >= Val(newValue) Then
                        newValue = .value
                     End If
                  End With
               
               Case "NUMMIN"
                  With EdiPr("NUMMAX")
                     If .value <= Val(newValue) Then
                        newValue = .value
                     End If
                  End With
               
               Case "MAX_AGE"
                  With EdiPr("MIN_AGE")
                     If .value >= Val(newValue) Then
                        newValue = .value
                        MsgBox "The maximum age has been specified as less than the minimum age. Maximum age reset to " & newValue, _
                               vbExclamation, "Invalid age"
                     End If
                  End With
'                  If newValue > 120 Then
'                     MsgBox "Maximum age specified as " & newValue & ". Age reset to: 120", vbExclamation, "Invalid Age"
'                     newValue = 120
'                  End If
                  
               Case "MIN_AGE"
                  With EdiPr("MAX_AGE")
                     If .value <= Val(newValue) Then
                        .value = newValue
                        MsgBox "The minimum age has been specified as greater than the maximum age. Maximum age reset to " & newValue, _
                               vbExclamation, "Invalid age"
                     End If
                  End With
'                  If newValue > 100 Then
'                     MsgBox "Maximum age specified as " & newValue & ". Age reset to: 100", vbExclamation, "Invalid Age"
'                     newValue = 100
'                  End If
               
            End Select
         End If
      End If
      
      If Left(.Key, 5) = "SP+MS" Then
         intType = Val(Right(.Key, 1))
         If intType > 4 And intType < 9 Then
            Cancel = False
            If newValue <> "" Then
               If IsDate(newValue) Then
                  newValue = Format(newValue, "HH:NN")
               Else
                  MsgBox "Please enter a valid time in the format 'HH:MM'", vbExclamation, "Invalid Time"
                  Cancel = True
               End If
            End If
         End If
      End If
      
'      If .Key = "ENABLED" Then
'         If ediPr("SCREEN_COLOUR").DialogTitle = "0-None Set" Or _
'            ediPr("PROV_ID").value = 0 Then
'            MsgBox "A caption colour and a Test Provider must be specified before " & _
'                   "this test can be enabled", vbExclamation, "Unable to enable test"
'            newValue = False
'         End If
'      End If
'
'      If .Key = "CFGSTYLE" Then
'         With ediPr("CFGDEF")
'            .Style = newValue
'            Select Case newValue
'               Case 7
'                  .value = CDate(Now())
'
'               Case 8
'                  .value = ""
'
'               Case 11
'                  .value = False
'
'               Case Else
'                  .value = 0
'
'            End Select
'         End With
'      End If
'
'      If .Key = "OVERSTYLE" Then
'         With ediPr("OVERVALUE")
'            .Style = newValue
'            Select Case newValue
'               Case 7
'                  .value = Format(Now(), "dd/mm/yyyy")
'
'               Case 8
'                  .value = ""
'
'               Case 11
'                  .value = False
'
'               Case Else
'                  .value = 0
'
'            End Select
'         End With
'      End If
'
'      If .Key = "STYLE" Then
'         loadCtrl.visibility CStr(newValue)
'      End If
'
'      If .Key = "TYPE" Then
'         loadCtrl.TypeEntry CStr(newValue)
'      End If
'
'      If .Key = "TEST_CODE" Then
'         If newValue = "" Then
'            ediPr("TESTID").value = ""
'         Else
'            ediPr("TESTID").value = cboTrust.Text & " " & Left(newValue, 20)
'         End If
'      End If
'
'      If .Key = "DCODE" Then
'         Cancel = loadCtrl.Validate(CStr(newValue))
'      End If
'
'      If .Key = "SCREEN_PANEL" Then
'         Select Case newValue
'            Case -1
'
'            Case 0
'               ediPr("PANEL_PAGE").ListItems.Clear
'               ediPr("PANEL_PAGE").value = ""
'               ediPr("SCREEN_POSN").ListItems.Clear
'               ediPr("SCREEN_POSN").value = ""
'
'            Case Else
'               If .value <> newValue Then
'                  With ediPr("PANEL_PAGE")
'                     .value = ""
'                     loadCtrl.SetUpPages CStr(newValue)
'                     .value = ""
'   '                  loadCtrl.GetVacantScreenPositions CStr(NewValue), .value, 0
'                  End With
'                  ediPr("SCREEN_POSN").value = ""
'               End If
'         End Select
'      End If
'
'      If .Key = "PANEL_PAGE" Then
'         If newValue = "<No Page>" Then
'            ediPr("SCREEN_POSN").ListItems.Clear
'            ediPr("SCREEN_POSN").value = ""
'            newValue = ""
'         ElseIf .value <> newValue Then
'            If loadCtrl.GetVacantScreenPositions(CStr(ediPr("SCREEN_PANEL").value), CStr(newValue), .value) = False Then
'               ediPr("SCREEN_POSN").ListItems.Clear
'               ediPr("SCREEN_POSN").value = ""
'               newValue = ""
'               MsgBox "No vacant screen positions available for this page", _
'                      vbInformation, "Unable to select requested page"
'            End If
'         End If
'      End If
'
'      If .Key = "SCREEN_POSN" Then
'         If newValue = "" Then
'            Cancel = True
'         End If
'      End If
            
      If .Key = "LOCAL" Then
         EdiPr("LOCCODE").value = cboTrust.Text & " " & newValue
      End If
      
      If .Key = "NAT" Then
         loadCtrl.PB_AfterEdit PropertyItem, newValue
      End If
      
   End With
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.FurtherInfo = PropertyItem.Key & ": " & newValue
   eClass.CurrentProcedure = "IceConfig.frmMain.edipr_AfterEdit"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

'Private Sub ediPr_BeforeEdit(PropertyItem As PropertiesListCtl.PropertyItem, Cancel As Boolean)
'   Dim RS As New ADODB.Recordset
'
'   If PropertyItem.Key = "COL_CODE" Then
'      Exit Sub
'   End If
'
'   If PropertyItem.Style = plpsColor Then
'      If PropertyItem.Key <> "TUBE_CODE" And PropertyItem.Key <> "PAED_TUBE_CODE" Then
'         If PropertyItem.DialogTitle <> "" Then
'            frmColour.CurrentColour = Left(PropertyItem.DialogTitle, InStr(1, PropertyItem.DialogTitle, "-") - 1)
'         End If
'         frmColour.Show 1
'         If PickedColIndex > 0 Then
'            PropertyItem.value = Format(PickedCol)
'            PropertyItem.DialogTitle = Format(PickedColIndex) & "-" & PickedColName
'         End If
'         Cancel = True
'      Else
''         frmTube.CurrentTube = Left(PropertyItem.DialogTitle, InStr(1, PropertyItem.DialogTitle, "-") - 1)
'         frmTube.Show 1
'         If PickedTubeIndex > -1 Then
'            PropertyItem.value = PickedTubeCol
'            PropertyItem.DialogTitle = Format(PickedTubeIndex) + "-" + PickedTube
'         End If
'         Cancel = True
'      End If
'   End If
'
'   If PropertyItem.Key = "STYLE" And (PropertyItem.value <> 0 And PropertyItem.value <> "") Then
'      Cancel = (MsgBox("Changing the style will cause changed property details to be lost. " & _
'                "Are you sure you wish to change the behaviour style?", vbQuestion + vbYesNo, "Change Behaviour Style") = vbNo)
'   End If
'
'   Set RS = Nothing
'End Sub
'
'Private Sub ediPr_BeforeExtendedEdit(PropertyItem As PropertiesListCtl.PropertyItem, Cancel As Boolean)
'   On Error GoTo procEH
'   Select Case PropertyItem.Key
'      Case "SCREEN_PANEL"
'         loadCtrl.NewPanel PropertyItem
'
'      Case "PANEL_PAGE"
'         loadCtrl.NewPanelPage PropertyItem
'
'   End Select
'   Exit Sub
'
'procEH:
'   If eClass.Behaviour = -1 Then
'      Stop
'      Resume
'   End If
'   eClass.CurrentProcedure = "IceConfig.frmmain.edipr_BeforeExtendedEdit"
'   eClass.Add Err.Number, Err.Description, Err.Source, False
'End Sub

Private Sub edipr_PropertyBrowseClick(PropertyItem As PropertiesListCtl.PropertyItem)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim strArray() As String
   Dim strTemp As String
   
   EdiPr.Tag = PropertyItem.Key
   
   Select Case PropertyItem.Key
      Case "CSPEC"
         loadCtrl.PB_Click PropertyItem
         
      Case "DMSPEC"
         loadCtrl.PB_Click PropertyItem
         
      Case "EC"
         frmReadCodes.CurrentReadCode = PropertyItem.value
         frmReadCodes.Show 1
         
      Case "LSA"
         If EdiPr("SA").value = "" Then
            MsgBox "Please set an Anatomical origin for this sample before adding descriptions", vbInformation, "Anatomical Origin not defined"
            EdiPr("SA").Selected = True
         Else
            frmSampleData.Tag = "EDI_Local_Sample_AnatOrigin"
            frmSampleData.txtNatCode.Text = EdiPr("SA").value
            frmSampleData.fraDesc.Caption = "Anatomical Origin Description"
            frmSampleData.Show 1
         End If
      
      Case "LSC"
         If EdiPr("SC").value = "" Then
            MsgBox "Please set a Collection Code for this sample before adding descriptions", vbInformation, "Collection Code not defined"
            EdiPr("SC").Selected = True
         Else
            frmSampleData.Tag = "EDI_Local_Sample_CollectionTypes"
            frmSampleData.txtNatCode.Text = EdiPr("SC").value
            frmSampleData.fraDesc.Caption = "Collection Type Description"
            frmSampleData.Show 1
         End If
      
      Case "LST"
         If EdiPr("ST").value = "" Then
            MsgBox "Please set a Sample Code for this sample before adding descriptions", vbInformation, "Sample Code not defined"
            EdiPr("ST").Selected = True
         Else
            frmSampleData.Tag = "EDI_Local_Sample_Types"
            frmSampleData.txtNatCode.Text = EdiPr("ST").value
            frmSampleData.fraDesc.Caption = "Sample Description"
            frmSampleData.Show 1
         End If
         
      Case "NatUOM"
         frmUOMCodes.Show 1
            
      Case "ST"
         frmSampCodes.DbTable = "CRIR_Sample_Type"
         frmSampCodes.NationalDescriptionField = "Sample_Text"
         frmSampCodes.ReturnDataTo = "ST"
         frmSampCodes.Show 1
'         edipr("SD").value = frmSampCodes.txtSpec.Text
         Unload frmSampCodes
'         ediPr_AfterEdit edipr("NS"), edipr("NS").value, False
         
      Case "RC"
         frmReadCodes.CurrentReadCode = PropertyItem.value
         frmReadCodes.Show 1
         
      Case "READ_CODE"
         frmReadCodes.CurrentReadCode = PropertyItem.value
         frmReadCodes.Show 1
         
      Case "ANATCODE"
         frmSampCodes.DbTable = "CRIR_Sample_AnatOrigin"
         frmSampCodes.NationalCodeField = "Origin_Code"
         frmSampCodes.NationalDescriptionField = "Origin_Text"
         frmSampCodes.ReturnDataTo = "ANATCODE"
         frmSampCodes.InitialValue = Trim(PropertyItem.value)
         frmSampCodes.Show 1
         Unload frmSampCodes
      
      Case "COLLCODE"
         frmSampCodes.DbTable = "CRIR_Sample_CollectionType"
         frmSampCodes.NationalCodeField = "Collection_Code"
         frmSampCodes.NationalDescriptionField = "Collection_Text"
         frmSampCodes.ReturnDataTo = "COLLCODE"
         frmSampCodes.InitialValue = Trim(PropertyItem.value)
         frmSampCodes.Show 1
         Unload frmSampCodes
         
      Case "SCODE"
         frmSampCodes.InitialValue = Trim(PropertyItem.value)
         frmSampCodes.DbTable = "CRIR_Sample_Type"
         frmSampCodes.NationalCodeField = "Sample_Code"
         frmSampCodes.NationalDescriptionField = "Sample_Text"
         frmSampCodes.ReturnDataTo = "SCODE"
         frmSampCodes.Show 1
         Unload frmSampCodes
      
      Case "IN+IN1"
         frmGP.PracticeId = EdiPr("NATIONAL").value
         frmGP.GPNatCode = EdiPr("IN+IN2").value
         frmGP.txtGPName = EdiPr("IN+IN1").value
         frmGP.txtGPNatCode = EdiPr("IN+IN2").value
         frmGP.txtPrName = EdiPr("NA").value
         frmGP.txtPrNatCode = EdiPr("NATIONAL").value
         frmGP.txtPrAdr = EdiPr("AD").value
         frmGP.Form_Init
         frmGP.Show 1
      
      Case "IN+IN3"
         frmGPMatching.GPDetails = EdiPr("SUBID").value & "|" & EdiPr("IN+IN1").value & "|" & EdiPr("IN+IN2").value & "|" & EdiPr("NATIONAL").value
'         frmGPMatching.NationalCode = edipr("IN+IN2").value
'         frmGPMatching.IndividualId = edipr("SUBID").value
         frmGPMatching.Show 1
      
'      Case "PROV_ID"
'         frmProvider.Show 1
'         If PickedProvIndex <> 0 Then
'             PropertyItem.value = PickedProvIndex = ""
'             PropertyItem.DialogTitle = PickedProv
'         End If

      Case "TD"   '  Trader details
'         strSQL = "SELECT EDI_NatCode " & _
'                  "FROM EDI_Recipients " & _
'                  "WHERE EDI_NatCode = '" & ediPr("NATIONAL").value & "'"
'         Set RS = iceCon.Execute(strSQL)
'
'         If RS.EOF Then
'            MsgBox "Please create and save the EDI Details before setting the Trader code", vbInformation, "Warning"
'         Else
         'If Val(ediPr("REFID")) = 0 Then
            'MsgBox "Save the recipient before setting the Trader Code", vbInformation, "Save required"
         'Else
            frmTraderDets.EDI_NatCode = IIf(EdiPr("NATIONAL").value = "", EdiPr("NC").value, EdiPr("NATIONAL").value)
            frmTraderDets.EDI_RefIndex = Val(EdiPr("REFID").value)
            
            frmTraderDets.Show 1
            EdiPr("REFID").value = frmTraderDets.EDI_RefIndex
            frmTraderDets.Hide
            If frmTraderDets.recipientStatus > 0 Then
               SSListBar1_ListItemClick SSListBar1.ListItems("EDI Recipients")
            End If
         'End If
         
'         RS.Close
         'SSListBar1_ListItemClick SSListBar1.ListItems("EDI Recipients")
               
      Case "NATCODE"
         Select Case EdiPr.Caption 'Pages(edipr.ActivePage).Caption
            Case "Samp"
               frmSampCodes.InitialValue = Trim(PropertyItem.value)
               frmSampCodes.DbTable = "CRIR_Sample_Type"
               frmSampCodes.NationalCodeField = "Sample_Code"
               frmSampCodes.NationalDescriptionField = "Sample_Text"
               frmSampCodes.ReturnDataTo = "NATCODE"
               frmSampCodes.Show 1
               Unload frmSampCodes
            
            Case "Anat"
               frmSampCodes.DbTable = "CRIR_Sample_AnatOrigin"
               frmSampCodes.NationalCodeField = "Origin_Code"
               frmSampCodes.NationalDescriptionField = "Origin_Text"
               frmSampCodes.ReturnDataTo = "NATCODE"
               frmSampCodes.InitialValue = Trim(PropertyItem.value)
               frmSampCodes.Show 1
               Unload frmSampCodes
            
            Case "Coll"
               frmSampCodes.DbTable = "CRIR_Sample_CollectionType"
               frmSampCodes.NationalCodeField = "Collection_Code"
               frmSampCodes.NationalDescriptionField = "Collection_Text"
               frmSampCodes.ReturnDataTo = "NATCODE"
               frmSampCodes.InitialValue = Trim(PropertyItem.value)
               frmSampCodes.Show 1
               Unload frmSampCodes
               
            Case "EDI Recipients"
               frmSpecs.Show
               
            Case Else
               frmSpecs.Show 1
               EdiPr("DCODE").value = Left(EdiPr("NATCODE").value, 3)
            
         End Select
         
      Case Else
         frmSpecs.CalledFrom = PropertyItem.Key
         frmSpecs.Show 1
         If EdiPr.Tag = "SP+MS1" Then
            ediPr_AfterEdit EdiPr("SP+MS2"), EdiPr("SP+MS2").value, False
         End If
         Unload frmSpecs
   
   End Select
   ediPr_AfterEdit PropertyItem, EdiPr(PropertyItem.Key).value, False
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.frmmain.edipr_PropertyBrowseClick"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
'
'Private Sub ediPr_RequestDisplayValue(PropertyItem As PropertiesListCtl.PropertyItem, DisplayValue As String)
'   Dim sPos As Integer
'   Dim ePos As Integer
'   Dim pLen As Integer
'   Dim dVal As String
'
'   If PropertyItem.DialogTitle = "" Then
'      Exit Sub
'   End If
'
'   sPos = InStr(PropertyItem.DialogTitle, "-") + 1
'   ePos = InStr(PropertyItem.DialogTitle, "-") - 1
'   pLen = (Len(PropertyItem.DialogTitle) - ePos) + 1
'
'   Select Case PropertyItem.Key
'      Case "SCREEN_COLOUR"
'         DisplayValue = Mid$(PropertyItem.DialogTitle, sPos, pLen) & " (" + Left$(PropertyItem.DialogTitle, ePos) + ")"
'
'      Case "HELP_COLOUR"
'         DisplayValue = Mid$(PropertyItem.DialogTitle, sPos, pLen) & " (" + Left$(PropertyItem.DialogTitle, ePos) + ")"
'
'      Case "TUBE_CODE"
'         DisplayValue = Mid$(PropertyItem.DialogTitle, sPos, pLen) & " (" + Left$(PropertyItem.DialogTitle, ePos) + ")"
'
'      Case "PAED_TUBE_CODE"
'         DisplayValue = Mid$(PropertyItem.DialogTitle, sPos, pLen) & " (" + Left$(PropertyItem.DialogTitle, ePos) + ")"
'
'      Case "PROV_ID"
'         DisplayValue = Trim(PropertyItem.DialogTitle & "") & " (" & PropertyItem.value & ")"
'
'      Case "PROFILE_COLOUR"
'         DisplayValue = Mid$(PropertyItem.DialogTitle, sPos, pLen) & " (" + Left$(PropertyItem.DialogTitle, ePos) + ")"
'
'      Case "PANEL_PAGE"
'         If PropertyItem.value = "" Then
'            DisplayValue = "<No Page>"
'         Else
'            If Not (PropertyItem.ListItems.SelectedItem Is Nothing) Then
'               DisplayValue = PropertyItem.ListItems.SelectedItem.value
'            End If
'         End If
'
''      Case "TEST_CODE"
''         DisplayValue = Mid(PropertyItem.value, 7)
''
'      Case "TEST_TYPE"
'         If PropertyItem.value = 2 Then
'            DisplayValue = "Histology"
'         ElseIf PropertyItem.value = 1 Then
'            DisplayValue = "Bloodbank"
'         Else
'            DisplayValue = "Standard"
'         End If
'
'      Case Else
'         If Not (PropertyItem.ListItems.SelectedItem Is Nothing) Then
'            DisplayValue = PropertyItem.ListItems.SelectedItem.value
'         End If
'
'   End Select
'End Sub

Public Function FileLocation(fPath As String, _
                             nodeSource As String, _
                             errorFile As Boolean) As String
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim iVar As String
   Dim newPath As String
   Dim cName As String
   Dim eDir As String
   
   cName = ""
   If fs.FolderExists(fPath) Then
      newPath = fPath
   Else
      strSQL = "SELECT * " & _
               "FROM Service_Types " & _
               "WHERE Type_Index = " & nodeSource
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.BOF = False And RS.EOF = False Then
         Select Case UCase(Trim(RS!Description))
            Case "LABORATORY REPORTS"
               cName = "IMPORTER"
               
            Case "EDI REPORTS"
               cName = "EXPORTER"
               
         End Select
         If errorFile Then
'            eDir = fs.GetBaseName(fPath)
            iVar = Left(cName, 3) & "Err"
         Else
'            eDir = ""
            iVar = Left(cName, 3) & "Hist"
         End If
      End If
      RS.Close
      
      strSQL = "SELECT * " & _
               "FROM Connections " & _
               "WHERE Connection_Name = '" & cName & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If errorFile Then
         eDir = Mid(fPath, Len(Trim(RS!Connection_ErrorDirs & "")) + 2)
      Else
         eDir = Mid(fPath, Len(Trim(RS!Connection_HistoryDirs & "")) + 2)
      End If
      RS.Close
      
      newPath = Read_Ini_Var("Directory_Overrides", iVar, iniFile)
      newPath = fs.BuildPath(newPath, eDir)
   End If
   
   FileLocation = newPath
   
   Set RS = Nothing
End Function

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   Dim blnAddSpec As Boolean
   
   If Source.Name = "TreeView1" Then
      Timer1.Enabled = False
   End If
   blnAddSpec = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        MsgBox "The character ' is not permitted, please use the ` character instead"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
   On Error GoTo procEH
   Dim i As Integer
   Dim newTop As Integer
   Dim CG As Integer
   Dim CI As Integer
   Dim iceCmd  As New ADODB.Command
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim rc As RECT
   Dim msg As String
    
   wb.NavigateTo fs.BuildPath(App.Path, "logo.html"), True
   
   frmMain.Height = 768 * Screen.TwipsPerPixelY
   frmMain.Width = 1024 * Screen.TwipsPerPixelX
   
   frmMain.cboTrust.Locked = True
   SSSplitter1.Panes(0).LockHeight = True
   SSSplitter1.Panes(1).LockWidth = False
   SSSplitter1.Panes(4).LockHeight = True
   SSSplitter1.Panes(5).LockHeight = True
   SSSplitter1.Panes(4).Width = SSSplitter1.Panes(1).Width
   SSSplitter1.Panes(4).LockWidth = False
   SSSplitter1.Panes(2).MinWidth = 100
    
'   SSListBar1.Enabled = False
   Timer2.Enabled = False
   LogBackColour = LogText.BackColor
   StatusBar1.Panels(3).Text = "Copyright © Sunquest Information Systems, 2010.  All Rights Reserved."
   
   frmMain.Caption = formHeader
   
   wb.LocationTitle = "Sunquest Systems"
   CheckVersion
   
   newTop = frmSplash.Top - (frmLogon.Height / 2) - 200
   
   frmLogon.Left = (Screen.Width - frmLogon.Width) / 2
   frmLogon.Top = newTop + (frmSplash.Height / 2)
   frmLogon.Show 1
   
   Me.Caption = Me.Caption + " (" & DB_Name & " Database) - Version " & Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
'    CfgPanelHelp = Label6.Caption
   Timer1.Enabled = False
   Timer1.Interval = 200
'   blnAddStatus = False
   fView.Hide
   
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rc, 0&)
   
   frmMain.Top = rc.Top
   frmMain.Left = rc.Left
   frmMain.Height = Screen.TwipsPerPixelX * rc.Bottom
   frmMain.Width = Screen.TwipsPerPixelX * rc.Right
   frmMain.WindowState = 2
   
   Unload frmSplash
   
   With EdiPr
      .Top = 350
      .Left = 250
      .Height = 4650
      .Width = 6900
   End With
   
   dtPFrom.value = DateAdd("m", -1, Now())
   dtPTo.value = Now()
   
   wb.Top = 400
   wb.Left = 400

'   If ShowRequesting Then
'      SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Configuration Items", "Configuration Items"
'      With SSListBar1.Groups(SSListBar1.Groups.Count)
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Request Details", "Request Details"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Panels", "Panels and Pages"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Term", "Terminology"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Profiles", "Profiles"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Picklists", "Picklists"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Rules", "Rules"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Clinicians", "Clinicians"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 2
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Colours", "Colours"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Location", "Location"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 23
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Request Sample Panels", "Request Sample Panels"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Request Sample Panel Options", "Request Sample Panel Options"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Blood/Histo/Micro", "Blood/Histo/Micro"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "System Configuration", "System Configuration"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
'      End With
'   End If
   
'   If ShowUsers Then
'      SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "User Management", "User Management"
'      With SSListBar1.Groups(SSListBar1.Groups.Count)
'         .ListItems.Add .ListItems.Count + 1, "Users", "Users"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 2
'         .ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Permissions", "Permissions"
'         .ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 3
'      End With
'   End If
   
   If Not hideEDI Then
      SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "EDI Management", "EDI Management"
      With SSListBar1.Groups(SSListBar1.Groups.Count)
         .ListItems.Add .ListItems.Count + 1, "My Settings", "My Settings"
         .ListItems(.ListItems.Count).IconLarge = 4
         .ListItems.Add .ListItems.Count + 1, "EDI Recipients", "EDI Recipients"
         .ListItems(.ListItems.Count).IconLarge = 3
         .ListItems.Add .ListItems.Count + 1, "EDI Clinicians", "EDI Clinicians"
         .ListItems(.ListItems.Count).IconLarge = 3
         .ListItems.Add .ListItems.Count + 1, "Result Mapping", "Result Mapping"
         .ListItems(.ListItems.Count).IconLarge = 1
         .ListItems.Add .ListItems.Count + 1, "UOM Mapping", "UOM Mapping"
         .ListItems(.ListItems.Count).IconLarge = 9
         .ListItems.Add .ListItems.Count + 1, "Specimen Types", "Specimen Types"
         .ListItems(.ListItems.Count).IconLarge = 6
         .ListItems.Add .ListItems.Count + 1, "Körner Medical Specialties", "Körner Medical Specialties"
         .ListItems(.ListItems.Count).IconLarge = 7
         .ListItems.Add .ListItems.Count + 1, "RepList Entries", "Current Rep List Entries"
         .ListItems(.ListItems.Count).IconLarge = 20
      End With
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Specimen Anatomical Origins", "Specimen Anatomical Origins"
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 12
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Specimen Collection Procedures", "Specimen Collection Procedures"
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 11
   End If
   
   If Not hideAudit Then
      SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Audit Logs", "Audit Logs"
      With SSListBar1.Groups(SSListBar1.Groups.Count)
         .ListItems.Add .ListItems.Count + 1, "Logs", "Logs"
         .ListItems(.ListItems.Count).IconLarge = 13
         .ListItems.Add .ListItems.Count + 1, "Search", "Search"
         .ListItems(.ListItems.Count).IconLarge = 21
         .ListItems.Add .ListItems.Count + 1, "Filter", "Filter"
         .ListItems(.ListItems.Count).IconLarge = 25
         .ListItems.Add .ListItems.Count + 1, "Requeue", "Requeue"
         .ListItems(.ListItems.Count).IconLarge = 24
         .ListItems.Add .ListItems.Count + 1, "System Activity", "System Activity"
         .ListItems(.ListItems.Count).IconLarge = 16
      End With
   End If
   
   If Not hideConnections Then
      SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Connections", "Connections"
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Connections", "Connections"
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 5
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Maps", "Maps"
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 15
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Modules", "Modules"
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 14
   End If
   
   If SSListBar1.Groups.Count > 1 Then
      SSListBar1.Groups.Remove SSListBar1.Groups(1)
      SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Help", "Help"
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "About", "About"
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 17
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Helpdesk", "Helpdesk"
      SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 22
      
      If Dir(App.Path + "\ICEKeyGen.EXE") <> "" Then
         SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Key Generation", "Key Generation"
         SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 18
      End If
   End If
   
   If GetOrganisations = False Then
      DefaultItem = "My Settings"
   End If
   
   Select Case DefaultItem
      Case "My Settings"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 1
      Case "EDI Recipients"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 2
      Case "READ Codes"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 3
      Case "Result Mapping"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 4
      Case "UOM Mapping"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 5
      Case "Specimen Type"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 6
      Case "Körner Medical Specialties"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 7
      Case "Specimen Anatomical Origins"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 8
      Case "Specimen Collection Procedures"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "EDI Management" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 4
         CI = 9
      Case "Logs"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "Audit Logs" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 5
         CI = 1
      Case "Search"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "Audit Logs" Then
               CG = i + 1
               Exit For
            End If
         Next i
         CG = 5
         CI = 2
      Case "Connections"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "Connections" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 6
         CI = 1
      Case "Maps"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "Connections" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 6
         CI = 2
      Case "Modules"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "Connections" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 6
         CI = 3
      Case "System Activity"
         For i = 1 To SSListBar1.Groups.Count
            If SSListBar1.Groups(i).Caption = "Connections" Then
               CG = i + 1
               Exit For
            End If
         Next i
         'CG = 6
         CI = 4
   
   End Select
   
   If CG > 0 Then
      SSListBar1.CurrentGroup = SSListBar1.Groups(CG - 1)
      SSListBar1_ListItemClick SSListBar1.CurrentGroup.ListItems(CI)
      Set LastLVItem = Nothing
   Else
      MsgBox "Unable to show default item with this database version", vbCritical, "Database version incorrect"
      End
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmMain.Form_Load"
   eClass.Add Err.Number, Err.Description, Err.Source, False
   HandleError "frmMain.Form_Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode <> 0 Then
       Cancel = False
       Exit Sub
   End If
   If MsgBox("Are you sure you wish to close ICE...Configuration?", vbQuestion + vbYesNo, "ICE...Configuration") = vbNo Then
      Cancel = True
   End If
   If transCount > 0 Then
      iceCon.RollbackTrans
      transCount = 0
   End If
'   For i = 0 To Forms.Count-1
'      MsgBox (Forms(i).Name)
'   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim frm As Form
   
   If transCount > 0 Then
      iceCon.RollbackTrans
   End If
   If Not (iceCon Is Nothing) Then
      iceCon.Close
   End If
   
   For Each frm In Forms
      If frm.Name <> Me.Name Then
         Unload frm
      End If
   Next
   
   Set iceCon = Nothing
   On Error Resume Next
   Kill fs.BuildPath(App.Path, "*.tmp")
   Set loadCtrl = Nothing
'   Set objTView = Nothing
'   Set objctrl = Nothing
   Set fView = Nothing
   Set eClass = Nothing
   Set fs = Nothing
'   Set tPList = Nothing
End Sub

Public Sub ReadINI()
   INI = App.Path + "\ICEConfig.INI"
   ConfigPath = Read_Ini_Var("General", "ConfigPath", INI)
   DefOrgID = Read_Ini_Var("General", "DefOrgID", INI)
End Sub

Public Function GetOrganisations() As Boolean
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim firstOrg As String
   
'   OrgList.AddItem "Create New..."
   strSQL = "SELECT Organisation_National_Code, Organisation_Name " & _
            "FROM Organisation"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   i = 0
   
   cboTrust.Clear
   OrgList.Clear
   firstOrg = RS!organisation_Name & ""
   
   Do Until RS.EOF
      cboTrust.AddItem RS!Organisation_National_Code
      
      If RS!Organisation_National_Code = DefOrgID Then
         cboTrust.ListIndex = i
         OrgName.Caption = RS!organisation_Name & ""
      End If
      
      i = i + 1
      RS.MoveNext
   Loop
   cboTrust.Locked = (RS.RecordCount < 2)
'   OrgName.Visible = cboTrust.Visible
   
   RS.Close
   
'   cboTrust.AddItem "Add New Trust"
   
   If cboTrust.ListIndex = -1 Then
     cboTrust.ListIndex = 0
     OrgName.Caption = firstOrg
     MsgBox "Default Organisation from config file (" & DefOrgID & ") not found. Organisation set to " & cboTrust.List(0), _
            vbInformation, "Invalid Organisation"
     DefOrgID = cboTrust.List(0)
   End If
   
   strSQL = "SELECT EDI_OrgCode, EDI_Msg_Type, EDI_LTS_Index " & _
            "FROM EDI_Local_Trader_Settings " & _
            "WHERE Organisation = '" & DefOrgID & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.RecordCount = 0 Then
      OrgList.AddItem "None..."
   Else
      Do While Not RS.EOF
         OrgList.AddItem Trim(RS!EDI_OrgCode) & " - " & RS!EDI_Msg_Type
         OrgList.ItemData(OrgList.ListCount - 1) = RS!EDI_LTS_Index
         RS.MoveNext
      Loop
   End If
   
'   OrgList.Locked = (RS.RecordCount < 2)
'   labOrgList.Visible = OrgList.Visible
   OrgList.ListIndex = 0

   RS.Close
   Set RS = Nothing
   GetOrganisations = (OrgList.ListCount > 1)
End Function
'

Private Sub MailBtn_Click()
   On Error GoTo procEH
   loadCtrl.RequeueOptions MailBtn.Caption
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.RequeueFile"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub OpenFileBtn_Click()
   Dim Tstr As String
   Dim tFile As String
   Dim tBuf As New StringBuffer
   Dim fStream As TextStream
   
   If loadCtrl.CurrentLogFile <> "" Then
      Open loadCtrl.CurrentLogFile For Input As #1
      tFile = fs.BuildPath(App.Path, fs.GetTempName)
      Set fStream = fs.CreateTextFile(tFile)
'      fData = tStr
      While Not EOF(1)
         Line Input #1, Tstr
         tBuf.Append Tstr
         If Len(Tstr) > 0 Then
            tBuf.Append vbCrLf
         End If
      Wend
      fStream.Write tBuf.value
      fStream.Close
      Close #1
      Shell "Notepad.EXE " & tFile, vbNormalFocus '  fraPanel(5).Caption, vbNormalFocus
   End If
End Sub

Private Sub optLogSearch_Click(Index As Integer)
   Dim btnId As String
   
   If Index = 0 Then
      chkRestrict.value = 1
      chkRestrict.Enabled = False
'      dtPTo.value = Now()
'      dtPFrom.value = DateAdd("m", -1, dtPTo.value)
      btnId = "0"
      chkRestrict.value = 1
   Else
      chkRestrict.Enabled = True
      btnId = CStr(Index)
   End If
   fView.SetUpPanel Fra_LOGSEARCH, btnId
End Sub

Private Sub itemAdd_Click()
   On Error GoTo procEH
   Dim pNode As MSComctlLib.Node
   Dim natCode As String
   Dim i As Integer
   Dim blnReset As Boolean
   
   blnReset = True
   
   EdiPr.Tag = "New"
   loadCtrl.MenuAddEntry
   
   blnReset = False
   
   If blnReset Then
      For i = 1 To EdiPr.PropertyItems.Count
         EdiPr(i).value = EdiPr(i).defaultValue
      Next i
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.frmMain.itemAdd_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub
'
Private Sub itemDelete_Click()
 On Error GoTo procEH
   If MsgBox("This will delete the entry for " & TreeView1.SelectedItem.Text & ". Are you sure?", vbYesNo, "ICEConfig") = vbYes Then
      loadCtrl.Delete objTV.ActiveNode
      
      If Not (TreeView1.SelectedItem Is Nothing) And nOrigin <> "T" Then
         objTV.ActiveNode = TreeView1.SelectedItem
         objTV.ActiveNode = TreeView1.SelectedItem
         NodeClick TreeView1.SelectedItem
      End If
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.frmMain.itemDelete_Click"
   eClass.FurtherInfo = TreeView1.SelectedItem.Text & " for origin " & nOrigin
   eClass.Add Err.Number, Err.Description, Err.Source, False
   HandleError "frmmMain.itemDelete_Click"
End Sub

Private Sub optShowErr_Click(Index As Integer)
   LogSearchOpt = Index
End Sub

Private Sub optStatPeriod_Click(Index As Integer)
   On Error GoTo procEH
   loadCtrl.ViewingOption = Index
   With DTStatEnd
      Select Case Index
         Case 0
            .MaxDate = DTStats.value
            .value = DTStats.value
         
         Case 1
            .value = DTStats.value - 7
         
         Case 2
            .value = DTStats.value - 28
         
         Case 3
            .value = DTStats.value
      
      End Select
   End With
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmMain.OptStatPeriod.Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Public Sub xOrgList_Click()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   If OrgList.Text <> "None..." Then
      If OrgList.Text = "Create New..." Then
         MsgBox "Not yet available"
      Else
         strSQL = "SELECT EDI_LTS_Index, EDI_Msg_Type, EDI_OrgCode " & _
                  "FROM EDI_Local_Trader_Settings " & _
                  "WHERE Organisation = '" & cboTrust.Text & "' " & _
                     "AND EDI_OrgCode = '" & Trim(Left(OrgList.Text, InStr(1, OrgList.Text, "-") - 2)) & "' " & _
                     "AND EDI_Msg_Type = '" & Trim(Mid(OrgList.Text, InStr(1, OrgList.Text, "-") + 1)) & "'"
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         LTS_DataStream = RS!EDI_Msg_Type
         LTS_Index = RS!EDI_LTS_Index
         LTS_OrgCode = RS!EDI_OrgCode
         RS.Close
         Set RS = Nothing
         
         OrgName.Caption = GetOrganisationName(OrgList.Text)
         SSListBar1.Enabled = True
         TreeView1.Visible = False
         TreeView1.Nodes.Clear
         TreeView1.Visible = True
         If Not loadCtrl Is Nothing Then
            loadCtrl.FirstView
         End If
      End If
   End If
End Sub

Private Sub ssCmdEdiCancel_Click()
   Dim strSQL As String
   
   If transCount > 0 Then
      iceCon.RollbackTrans
      transCount = 0
   End If
   
   fView.Show Fra_HELP, ""
   fView.RefreshProc = ""
   fView.RefreshProcParams = ""
'   If objTView.NodeOrigin = "U" Then
'      loadCtrl.Practices OrgList.Text
'   Else
      fView.RefreshDisplay "objTView"
'   End If
'   blnAddSpec = False
      
End Sub

Private Sub ssCmdEDIok_Click()
   On Local Error GoTo procEH
   Dim PageId As Integer
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim strSQL As String
   Dim sqlField As String
   Dim sqlValue As String
   Dim vSet As String
   Dim pNode As MSComctlLib.Node
   Dim tableOrProc As String
   Dim blnTable As Boolean
   Dim blnNew As Boolean
   Dim intRet As Integer
   Dim tNode As MSComctlLib.Node
   Dim natCode As String
   Dim newValue As String
   Dim strArray() As String
   Dim fieldId As String
   Dim failInfo As String
   Dim indId As Long
   Dim sPos As Integer
   Dim ePos As Integer
   Dim pLen As Integer
   
   '  Update database with new values
   EdiPr.Redraw = False
   If TypeName(loadCtrl) <> "Nothing" Then
      PageId = EdiPr.ActivePage
      tableOrProc = EdiPr.Pages(PageId).Caption
      blnTable = (dbObject(EdiPr.Pages(PageId).Caption) = "UserTable")
      blnNew = (Left(objTV.ActiveNode.Text, 3) = "New") Or objTV.newNode
      
      For i = 1 To EdiPr.PropertyItems.Count
'        Is the key on the active page?
         If EdiPr(i).PageKeys = EdiPr.Pages(PageId).Key Then
            If EdiPr(i).Tag <> "" And EdiPr(i).Tag <> "MENU" Then
               If InStr(1, EdiPr(i).Description, "Mandatory") > 0 Then
                  If EdiPr(i).value = "" Then
                     EdiPr.Redraw = True
                     Err.Raise 16001, "Properties List", "Mandatory Field (" & EdiPr(i).Caption & ") not supplied"
                  End If
               End If
                  
               If blnTable Then
                  fieldId = Mid(EdiPr(i).Tag, InStr(EdiPr(i).Tag, ".") + 1)
   '              Is this value boolean?
                  If (EdiPr(i).value = "True" Or EdiPr(i).value = "False") Then
   '                 Yes, so set the bit value accordingly
                     If EdiPr(i).value = "True" Then
                        vSet = "1"
                     Else
                        vSet = "0"
                     End If
                  Else
                     If EdiPr(i).DialogTitle = "" Or EdiPr(i).DialogTitle = "*" Or EdiPr(i).Key = "PROV_ID" Then
                        vSet = "'" & Trim(EdiPr(i).value) & "'"
                     Else
'                        sPos = InStr(ediPr(i).DialogTitle, "-") + 1
                        ePos = InStr(EdiPr(i).DialogTitle, "-") - 1
'                        pLen = (Len(ediPr(i).DialogTitle) - ePos) + 1
                        vSet = "'" & Left$(EdiPr(i).DialogTitle, ePos) & "'"
'                        vSet = "'" & Mid$(ediPr(i).DialogTitle, sPos, pLen) & "'"
                     End If
                  End If
                  
                  If blnNew Then
                     sqlField = sqlField & fieldId & ", "
                     sqlValue = sqlValue & vSet & ", "
                  Else
                     strSQL = strSQL & fieldId & " = " & vSet & ", "
                  End If
               End If
            End If
         End If
      Next
      transId.StartTransaction
'      If transCount = 0 Then
'         ICECon.BeginTrans
'         transCount = 1
'      End If
      
      If blnTable Then
         If blnNew Then
            strSQL = "INSERT INTO " & EdiPr.Pages(PageId).Caption & " (" & strSQL & _
                      Left(sqlField, Len(sqlField) - 2) & ") VALUES (" & Left(sqlValue, Len(sqlValue) - 2) & ") "
            sqlField = ""
            sqlValue = ""
         Else
            strSQL = "UPDATE " & EdiPr.Pages(PageId).Caption & _
                     " SET " & Left(strSQL, Len(strSQL) - 2) & EdiPr.Caption
         End If
         iceCon.Execute strSQL
         sqlField = ""
         sqlValue = ""
         
      Else
   '     A Stored Procedure needs to be run
         
         Select Case EdiPr.Pages(PageId).Key
            Case "Rules"
               loadCtrl.Update objTV.ActiveNode  ' objTView.NodeKey(objTView.ActiveNode.Key)
            
'            Case "Config"
'               loadCtrl.Update
'               vSet = objTView.NodeKey(TreeView1.SelectedItem.Key)
'               If vSet <> "DFLT" Then
'                  NewValue = ediPr(vSet).value
'                  TreeView1.SelectedItem.Text = NewValue
'               End If
            
            Case Else
               newValue = loadCtrl.Update(EdiPr.Pages(PageId).Key)
         
         End Select
      End If
      
      transId.EndTransaction
'      If transCount > 0 Then
'         ICECon.CommitTrans
'      End If
'      transCount = 0
      
      newValue = loadCtrl.Refresh
      
      loadCtrl.RunWhat objTV.RefreshNode, False
      If TypeName(loadCtrl) = "LoadMySettings" Then
         SSListBar1_ListItemClick frmMain.SSListBar1.ListItems("My Settings")
      End If
   End If
   EdiPr.Redraw = True
   Exit Sub
         
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   If Err.Number = 3157 Then
      loadCtrl.TidyUp
      eClass.ClearErrors
      transId.AbandonTransaction
   Else
      eClass.CurrentProcedure = "ssCmdEDIok.Click"
      eClass.FurtherInfo = IIf(iceCmd.CommandText = "", strSQL, iceCmd.CommandText)
      eClass.Add Err.Number, Err.Description, Err.Source, False
      HandleError "ssCmdEDIok", (transCount > 0)
   End If
End Sub

Private Sub SSListBar1_GroupClick(ByVal GroupClicked As Listbar.SSGroup, ByVal PreviousGroup As Listbar.SSGroup)
   Dim i As Integer
   'fView.Show Fra_HELP, "1"
   Timer2.Enabled = False
   
   SSListBar1.Groups("EDI Management").ListItems("RepList Entries").Text = "Current Rep List Entries"
   If GroupClicked.Key = "Audit Logs" Then
      chkRestrict.value = 1
'      dtPFrom.value = DateAdd("m", -1, Now())
'      dtPTo.value = Now()
'      Set rCtrl = New requeueControl
   End If
   'fView.Show Fra_HELP
   Select Case GroupClicked.Index
      Case 6
         'fView.Show Fra_HELP, "1"
    
   End Select
End Sub

Public Sub SSListBar1_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
   On Local Error GoTo procEH
   Dim i As Integer
   Dim Temp(2) As String
   Dim objDisplay As Object
   Dim strObjId As String
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
    
   itemId = ""
   Timer2.Enabled = False
   MailBtn.Visible = False
   OpenFileBtn.Visible = False
   TreeView1.Visible = False
   TreeView1.Nodes.Clear
   Set loadCtrl = Nothing
   eClass.FurtherInfo = ItemClicked.Text
   frapanel(12).Visible = False
   fView.ShowRCFrame = False
   EdiPr.Redraw = False
   objTV.Origin = 22
   
   SSListBar1.Groups("EDI Management").ListItems("RepList Entries").Text = "Current Rep List Entries"
   
   Set loadCtrl = Nothing
   fView.RefreshProcParams = ""
   
   frmMain.OrgList.Visible = False
   frmMain.labOrgList.Visible = False
   
   fView.Hide
   blnShowBrowser = True
   
   Select Case ItemClicked.Text
'      Case "Request Details"
'         Set loadCtrl = New LoadTests
'         fView.RefreshProc = "Providers"
'
'      Case "Panels and Pages"
'         Set loadCtrl = New LoadPanels
'         fView.RefreshProc = "Panels"
'
'      Case "Profiles"
'         Set loadCtrl = New LoadProfiles
'         fView.RefreshProc = "LoadOrgProfiles"
'         fView.RefreshProcParams = cboTrust.Text
'
'      Case "Picklists"
'         Set loadCtrl = New LoadPickLists
'         fView.RefreshProc = "FirstView"
'
'      Case "Rules"
'         Set loadCtrl = New LoadRules
'         fView.RefreshProc = "ListRules"
'
'      Case "Colours"
'         Set loadCtrl = New LoadColours
'         fView.RefreshProc = "LoadColours"
'
'      Case "Location"
'         Set loadCtrl = New LoadLocations
'         fView.RefreshProc = "FirstView"
'
'      Case "Clinicians"
'         Set loadCtrl = New LoadClinicians
'         fView.RefreshProc = "FirstView"

'      Case "Users"
'
'      Case "Permissions"
'         Set loadCtrl = New LoadUsers
'         fView.RefreshProc = "FirstView"
'
'      Case "Terminology"
'         Set loadCtrl = New LoadCatsAndPriorities ' New LoadTerminology
'         fView.RefreshProc = "FirstView"
'
'      Case "Request Sample Panel Options"
'         Set loadCtrl = New LoadSamplePanelOptions
'         fView.RefreshProc = "FirstView"
'
'      Case "Request Sample Panels"
'         Set loadCtrl = New LoadSamplePanels
'         fView.RefreshProc = "FirstView"
'
'      Case "Blood/Histo/Micro"
'         Set loadCtrl = New LoadBHM
'         fView.RefreshProc = "FirstView"
'
'      Case "System Configuration"
'         Set loadCtrl = New LoadConfiguration
'         fView.RefreshProc = "FirstView"
'
      Case "Connections"
         Set loadCtrl = New LoadConnections
         fView.RefreshProc = "FirstView"
        
      Case "Maps"
         Set loadCtrl = New LoadMaps
         fView.RefreshProc = "FirstView"
        
      Case "Modules"
         Set loadCtrl = New LoadModules
         fView.RefreshProc = "FirstView"
        
      Case "System Activity"
         fView.FrameToShow = Fra_TESTDETAILS
         Set loadCtrl = New LoadMonitors
         fView.RefreshProc = "FirstView"
         'fView.Show fra_None
         blnShowBrowser = True
        
      Case "My Settings"
         Set loadCtrl = New LoadMySettings
         fView.RefreshProc = "FirstView"
        
      Case "EDI Recipients"
         Set loadCtrl = New LoadEDIRecipients
         fView.FrameToShow = Fra_EDI
         fView.RefreshProc = "FirstView"
        
      Case "EDI Clinicians"
         Set loadCtrl = New LoadEDIClinicians
         fView.RefreshProc = "FirstView"
        
      Case "Result Mapping"
         Set loadCtrl = New LoadResultMapping
         'fView.ShowRCFrame = True
         fView.RefreshProc = "FirstView"
         fView.ShowReadCodes
        
      Case "UOM Mapping"
         Set loadCtrl = New LoaduommAP
         fView.RefreshProc = "FirstView"
        
      Case "Specimen Types"
         Set loadCtrl = New LoadSpecimenCodes
         fView.RefreshProc = "FirstView"
        
      Case "Körner Medical Specialties"
         Set loadCtrl = New LoadDisciplineMap
         fView.RefreshProc = "ServiceDisciplines"
        
      Case "Logs"
         Set loadCtrl = New LoadLogs
         loadCtrl.SearchInProgress = False
         LogSearchOpt = 0
         LogText.Text = ""
         dtPFrom.value = DateAdd("m", -1, Now())
         dtPTo.value = Now()
         txtSrchSurname.Text = ""
         txtSrchForename.Text = ""
         'fView.Show Fra_LOGVIEW
'         fView.Show Fra_INFO, "Please wait whilst Log details for " & OrgList.Text & " are loaded..."
         fView.RefreshProc = "FirstView"
         
      Case "Search"
         blnShowBrowser = False
         LogText.Text = ""
         fView.RefreshProc = ""
         
         frmMain.MousePointer = vbNormal
         optLogSearch(0).value = True
         ComboSrchPractice.Clear
   
         strSQL = "SELECT DISTINCT er.EDI_NatCode, EDI_Name " & _
                  "FROM EDI_Recipients er " & _
                     "INNER JOIN EDI_Recipient_Individuals ei " & _
                        "INNER JOIN EDI_Matching em " & _
                        "ON ei.Individual_Index=em.Individual_Index " & _
                     "ON er.EDI_NatCode = EDI_Org_NatCode " & _
                  "ORDER BY EDI_Name"
   
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         Do Until RS.EOF
            ComboSrchPractice.AddItem Trim(RS!EDI_NatCode & " - " & RS!EDI_Name)
            RS.MoveNext
         Loop
         
         RS.Close
         Set RS = Nothing
         
         Set loadCtrl = New LoadLogs
         loadCtrl.SearchInProgress = True
         
         fView.FrameToShow = Fra_LOGVIEW
         fView.Show Fra_LOGSEARCH, "0"
         
      Case "Requeue"
         Set loadCtrl = New LoadRequeue
         fView.RefreshProc = "FirstView"
         
      Case "Current Rep List Entries"
         Set loadCtrl = New LoadRepList
         fView.RefreshProc = "firstView" '"ReadRepList"
         
      Case "Filter"
         Set loadCtrl = New LoadFilter
         fView.RefreshProc = "FirstView"
      
      Case "About"
        frmAbout.Show 1
        Exit Sub
        
      Case "Key Generation"
        Shell App.Path + "\ICEKeyGen.EXE", vbNormalFocus
        Exit Sub
        
      Case "Helpdesk"
         fView.RefreshProc = ""
         HelpDesk_Click
      
      Case Else
         MsgBox "This section has not yet been implemented.  Please check with your supplier for a later version.", vbInformation + vbOKOnly, "ICE...Configuration"
         Set loadCtrl = Nothing
         Set loadCtrl = New LoadEDIRecipients
         fView.RefreshProc = "FirstView"
         fView.RefreshProcParams = cboTrust.Text
         SSListBar1.CurrentGroup = 3
        
   End Select
   fView.RefreshDisplay
   EdiPr.Redraw = True
   wb.Visible = blnShowBrowser
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "SSListbar1.ListItemClick"
   eClass.FurtherInfo = ItemClicked.Text
   eClass.Add Err.Number, Err.Description, Err.Source, False
   HandleError "frmMain.ListItemClick"
End Sub

Private Sub Text1_Change()
'    Command7.Enabled = True
'    Command8.Enabled = True
'    Command5.Enabled = False
'    Command6.Enabled = False
'    TreeView1.Enabled = False
'    SSListBar1.Enabled = False
'    OrgList.Enabled = False
End Sub
Private Sub Text5_Change()
   Text5.Text = Replace(Text5.Text, "'", "`")
   Text5.SelStart = Len(Text5.Text)
End Sub

Private Sub Timer1_Timer()
    Set TreeView1.DropHighlight = TreeView1.HitTest(mfx, mfy)
    If m_iScrollDir = -1 Then 'Scroll Up
    ' Send a WM_VSCROLL message 0 is up and 1 is down
      SendMessage TreeView1.hwnd, 277&, 0&, vbNull
    Else 'Scroll Down
      SendMessage TreeView1.hwnd, 277&, 1&, vbNull
    End If
End Sub

Private Sub Timer2_Timer()
   loadCtrl.Timer
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
   Dim vData As Variant
   Dim tNode As Node
   Dim strSQL As String
   
   If NewString <> "" Then
      vData = objTV.ReadNodeData(objTV.ActiveNode)
      objTV.ActiveNode.Text = NewString
      If vData(0) = "CAT" Then
         strSQL = "DELETE FROM Request_Category"
         iceCon.Execute strSQL
         Set tNode = objTV.TopLevelNode.child
         Do Until tNode Is Nothing
            strSQL = "INSERT INTO Request_Category (Category) " & _
                        "VALUES ('" & Left(tNode.Text, 10) & "')"
            iceCon.Execute strSQL
            Set tNode = tNode.Next
         Loop
      Else
         strSQL = "DELETE FROM Request_Priority"
         iceCon.Execute strSQL
         Set tNode = objTV.TopLevelNode.child
         Do Until tNode Is Nothing
            strSQL = "INSERT INTO Request_Priority (Priority) " & _
                        "VALUES ('" & Left(tNode.Text, 10) & "')"
            iceCon.Execute strSQL
            Set tNode = tNode.Next
         Loop
      End If
      TreeView1.HotTracking = True
   End If
   loadCtrl.FirstView
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
   On Error GoTo NoAction
   loadCtrl.Collapse Node
   
NoAction:
   ' procedure doesn't exist - ignore
End Sub

Public Sub TreeView1_DblClick()
   Dim vData As Variant
   
   vData = objTV.ReadNodeData(objTV.ActiveNode)
   Set TreeView1.SelectedItem = objTV.ActiveNode
   If vData(1) = "EDIT" Then
      TreeView1.HotTracking = False
      TreeView1.StartLabelEdit
   End If
End Sub

Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
   Debug.Print "Drag_Drop"
'   Dim dhData As Variant
'   Dim siData As Variant
'   Dim BN1 As String
'   Dim BN2 As String
'   Dim tt As String
'   Dim tk As String
'   Dim moNode As MSComctlLib.node
'
'   If Not TreeView1.DropHighlight Is Nothing Then
'      siData = objTV.ReadNodeData(TreeView1.SelectedItem)
'      dhData = objTV.ReadNodeData(TreeView1.DropHighlight)
'
'      Select Case Left(objTV.NodeOrigin, 1)
'         Case "T"
'            If objTV.NodeKey(TreeView1.DropHighlight) = siData(1) And _
'               TreeView1.DropHighlight.Key <> TreeView1.SelectedItem.Key Then
'               tk = TreeView1.SelectedItem.Key
'               tt = TreeView1.SelectedItem.Text
'               If Not TreeView1.DropHighlight Is Nothing Then
'                  If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' below '" & TreeView1.DropHighlight.Text & "'?", vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
'                     TreeView1.Nodes.Remove (tk)
'                     Set moNode = TreeView1.Nodes.Add(TreeView1.DropHighlight, tvwNext, tk, tt, 1, 1)
'                     moNode.Tag = "*"
'                     moNode.Checked = True
'                     loadCtrl.WriteTestRules TreeView1.DropHighlight.Parent
'                  End If
'               End If
'
'            ElseIf TreeView1.DropHighlight.Text = "Rules" And _
'               siData(1) = dhData(0) Then
'               tk = TreeView1.SelectedItem.Key
'               tt = TreeView1.SelectedItem.Text
'               If Not TreeView1.DropHighlight Is Nothing Then
'                  If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' above '" & TreeView1.DropHighlight.Child.Text & "'?", vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
'                     TreeView1.Nodes.Remove (tk)
'                     Set moNode = TreeView1.Nodes.Add(TreeView1.DropHighlight.Child, tvwFirst, tk, tt, 1, 1)
'                     moNode.Checked = True
'                     moNode.Tag = "*"
'                     loadCtrl.WriteTestRules TreeView1.DropHighlight
'                  End If
'               End If
'            End If
'
'         Case "P"
'            If TreeView1.DragIcon = ImageList1.ListImages("Cogs").Picture Then
''              Move a page below a page
'               If objTV.NodeKey(TreeView1.DropHighlight) = "PAGE" Then
'                  tk = TreeView1.SelectedItem.Key
'                  tt = TreeView1.SelectedItem.Text
'                  If Not TreeView1.DropHighlight Is Nothing Then
'                     If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' below '" & TreeView1.DropHighlight.Text & "'?", vbYesNo + vbQuestion, "Re-order Pages") = vbYes Then
'                        TreeView1.Nodes.Remove (tk)
'                        Set moNode = TreeView1.Nodes.Add(TreeView1.DropHighlight, tvwNext, tk, tt, 2, 2)
'                        moNode.Checked = True
'                        loadCtrl.WritePageSequence TreeView1.DropHighlight.Parent
'                        frmMain.NodeClick moNode
'                     End If
'                  End If
'
'               ElseIf objTV.NodeKey(TreeView1.DropHighlight) = "PANEL" Then
'                  If objTV.NodeKey(TreeView1.SelectedItem) = "PANEL" Then
'                     tk = TreeView1.SelectedItem.Key
'                     tt = TreeView1.SelectedItem.Text
'                     If Not TreeView1.DropHighlight Is Nothing Then
'                        If TreeView1.DropHighlight = TreeView1.SelectedItem.Root Then
'                           If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' above '" & TreeView1.DropHighlight.Next.Text & "'?", vbYesNo + vbQuestion, "Re-order Panels") = vbYes Then
'                              loadCtrl.WritePanelSequence TreeView1.SelectedItem, TreeView1.DropHighlight
'                           End If
'                        Else
'                           If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' below '" & TreeView1.DropHighlight.Text & "'?", vbYesNo + vbQuestion, "Re-order Panels") = vbYes Then
'                              loadCtrl.WritePanelSequence TreeView1.SelectedItem, TreeView1.DropHighlight
'                           End If
'                        End If
'                     End If
'
'                  ElseIf objTV.NodeKey(TreeView1.SelectedItem) = "PAGE" Then
'                     tk = TreeView1.SelectedItem.Key
'                     tt = TreeView1.SelectedItem.Text
'                     If Not TreeView1.DropHighlight Is Nothing Then
'                        If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' above '" & TreeView1.DropHighlight.Child.Text & "'?", vbYesNo + vbQuestion, "Re-order Panels") = vbYes Then
'                           TreeView1.Nodes.Remove (tk)
'                           Set moNode = TreeView1.Nodes.Add(TreeView1.DropHighlight.Child, tvwFirst, tk, tt, 2, 2)
'                           moNode.Checked = True
'                           loadCtrl.WritePageSequence TreeView1.DropHighlight
'                        End If
'                     End If
'                  End If
'               End If
'               Set TreeView1.SelectedItem = moNode
'            End If
'
'      End Select
'   End If
'   Set TreeView1.DropHighlight = Nothing
'   Set moNode = Nothing
'   Timer1.Enabled = False
End Sub

Private Sub TreeView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
'   Dim mfx As Single
'   Dim mfy As Single
'   Dim BN1 As String
'   Dim BN2 As String
'   Dim dhData As Variant
'   Dim siData As Variant
'
'   Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
'   mfx = x
'   mfy = y
'
'   If y > 0 And y < 100 Then 'scroll up
'      m_iScrollDir = -1
'      Timer1.Enabled = True
'   ElseIf y > (TreeView1.Height - 200) And y < TreeView1.Height Then
'   'scroll down
'      m_iScrollDir = 1
'      Timer1.Enabled = True
'   Else
'      Timer1.Enabled = False
'   End If
'
'   If TreeView1.DropHighlight Is Nothing Then
'      TreeView1.DragIcon = ImageList1.ListImages(3).Picture
'      Exit Sub
'   End If
'
'   dhData = objTV.ReadNodeData(TreeView1.DropHighlight)
'   siData = objTV.ReadNodeData(TreeView1.SelectedItem)
'
'   BN1 = objTV.NodeKey(TreeView1.DropHighlight)
'   BN2 = objTV.NodeKey(TreeView1.SelectedItem)
'   Select Case Left(objTV.NodeOrigin, 1)
'      Case "T"
'         If dhData(2) <> "RuleDetails" And dhData(2) <> "Rules" Then
'            TreeView1.DragIcon = ImageList1.ListImages(3).Picture
'            Exit Sub
'         End If
'
'         If (BN2 <> BN1) And (dhData(0) <> siData(1)) Then
'            TreeView1.DragIcon = ImageList1.ListImages(3).Picture
'            Exit Sub
'         End If
'         TreeView1.DragIcon = ImageList1.ListImages(1).Picture
'
'      Case "P"
'         If objTV.NodeKey(TreeView1.SelectedItem) = "PAGE" Then
''         Debug.Print TreeView1.SelectedItem.Text & ": " & siData(0) & TreeView1.DropHighlight.Text & ": " & dhData(0)
'            If siData(0) <> dhData(0) Or _
'               TreeView1.DropHighlight.Key = TreeView1.SelectedItem.Key Then
'               TreeView1.DragIcon = ImageList1.ListImages(3).Picture
'               Exit Sub
'            End If
'            TreeView1.DragIcon = ImageList1.ListImages("Cogs").Picture
'         Else
'            If objTV.NodeKey(TreeView1.DropHighlight) <> "PANEL" Or _
'               TreeView1.DropHighlight.Key = TreeView1.SelectedItem.Key Then
'               TreeView1.DragIcon = ImageList1.ListImages(3).Picture
'               Exit Sub
'            End If
'            TreeView1.DragIcon = ImageList1.ListImages("Cogs").Picture
'         End If
'
'   End Select
''   Debug.Print "dh: " & dhData(0) & " " & dhData(1) & vbCrLf & _
'          "SI: " & siData(0) & " " & siData(1)
'
End Sub

Public Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
   On Error GoTo procEH
'   Dim blnCollapsed As Boolean
'   blnCollapsed = (node.Tag <> "E")
   objTV.ActiveNode = Node
   
   If Node.Expanded = True Then
      loadCtrl.RunWhat Node, True
'      node.Tag = "E"
'      If Not (node.Child Is Nothing) Then
'         node.Child.LastSibling.EnsureVisible
'      End If
   End If
   wb.Visible = blnShowBrowser
   
   If blnShowBrowser Then
      If fView.FrameToShow <> Fra_TESTDETAILS Then
         fView.Hide
      End If
'      fView.FrameToShow = fra_None
'      fView.Show fra_None
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.frmMain.Treeview1_Expand"
   eClass.FurtherInfo = TypeName(loadCtrl) & " - Node Id = " & Node.Text
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Set TreeView1.SelectedItem = TreeView1.HitTest(x, y)
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim vData As Variant
   
   TreeView1.DropHighlight = TreeView1.HitTest(x, y)
   'Make sure we are over a Node
   If Not TreeView1.DropHighlight Is Nothing Then
'      Set moNode = TreeView1.HitTest(x, y)
      TreeView1.SelectedItem = TreeView1.HitTest(x, y)
      Set moNode = TreeView1.SelectedItem ' Set the item being dragged.
   
      If Button = 1 Then
         blnShowBrowser = False
         NodeClick TreeView1.SelectedItem
         wb.Visible = blnShowBrowser
      Else
         nOrigin = objTV.NodeOrigin(TreeView1.SelectedItem)
         
         If nOrigin <> "T" Then
            NodeClick TreeView1.SelectedItem
         Else
            vData = objTV.ReadNodeData(TreeView1.SelectedItem)
            objTV.MenuStatus = vData(3)
         End If
         
         If objTV.MenuStatus <> 0 And _
            objTV.newNode = False Then
            PopupMenu item
         End If
      End If
   End If
   Set TreeView1.DropHighlight = Nothing
'   Debug.Print "End of MouseDown: " & TreeView1.SelectedItem.Text
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If objTV.NodeOrigin = "COL" Then
      If TreeView1.HitTest(x, y) Is Nothing Then
         frapanel(13).Visible = False
         frmMain.labShowCol.BackColor = &H808080
      Else
         frapanel(13).Visible = True
'         objTV.ActiveNode = TreeView1.HitTest(x, y)
         frmMain.labShowCol.BackColor = Val(objTV.nodeKey(objTV.TopLevelNode(TreeView1.HitTest(x, y))))
      End If
   Else
      frapanel(13).Visible = False
   End If
End Sub

Public Sub NodeClick(ByVal Node As MSComctlLib.Node)
   On Local Error GoTo procEH
   Dim vData As Variant
   
   EdiPr.Redraw = False
   Set newNode = Node
   objTV.ActiveNode = TreeView1.SelectedItem
   vData = objTV.ReadNodeData(Node)
   objTV.MenuStatus = vData(3)
   loadCtrl.RunWhat newNode, True
   frmMain.MousePointer = 0
   TreeView1.Visible = True
   EdiPr.Redraw = True
   
   If blnShowBrowser Then
      If fView.FrameToShow <> Fra_TESTDETAILS Then
         fView.Hide
      End If
   Else
      fView.Show
   End If
   
   Exit Sub
    
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.frmMain.NodeClick"
   eClass.Add Err.Number, Err.Description, Err.Source, False
   frmMain.MousePointer = vbNormal
End Sub

Private Sub TreeView1_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   Effect = loadCtrl.TV_DragDrop(TreeView1.HitTest(x, y), data)
End Sub

Private Sub TreeView1_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   Dim tNode As Node
   
   If Button = vbLeftButton Then
      Set tNode = TreeView1.HitTest(x, y)
      If tNode Is Nothing Then
         Effect = vbDropEffectNone
      Else
         Effect = loadCtrl.TV_DragOver(tNode)
      End If
   Else
      Effect = vbDropEffectNone
   End If
End Sub

Private Sub TreeView1_OLEStartDrag(data As MSComctlLib.DataObject, AllowedEffects As Long)
   Dim vData As Variant
   
   objTV.ActiveNode = TreeView1.SelectedItem
   vData = objTV.ReadNodeData(TreeView1.SelectedItem)
   If Left(vData(4), 4) = "DRAG" Then
      AllowedEffects = vbDropEffectMove
      data.SetData vData(1)
   Else
      AllowedEffects = vbDropEffectNone
   End If
End Sub

Private Sub txtSrchForename_Validate(Cancel As Boolean)
   txtSrchForename.Text = Replace(txtSrchForename.Text, "`", "'")
   
   If InStr(1, txtSrchForename.Text, "'") = 0 Then
      txtSrchForename.ToolTipText = ""
   Else
      txtSrchForename.Text = Replace(txtSrchForename.Text, "'", "''")
      txtSrchForename.ToolTipText = "Apostrophe escaped"
   End If
End Sub

Private Sub txtSrchSurname_Validate(Cancel As Boolean)
   txtSrchSurname.Text = Replace(txtSrchSurname.Text, "`", "'")
   
   If InStr(1, txtSrchSurname.Text, "'") = 0 Then
      txtSrchSurname.ToolTipText = ""
   Else
      txtSrchSurname.Text = Replace(txtSrchSurname.Text, "'", "''")
      txtSrchSurname.ToolTipText = "Apostrophe escaped"
   End If
End Sub
