VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5370
   ClientLeft      =   2100
   ClientTop       =   2670
   ClientWidth     =   8070
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7560
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3405
      ScaleHeight     =   195
      ScaleWidth      =   2535
      TabIndex        =   16
      Top             =   4755
      Width           =   2535
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "http://www.ahsl.com"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   0
         Width           =   1500
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   975
      Left            =   2040
      TabIndex        =   10
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1720
      _Version        =   131073
      BevelOuter      =   1
      Begin VB.Label Label8 
         Caption         =   "A Company"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "A Person"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "System Info..."
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   30
      Left            =   2040
      TabIndex        =   2
      Top             =   4305
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   53
      _Version        =   131073
      BevelOuter      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   9128
      _Version        =   131073
      BevelInner      =   1
      Begin VB.Image Image5 
         Height          =   480
         Left            =   600
         Picture         =   "frmAbout.frx":030A
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   600
         Picture         =   "frmAbout.frx":0BD4
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   600
         Picture         =   "frmAbout.frx":0EDE
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   600
         Picture         =   "frmAbout.frx":17A8
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   600
         Picture         =   "frmAbout.frx":2072
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   825
         Left            =   75
         Picture         =   "frmAbout.frx":293C
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   1560
      End
   End
   Begin VB.Label Label12 
      Caption         =   "or call our sales team on +44 (0) 1603 819600."
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   4950
      Width           =   4095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Member of the ICE family of products"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   14
      Top             =   720
      Width           =   2595
   End
   Begin VB.Label Label9 
      Caption         =   "For information on other products in the ICE family, please see our website at "
      Height          =   735
      Left            =   2040
      TabIndex        =   13
      Top             =   4560
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "This product is licensed to:"
      Height          =   195
      Left            =   2040
      TabIndex        =   9
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Label Label5 
      Caption         =   $"frmAbout.frx":DBA0
      Height          =   675
      Left            =   2040
      TabIndex        =   8
      Top             =   3600
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ICE...Products use signing and encryption technology from Entrust."
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   4710
   End
   Begin VB.Label Label3 
      Caption         =   "ICE...Configuration is Copyright © Anglia Healthcare Systems Ltd, 1994 - 2001. All Rights Reserved."
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
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
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ICE...Configuration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   5355
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------------
' Copyright ©1996-2001 VBnet, Randy Birch, All Rights Reserved.
' Terms of use http://www.mvps.org/vbnet/terms/pages/terms.htm
'--------------------------------------------------------------


Private Const clrLinkActive = vbBlue
Private Const clrLinkHot = vbRed
Private Const clrLinkInactive = vbBlack

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
  (lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, _
   lpPoint As POINTAPI) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
Private Sub Timer1_Timer()
   Dim pt As POINTAPI
   Dim x As Long
   Dim y As Long
   With Picture1
      GetCursorPos pt
      ScreenToClient .hwnd, pt
      x = pt.x * Screen.TwipsPerPixelX
      y = pt.y * Screen.TwipsPerPixelY
      If (x < 0) Or (x > .Width) Or _
         (y < 0) Or (y > .Height) Then
            Label11.ForeColor = clrLinkInactive
            Label11.Font.Underline = False
            Timer1.Enabled = False
      End If
   End With
End Sub
Private Sub Label11_Click()
   Dim sURL As String
   sURL = Label11.Caption
   Call RunShellExecute("open", sURL, 0&, 0&, SW_SHOWNORMAL)
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   With Label11
      If .ForeColor = clrLinkActive Then
         .ForeColor = clrLinkHot
         .Refresh
      End If
   End With
End Sub
Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   With Label11
      If .ForeColor = clrLinkHot Then
         .ForeColor = clrLinkActive
         .Refresh
      End If
   End With
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   With Label11
      If .ForeColor = clrLinkInactive Then
         .ForeColor = clrLinkActive
         .Font.Underline = True
         Timer1.Interval = 100
         Timer1.Enabled = True
      End If
   End With
End Sub
Private Sub RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long)
    Call ShellExecute(GetDesktopWindow(), sTopic, sFile, sParams, sDirectory, nShowCmd)
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    Call StartSysInfo
End Sub
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Private Sub Form_Load()
    Dim Name As String
    Dim Company As String
    Dim SerialNo As String
    Label2.Caption = Label2.Caption + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
    GetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Anglia Healthcare Systems Ltd\ICE", "Name", Name
    GetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Anglia Healthcare Systems Ltd\ICE", "Company", Company
    GetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Anglia Healthcare Systems Ltd\ICE", "Serial Number", SerialNo
    Label7.Caption = Name
    Label8.Caption = Company
    Label13.Caption = SerialNo
    With Label11
        .Move 0, 0
        .ForeColor = clrLinkInactive
        Picture1.Move Picture1.Left, Picture1.Top, .Width, .Height
    End With
End Sub
