VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFilter 
   Caption         =   "EDI Report Filtering"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   525
      Left            =   2595
      TabIndex        =   30
      Top             =   5400
      Width           =   1650
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   375
      Left            =   4470
      TabIndex        =   25
      Top             =   315
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   16449539
      CurrentDate     =   37666
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   375
      Left            =   1620
      TabIndex        =   24
      Top             =   315
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   16449539
      CurrentDate     =   37666
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   480
      TabIndex        =   3
      Top             =   930
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "EDI Reports"
      TabPicture(0)   =   "frmFilter.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "optAll"
      Tab(0).Control(2)=   "optAck"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Message Errors"
      TabPicture(1)   =   "frmFilter.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Crypt/DTS"
      TabPicture(2)   =   "frmFilter.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ack Details"
      TabPicture(3)   =   "frmFilter.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).ControlCount=   1
      Begin VB.OptionButton optAck 
         Caption         =   "Unacknowledged"
         Height          =   270
         Left            =   -71670
         TabIndex        =   29
         Top             =   405
         Width           =   1740
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   270
         Left            =   -72870
         TabIndex        =   28
         Top             =   405
         Width           =   675
      End
      Begin VB.Frame Frame4 
         Caption         =   "Show files with..."
         Height          =   2295
         Left            =   -74520
         TabIndex        =   14
         Top             =   600
         Width           =   5295
         Begin VB.CheckBox optOKLevel 
            Caption         =   "Inactive Tests"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox optOKLevel 
            Caption         =   "Suppressed Samples"
            Height          =   315
            Index           =   6
            Left            =   255
            TabIndex        =   21
            Top             =   1800
            Width           =   2610
         End
         Begin VB.CheckBox optOKLevel 
            Caption         =   "Deleted Read Codes"
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox optOKLevel 
            Caption         =   "Inapplicable Read Codes"
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   19
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox optOKLevel 
            Caption         =   "No UOM with Read Code"
            Height          =   315
            Index           =   3
            Left            =   2880
            TabIndex        =   18
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox optOKLevel 
            Caption         =   "Suppressed Tests"
            Height          =   315
            Index           =   5
            Left            =   2880
            TabIndex        =   17
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox optOKLevel 
            Caption         =   "Suppressed Battery Headers"
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CheckBox optOKLevel 
            Caption         =   "Others"
            Height          =   315
            Index           =   7
            Left            =   2880
            TabIndex        =   15
            Top             =   1800
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Show files with..."
         Height          =   1455
         Left            =   -74520
         TabIndex        =   11
         Top             =   600
         Width           =   5295
         Begin VB.CheckBox optAckLevel 
            Caption         =   "Invalid Acknowledgement"
            Height          =   315
            Index           =   2
            Left            =   1560
            TabIndex        =   23
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox optAckLevel 
            Caption         =   "Interchange rejected"
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox optAckLevel 
            Caption         =   "Interchange partially rejected"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Show files with..."
         Height          =   2055
         Left            =   -74520
         TabIndex        =   8
         Top             =   600
         Width           =   5295
         Begin VB.CheckBox optDTSLevel 
            Caption         =   "DTS errors"
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   10
            Top             =   360
            Width           =   2295
         End
         Begin VB.CheckBox optDTSLevel 
            Caption         =   "Cryptography Errors"
            Height          =   315
            Index           =   0
            Left            =   255
            TabIndex        =   9
            Top             =   360
            Width           =   1860
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Show files with..."
         Height          =   1455
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   5295
         Begin VB.CheckBox optMsgLevel 
            Caption         =   "Parsing errors"
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox optMsgLevel 
            Caption         =   "All tests suppressed"
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   6
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox optMsgLevel 
            Caption         =   "Data Integrity Errors"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox txtInstruct 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   495
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmFilter.frx":0070
      Top             =   4305
      Width           =   6135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   4755
      TabIndex        =   1
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Set as default"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   5415
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Up to..."
      Height          =   255
      Left            =   3810
      TabIndex        =   27
      Top             =   390
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Show from..."
      Height          =   255
      Left            =   705
      TabIndex        =   26
      Top             =   405
      Width           =   960
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msgFlag(4) As Long
Private blnDateChanged As Boolean
Private strSQL As String
Private blnChangeFilter As Boolean
Private LTS_Index As Long
Private LTS_OrgCode As String
Private LTS_DataStream As String

Public Property Get ApplyChanges() As Boolean
   ApplyChanges = blnChangeFilter
End Property

Public Sub BuildView(StartDate As String, _
                     EndDate As String)
   strSQL = "ALTER VIEW ImpExp_View AS (" & vbCrLf & _
            "SELECT Service_ImpExp_Headers.Service_ImpExp_ID, " & vbCrLf & _
                  "  EDI_NatCode, " & vbCrLf & _
                  "  EDI_Name, " & vbCrLf & _
                  "  ImpExp_File, " & vbCrLf & _
                  "  EDI_LTS_Index, " & vbCrLf & _
                  "  Control_Ref, " & vbCrLf & _
                  "  Header_Status, " & vbCrLf & _
                  "  Comment_Status, " & vbCrLf & _
                  "  Case (Header_Status & 0x40000110) " & vbCrLf & _
                  "     When 0x40000110 Then 256 " & vbCrLf & _
                  "     When 0x40000010 Then 0 " & vbCrLf & _
                  "     Else 256 " & vbCrLf & _
                  "  End As Ack_Status, " & vbCrLf & _
                  "  Last_Comment_Added, " & vbCrLf & _
                  "  Date_Added " & vbCrLf & _
            "FROM Service_ImpExp_Headers " & vbCrLf & _
               "  LEFT JOIN EDI_Recipients er " & vbCrLf
               
   If phoenix Then
      strSQL = strSQL & "  ON PatIndex('%' + EDI_NatCode + '%',ImpExp_File) > 0 " & vbCrLf
   Else
      strSQL = strSQL & _
               "    INNER JOIN EDI_Recipient_Ref err " & vbCrLf & _
               "    ON er.Ref_Index=err.Ref_Index " & vbCrLf & _
               "  ON Trader_Code = EDI_Trader_Account + EDI_Free_Part " & vbCrLf

   End If
      
   'strSQL = strSQL & vbCrLf & _
            "WHERE Service_Type = 2 " & vbCrLf & _
               "  AND DateDiff(d,'" & StartDate & "',Last_Comment_Added) >= 0 " & vbCrLf & _
               "  AND DateDiff(d, Last_Comment_Added, '" & EndDate & "') >= 0)" & vbCrLf
            
'            "GROUP BY Service_ImpExp_Headers.Service_ImpExp_ID, EDI_NatCode, ImpExp_File, Comment_Status, Header_Status,Service_ImpExp_Headers.Date_Added)"
   
   strSQL = strSQL & vbCrLf & _
            "WHERE Service_Type = 2 " & vbCrLf & _
            "   AND Last_Comment_Added between '" & StartDate & " 00:00:00' " & " and '" & EndDate & " 23:59:59')"

   iceCon.Execute strSQL
End Sub

Private Sub cmdApply_Click()
   MessageFlag SSTab1.Tab
   If blnDateChanged Then
      BuildView Format(dtpStart.value, "yyyymmdd"), Format(dtpEnd.value, "yyyymmdd")
   End If
   blnDateChanged = False
   Me.Hide
   blnChangeFilter = True
End Sub

Private Sub cmdClose_Click()
   blnChangeFilter = False
   Me.Hide
End Sub

Private Sub cmdDefault_Click()
   Dim i As Integer
   
   For i = 0 To 3
      Write_Ini_Var "DEFAULT", "LogFilter_" & i, CStr(MessageFlag(i)), iniFile
   Next i
End Sub

Private Sub dtpEnd_CloseUp()
   blnDateChanged = True
   If DateDiff("d", dtpEnd.value, Now()) < 0 Then
      dtpEnd.value = Now()
   End If
   
   If DateDiff("d", dtpStart.value, dtpEnd.value) < 0 Then
      dtpStart.value = dtpEnd.value
   End If
End Sub

Private Sub dtpStart_CloseUp()
   blnDateChanged = True
   If DateDiff("d", dtpStart.value, Now()) < 0 Then
      dtpStart.value = Now()
   End If
   
   If DateDiff("d", dtpStart.value, dtpEnd.value) < 0 Then
      dtpEnd.value = dtpStart.value
   End If
End Sub

Private Sub Form_Activate()
   SSTab1.Tab = frmNewLogView.SSTab1.Tab
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   For i = 0 To 3
      msgFlag(i) = Val(Read_Ini_Var("Default", "logFilter_" & i, iniFile))
      BuildFlagOptions i
   Next i
   
End Sub

Private Sub BuildFlagOptions(tabIndex As Integer)
   Dim i As Integer
   Dim flag As Long
   
   flag = msgFlag(tabIndex)
   Select Case tabIndex
      Case 0
         optOKLevel(0).value = IIf((flag And (IS_TEST_INACTIVE Or IS_INV_INACTIVE)) = 0, 0, 1)
         optOKLevel(1).value = IIf((flag And IS_RC_NA) = 0, 0, 1)
         optOKLevel(2).value = IIf((flag And IS_RC_DELETED) = 0, 0, 1)
         optOKLevel(3).value = IIf((flag And IS_RC_REMOVED) = 0, 0, 1)
         optOKLevel(4).value = IIf((flag And IS_INV_SUPPRESSED) = 0, 0, 1)
         optOKLevel(5).value = IIf((flag And IS_TEST_SUPPRESSED) = 0, 0, 1)
         optOKLevel(6).value = IIf((flag And IS_SAMPLE_SUPPRESSED) = 0, 0, 1)
         optOKLevel(7).value = IIf((flag And (MS_CONFORMANCE Or MS_REQUEUE)) = 0, 0, 1)
         optAck.value = (flag And IS_NO_ACK) = IS_NO_ACK
         optAll.value = Not (optAck.value)
         
      Case 1
         optMsgLevel(0).value = IIf((flag And MS_DATA_INTEGRITY) = 0, 0, 1)
         optMsgLevel(1).value = IIf((flag And RS_SUPPRESSION) = 0, 0, 1)
         optMsgLevel(2).value = IIf((flag And MS_PARSE_FAIL) = 0, 0, 1)
'         optMsgLevel(3).value = IIf((flag And (MS_ACK_REJECT_ALL Or MS_ACK_REJECT_PART)) = 0, 0, 1)
      
      Case 2
         optDTSLevel(0).value = IIf((flag And MS_CRYPT_FAIL) = 0, 0, 1)
         optDTSLevel(1).value = IIf((flag And MS_DTS_FAIL) = 0, 0, 1)
         
      Case 3
         optAckLevel(0).value = IIf((flag And MS_ACK_REJECT_PART) = 0, 0, 1)
         optAckLevel(1).value = IIf((flag And MS_ACK_REJECT_ALL) = 0, 0, 1)
         optAckLevel(2).value = IIf((flag And MS_ACK_FAIL) = 0, 0, 1)
      
   End Select
   
   If flag = 0 Then
      flag = &HFFFFFFFF
   End If
   
   msgFlag(tabIndex) = flag Or &H10000000
End Sub

Public Property Let CurrentFlag(tabIndex As Integer, lngNewValue As Long)
   msgFlag(tabIndex) = lngNewValue
End Property

Public Property Get CurrentLTSDataStream() As String
   CurrentLTSDataStream = LTS_DataStream
End Property

Public Property Get CurrentLTSIndex() As Long
   CurrentLTSIndex = LTS_Index
End Property

Public Property Get CurrentLTSOrg() As String
   CurrentLTSOrg = LTS_OrgCode
End Property
   
Public Function MessageFlag(FlagId As Integer) As Long
   Dim i As Integer
   Dim flag As Long
   Dim blnSetFlag As Boolean
   
   flag = 0
   Select Case FlagId
      Case 0
         For i = 0 To optOKLevel.Count - 1
            blnSetFlag = (optOKLevel(i).value = 1)
            If blnSetFlag Then
               Select Case i
                  Case 0
                     flag = flag Or (IS_TEST_INACTIVE Or IS_INV_INACTIVE)
                  
                  Case 1
                     flag = flag Or IS_RC_NA
                  
                  Case 2
                     flag = flag Or IS_RC_DELETED
                  
                  Case 3
                     flag = flag Or IS_RC_REMOVED
                  
                  Case 4
                     flag = flag Or IS_INV_SUPPRESSED
                  
                  Case 5
                     flag = flag Or IS_TEST_SUPPRESSED
                     
                  Case 6
                     flag = flag Or IS_SAMPLE_SUPPRESSED
                  
                  Case 7
                     flag = flag Or (MS_CONFORMANCE Or MS_REQUEUE)
                  
               End Select
            End If
         Next i
         
         If optAck.value Then
            flag = flag Or IS_NO_ACK
         End If
         
         flag = flag Or &H80000000
      
      Case 1
         For i = 0 To optMsgLevel.Count - 1
            blnSetFlag = (optMsgLevel(i).value = 1)
            If blnSetFlag Then
               Select Case i
                  Case 0
                     flag = flag Or MS_DATA_INTEGRITY
                  
                  Case 1
                     flag = flag Or RS_SUPPRESSION 'Or IS_TEST_SUPPRESSED Or IS_INV_SUPPRESSED Or IS_SAMPLE_SUPPRESSED
                  
                  Case 2
                     flag = flag Or MS_PARSE_FAIL
                  
'                  Case 3
'                     flag = flag Or (MS_ACK_REJECT_ALL Or MS_ACK_REJECT_PART)
'
               End Select
            End If
         Next i
         
         If flag = 0 Then
            flag = &H2C600
         End If
      
      Case 2
         For i = 0 To optDTSLevel.Count - 1
            blnSetFlag = (optDTSLevel(i).value = 1)
            If blnSetFlag Then
               Select Case i
                  Case 0
                     flag = flag Or MS_CRYPT_FAIL
                  Case 1
                     flag = flag Or MS_DTS_FAIL
                  
               End Select
            End If
         Next i
      
      Case 3
         For i = 0 To optAckLevel.Count - 1
            blnSetFlag = (optAckLevel(i).value = 1)
            If blnSetFlag Then
               Select Case i
                  Case 0
                     flag = flag Or MS_ACK_REJECT_PART
                     
                  Case 1
                     flag = flag Or MS_ACK_REJECT_ALL
                     
                  Case 2
                     flag = flag Or MS_ACK_FAIL
                     
               End Select
            End If
         Next i
         
         If flag = 0 Then
            flag = &HE00
         End If
         
   End Select
   
   msgFlag(FlagId) = flag
   MessageFlag = flag
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
   MessageFlag PreviousTab
End Sub

