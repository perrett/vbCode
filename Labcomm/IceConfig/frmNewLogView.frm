VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2D47C3AF-9C7A-44E4-9FCC-CDE675667A2D}#1.0#0"; "Web_Browser_Control.ocx"
Begin VB.Form frmNewLogView 
   Caption         =   "EDI Report Summary"
   ClientHeight    =   10860
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Web_Browser_Control_V2.wbCtrl wb 
      Height          =   6855
      Left            =   7080
      TabIndex        =   18
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12091
   End
   Begin RichTextLib.RichTextBox rtfKey 
      Height          =   1995
      Left            =   7065
      TabIndex        =   16
      Top             =   7860
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3519
      _Version        =   393217
      BackColor       =   12648447
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmNewLogView.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRCKeys 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Show Read Code Key"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7320
      Width           =   3735
   End
   Begin VB.CommandButton cmdRequeue 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Requeue selected..."
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10080
      Width           =   2535
   End
   Begin VB.CommandButton cmdFilter 
      BackColor       =   &H0080FFFF&
      Caption         =   "Amend Tracking Filters..."
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Change the tracking filter"
      Top             =   10080
      Width           =   2895
   End
   Begin RichTextLib.RichTextBox rtfComments 
      Height          =   2055
      Left            =   600
      TabIndex        =   4
      Top             =   7800
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   3625
      _Version        =   393217
      BackColor       =   12648447
      ScrollBars      =   3
      TextRTF         =   $"frmNewLogView.frx":0080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6885
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   12144
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "EDI Reports"
      TabPicture(0)   =   "frmNewLogView.frx":010F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPbar(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tvMsg(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Message Errors"
      TabPicture(1)   =   "frmNewLogView.frx":012B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvMsg(1)"
      Tab(1).Control(1)=   "lblPbar(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Crypt/DTS"
      TabPicture(2)   =   "frmNewLogView.frx":0147
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tvMsg(2)"
      Tab(2).Control(1)=   "lblPbar(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Ack Errors"
      TabPicture(3)   =   "frmNewLogView.frx":0163
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblPbar(3)"
      Tab(3).Control(1)=   "tvMsg(3)"
      Tab(3).ControlCount=   2
      Begin MSComctlLib.TreeView tvMsg 
         Height          =   6135
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   10821
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvMsg 
         Height          =   6135
         Index           =   1
         Left            =   -74880
         TabIndex        =   3
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   10821
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvMsg 
         Height          =   6135
         Index           =   2
         Left            =   -74880
         TabIndex        =   6
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   10821
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvMsg 
         Height          =   6135
         Index           =   3
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   10821
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.Label lblPbar 
         Alignment       =   2  'Center
         Caption         =   "Processing, please wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -74280
         TabIndex        =   13
         Top             =   3240
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblPbar 
         Alignment       =   2  'Center
         Caption         =   "Processing, please wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74280
         TabIndex        =   12
         Top             =   3240
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblPbar 
         Alignment       =   2  'Center
         Caption         =   "Processing, please wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74160
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblPbar 
         Alignment       =   2  'Center
         Caption         =   "Processing, please wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   3000
         Visible         =   0   'False
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10800
      TabIndex        =   0
      Top             =   10080
      Width           =   2415
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5190
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewLogView.frx":017F
            Key             =   "Scroll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewLogView.frx":0499
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewLogView.frx":07B3
            Key             =   "PhoneFolder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewLogView.frx":108D
            Key             =   "Query"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewLogView.frx":13A7
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewLogView.frx":16C1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   6255
      Left            =   6720
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   11033
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label lblLTS 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   255
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tracking Messages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Menu mnuLocalTrader 
      Caption         =   "Local Trader Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuLTS 
         Caption         =   "Local Trader Settings"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmNewLogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCtrl As New ManageControls
Private vData As Variant
Private cmtFlag As Long
Private blnKeyVisible As Boolean
Private sBuf As New StringBuffer
Private chkStatus As Integer
Private maxLen As Long
Private MaxUOM As Long
Private MaxRES  As Long
Private maxRange As Long
Private tStop(6) As Long
Private invStop(4) As Long
Private sampText As String
Private sampCnt As Integer
Private blnRCInapplicable As Boolean
Private blnRCDelete As Boolean
Private blnRCNoUOM As Boolean
Private blnRCSuppBatt As Boolean
Private blnRCSuppTest As Boolean
Private blnRCInactive As Boolean
Private curRepId As Long
Private repStatus As String
'Private blob As New clsDBBlob
Private resTable As String
Private NodeToClear As Node

Private repHTMLHeader As String
Private outOfRange As String
Private sampHTML As String
Private repHTMLComment As String
Private invHTML As String
Private repHTMLHR As String
Private resHTML As String
Private resHTML_OOR As String
Private repHTMLPIndent As String
Private repHTMLPEnd As String
Private repHTMLTableEnd As String
Private repHTMLTableStart As String
Private resHTMLTableStart As String

Private repLTS_Index As Long
Private LTS_Index As Long
Private LTS_OrgCode As String
Private LTS_DataStream As String

Private rTot As Integer
Private dtsAPI As New clsDtsAPI
Private lastDTSId As String

Private objRTF As rtfSelection
Private collRTF As New Collection
                 
Private Sub cmdClose_Click()
   Unload frmFilter
   Unload Me
End Sub

Private Sub cmdFilter_Click()
   frmFilter.Show 1
   If frmFilter.ApplyChanges Then
      SSTab1_Click SSTab1.Tab
   End If
End Sub

Private Sub cmdRCKeys_Click()
   Dim tFile As String
   Dim Tstr As String
   Dim tBuf As New StringBuffer
   Dim fStream As TextStream
   
   If blnKeyVisible Then
      rtfKey.Visible = False
      cmdRCKeys.Caption = "Show Read Code Key"
   Else
      rtfKey.Visible = True
      cmdRCKeys.Caption = "Hide Read Code Key"
   End If
   blnKeyVisible = rtfKey.Visible
End Sub

Private Sub cmdRequeue_Click()
   On Error GoTo procEH
   Dim rNode As Node
   Dim tvid As TreeView
   Dim pbStep As Long
   
   If blnUseDMO Then
      rCtrl.CallerForm = Me
      With frmEDIRequeue
         .optRequeue(0).Enabled = (SSTab1.Tab <> 1)
         .Show 1
         Set tvid = tvMsg(SSTab1.Tab)
      
         Me.Hide
      
         frmShowRequeue.fgReq.Tag = 1
         Me.Show
         
         If .RequeueValue <> "Cancel" Then
            pBar.Visible = True
            rTot = 0
            
            If .RequeueValue = "Reprocess" Then
               rCtrl.RequeueReport = True
               
               Set rNode = tvid.Nodes(1)
               nodeRequeue rNode, rNode.Checked
               Do Until rNode.Next Is Nothing
                  Set rNode = rNode.Next
                  nodeRequeue rNode, rNode.Checked
               Loop
               
            ElseIf .RequeueValue = "Resend" Then
               rCtrl.RequeueReport = False
               
               Set rNode = tvid.Nodes(1)
               nodeFiles rNode, rNode.Checked
               
               Do Until rNode.Next Is Nothing
                  Set rNode = rNode.Next
                  nodeFiles rNode, rNode.Checked
               Loop
               
            End If
               
            If rTot = 0 Then
               MsgBox "No valid reports or files selected!", vbExclamation, "Nothing to Requeue"
               Unload frmShowRequeue
            Else
               pBar.value = 0
               pbStep = rTot / 10
               pbStep = (pbStep - (pbStep Mod 8)) / 8
               pBar.max = rTot + (pbStep * 8)
               pBar.Visible = True
               rCtrl.UseProgressBar pBar, pbStep
               rCtrl.RequeueData
               pBar.Visible = False
            End If
         End If
      End With
      GetDates
   
      Unload frmEDIRequeue
   Else
      MsgBox "Requeue facility not available on this machine." & vbCrLf & _
             "The 'SQLDMO' option is unavailable. Either ask your IT department to install " & vbCrLf & _
             "the option or use the management machine to requeue", vbInformation, "SQL Option not installed"
   End If
Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.cmdRequeue_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub xcmdRequeue_Click()
   On Error GoTo procEH
   Dim nd(4) As Node
   Dim tvid As TreeView
   Dim rTot As Integer
   Dim pbStep As Long
'   Dim rCtrl As New requeueControl
   
   rCtrl.CallerForm = Me
   With frmEDIRequeue
      .optRequeue(0).Enabled = (SSTab1.Tab <> 1)
      .Show 1
      Set tvid = tvMsg(SSTab1.Tab)
   
      Me.Hide
   
      If .RequeueValue = "Reprocess" Then
         rCtrl.RequeueReport = True
      ElseIf .RequeueValue = "Resend" Then
         rCtrl.RequeueReport = False
      End If
      
      frmShowRequeue.fgReq.Tag = 1
      Me.Show
      If .RequeueValue <> "Cancel" Then
         pBar.Visible = True
         rTot = 0
         Set nd(0) = tvid.Nodes(1)
'         Set nd(1) = nd(0).Child
         
         Do Until nd(0) Is Nothing
            Set nd(1) = nd(0).child
            If InStr(1, nd(0).Key, "GetPractices") > 0 Then
               Do Until nd(1) Is Nothing
                  Set nd(2) = nd(1).child
                  Do Until nd(2) Is Nothing
                     If .RequeueValue = "Reprocess" Then
                        Set nd(3) = nd(2).child
                        Do Until nd(3) Is Nothing
                           If nd(3).Checked Then
                              vData = objTV.ReadNodeData(nd(3))
                              rCtrl.RequeueItem = vData(1)
                              Debug.Print vData(1) & ": " & nd(3).Text
                              rTot = rTot + 1
                           End If
                           Set nd(3) = nd(3).Next
                        Loop
                     Else
                        If nd(2).Checked Then
                           vData = objTV.ReadNodeData(nd(2))
                           rCtrl.RequeueItem = vData(0)  '  fs.GetFileName(vData(1))
                           rTot = rTot + 1
                        End If
                     End If
                     
                     Set nd(2) = nd(2).Next
                  Loop
                  Set nd(1) = nd(1).Next
               Loop
            Else
               Do Until nd(1) Is Nothing
                  If frmEDIRequeue.RequeueValue = "Reprocess" Then
                     Set nd(2) = nd(1).child
                     Do Until nd(2) Is Nothing
                        If nd(2).Checked Then
                           vData = objTV.ReadNodeData(nd(2))
                           rCtrl.RequeueItem = vData(1)
                           rTot = rTot + 1
                        End If
                        Set nd(2) = nd(2).Next
                     Loop
                  Else
                  
                     If nd(1).Checked And nd(1).Children > 0 Then
                        vData = objTV.ReadNodeData(nd(1))
                        rCtrl.RequeueItem = vData(0)  '  fs.GetFileName(vData(1))
                        rTot = rTot + 1
                     End If
                  End If
                  Set nd(1) = nd(1).Next
               Loop
            End If
            Set nd(0) = nd(0).Next
         Loop
         
         If rTot = 0 Then
            MsgBox "No valid reports or files selected!", vbExclamation, "Nothing to Requeue"
            Unload frmShowRequeue
         Else
            pBar.value = 0
            pbStep = rTot / 10
            pbStep = (pbStep - (pbStep Mod 8)) / 8
            pBar.max = rTot + (pbStep * 8)
            pBar.Visible = True
            rCtrl.UseProgressBar pBar, pbStep
            rCtrl.RequeueData
            pBar.Visible = False
         End If
      End If
   End With
   GetDates
   
'   Set rCtrl = Nothing
   Unload frmEDIRequeue
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.cmdRequeue_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Public Function CallBack(Operation As String, _
                         Optional Parameter As Variant) As String
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   
   eClass.FurtherInfo = "Parameter = " & IIf(VarType(Parameter) = vbError, "<None>", Parameter)
   
   Select Case Operation
      Case "FileToNotepad"
         OpenInNotepad CStr(Parameter)
         
      Case "LocationFileName"
         CallBack = wb.Tag
      
      Case Else
         If curRepId > 0 Then
            If CInt(Parameter) = 1 Then
               With iceCmd
                  .ActiveConnection = iceCon
                  .CommandType = adCmdStoredProc
                  .CommandText = Operation
                  .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , curRepId)
                  .Parameters.Append .CreateParameter("Type", adInteger, adParamInput, , CInt(Parameter))
                  .Parameters.Append .CreateParameter("File", adVarChar, adParamOutput, 255)
                  .Execute
                  wb.Tag = .Parameters("File").value & ""
                  wb.LocationTitle = fs.GetFileName(wb.Tag)
               End With
            Else
               vData = objTV.ReadNodeData(tvMsg(SSTab1.Tab).SelectedItem.Parent)
               wb.Tag = vData(1)
               wb.LocationTitle = tvMsg(SSTab1.Tab).SelectedItem.Parent.Text
            End If
         End If
   End Select
   Exit Function
   
procEH:
'   MsgBox Err.Number & ": " & Err.Description
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.CallBack(" & Operation & ")"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function ConvertToRGB(ByVal Color As String) As String
    ' Are we dealing with a hex color
    If Left$(Color, 1) <> "&" Then
        Dim lngColor As Long
        Dim lngRemainder As Long
        Dim strHexPart As String
        Dim strOutput As String
        
        lngColor = CLng(Color)
        
        While lngColor > 0
            lngRemainder = lngColor Mod 16
            If lngRemainder < 10 Then
                strHexPart = CStr(lngRemainder)
            Else
                Select Case lngRemainder
                    Case 10
                        strHexPart = "A"
                    Case 11
                        strHexPart = "B"
                    Case 12
                        strHexPart = "C"
                    Case 13
                        strHexPart = "D"
                    Case 14
                        strHexPart = "E"
                    Case 15
                        strHexPart = "F"
                End Select
            End If
            strOutput = strHexPart & strOutput
            lngColor = (lngColor - lngRemainder) / 16
            
        Wend
        Color = strOutput
        Color = Left("00000", 6 - Len(Color)) & Color
    Else
        Color = Mid$(Color, 5, 6)
        
    End If
    Color = Right$(Color, 2) & Mid$(Color, 3, 2) & Left$(Color, 2)
    ConvertToRGB = Color
End Function

Private Sub Form_Load()
   Dim keyLine(7) As String
   Dim i As Integer
   Dim txtPos As Integer
   Dim oText As TextStream
   Dim strArray() As String
   Dim lBuf As String
   Dim daysToShow As Integer
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim dtsClient As String
   
   wb.DBConnection = iceCon
   Me.ScaleMode = 4
   chkStatus = -1
   
   daysToShow = -Val(Read_Ini_Var("Default", "FilterHistory", iniFile))
   
   If daysToShow = 0 Then
      daysToShow = -21
   End If
   
   With frmFilter
      .dtpStart = DateAdd("d", daysToShow, Now())
      .dtpEnd.value = Now()
      .BuildView Format(DateAdd("d", -21, Now()), "yyyymmdd"), Format(DateAdd("d", 1, Now()), "yyyymmdd")
   End With
   
   keyLine(0) = "[ N/A ]" & vbTab & "Read code not required (Standalone test)" & vbCrLf
   keyLine(1) = "[NONE ]" & vbTab & "No read code" & vbCrLf
   keyLine(2) = "[9999.]" & vbTab & "Not active" & vbCrLf
   keyLine(3) = "[9999.]" & vbTab & "Suppressed" & vbCrLf
   keyLine(4) = "[9999.]" & vbTab & "Flagged as inapplicable" & vbCrLf
   keyLine(5) = "[9999.]" & vbTab & "Flagged as deleted" & vbCrLf
   keyLine(6) = "[9999.]" & vbTab & "Removed - mandatory UOM not present"
   
   With rtfKey
      .Visible = False
      .Text = ""
      txtPos = 0
      
      For i = 0 To 6
         .Text = .Text & keyLine(i)
      Next i
      
      For i = 0 To 6
         .SelStart = txtPos + 1
         .SelLength = 5
         Select Case i
            Case 2
               .SelColor = vbRed
            
            Case 3
               .SelColor = vbRed
               .SelStrikeThru = True
            
            Case 4
               .SelColor = BPBLUE
               
            Case 5
               .SelColor = BPBLUE
               .SelStrikeThru = True
               
            Case 6
               .SelColor = BPBLUE
               .SelUnderline = True
              
         End Select
         txtPos = txtPos + Len(keyLine(i))
         
      Next i
   End With
   
'   Set oText = fs.OpenTextFile(fs.BuildPath(App.Path, "repTemplate1.html"))
   repHTMLHeader = ""
   outOfRange = ""
   sampHTML = ""
   repHTMLComment = ""
   invHTML = ""
'   resHTMLStart = ""
   resHTML = ""
   resHTML_OOR = ""
'   resHTMLEnd = ""
'   repHTMLTrailer = ""
   repHTMLTableStart = ""
   repHTMLTableEnd = ""
   repHTMLHR = ""
   repHTMLPIndent = ""
   
   strArray = wb.TemplateToArray("Report", TT_HTML)
   
   For i = 0 To UBound(strArray)
      lBuf = strArray(i) & vbCrLf
      Select Case i
         Case 0 To 50
            repHTMLHeader = repHTMLHeader & lBuf
            
         Case 51
            repHTMLHR = lBuf
   
         Case 52
            repHTMLTableStart = lBuf
            
         Case 53 To 57
            repHTMLComment = repHTMLComment & lBuf
            
         Case 58
            repHTMLTableEnd = lBuf
            
         Case 59 To 73
            sampHTML = sampHTML & lBuf
      
         Case 74 To 78
            outOfRange = outOfRange & lBuf
   
         Case 79 To 86
            invHTML = invHTML & lBuf
         
         Case 87
            repHTMLPIndent = lBuf
         
         Case 88
            resHTMLTableStart = lBuf
            
         Case 89 To 95
            resHTML = resHTML & lBuf
            
'         Case 98
'            repHTMLPEnd = lBuf
'
'         Case 100 To 105
'            resHTML_OOR = resHTML_OOR & lBuf
'
      End Select
   Next i
   
   With frmMain.OrgList
      If .ListCount = 1 Then
         For i = 0 To 3
            tvMsg(i).ToolTipText = ""
         Next i
         
      Else
         For i = 0 To .ListCount - 1
            If i > mnuLTS.UBound Then
               Load mnuLTS(i)
            End If
            mnuLTS(i).Caption = .List(i)
            mnuLTS(i).Visible = True
         Next i
         
         Load mnuLTS(i)
         mnuLTS(i).Caption = "All datastreams"
         mnuLTS(i).Visible = True
   
      End If
   End With
            
   LTS_Index = 0
   LTS_DataStream = "All datastreams"
   LTS_OrgCode = "All"
   
   wb.InfoCallBack = Me
   SSTab1_Click 0
   
   strSQL = "SELECT Module_Launch_Folder " & _
            "FROM Connect_Modules " & _
            "WHERE Module_Name = 'ICELABDTS.EXE'"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   If Not RS.EOF Then
      dtsClient = fs.BuildPath(RS!Module_Launch_Folder, "DTS\Client\DTSClient.cfg")
   End If
   
   RS.Close
   
   dtsAPI.DTSServer = Read_Ini_Var("General", "DTSServer", iniFile)
   
   If blnUseDTS Then
      blnUseDTS = dtsAPI.FindSiteDTS(dtsClient)
   End If
   
   Set RS = Nothing
End Sub

Public Property Get CurrentDTSEnquiry()
   CurrentDTSEnquiry = lastDTSId
End Property

Private Sub GetDates()
   Dim cText As String
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim msgFlag As Long
   Dim tvid As TreeView
   Dim nd As Node
   
   rtfComments.Text = ""
   cmtFlag = frmFilter.MessageFlag(SSTab1.Tab)
   
   Select Case SSTab1.Tab
      Case 0
         cText = "ICECONFIG_Logs_ReadDates"
         
         If (cmtFlag And IS_NO_ACK) = IS_NO_ACK Then
            msgFlag = (MS_MSGOK Or MS_AWAIT_ACK Or MS_ACK_RECEIVED)
         Else
            msgFlag = MS_MSGOK
         End If
         
'         cmtFlag = &H800011FE
      
      Case 1
         cText = "ICECONFIG_Logs_ReaderrorDates"
         msgFlag = cmtFlag 'MS_PARSE_FAIL Or MS_NO_OUTPUT Or MS_DATA_INTEGRITY
         cmtFlag = -1
         
      Case 2
         cText = "ICECONFIG_Logs_ReaderrorDates"
         msgFlag = cmtFlag 'MS_DTS_FAIL Or MS_CRYPT_FAIL
         cmtFlag = &HFFFFFFFF
      
      
      Case 3
         cText = "ICECONFIG_Logs_ReaderrorDates"
         msgFlag = cmtFlag 'MS_ACK_REJECT_PART Or MS_ACK_REJECT_ALL Or MS_ACK_FAIL
         cmtFlag = -1
      
   End Select
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = cText
      .Parameters.Append .CreateParameter("Status", adInteger, adParamInput, , msgFlag)
      .Parameters.Append .CreateParameter("cStatus", adInteger, adParamInput, , cmtFlag)
      Set RS = .Execute
   End With
   
'  Get the comment flag setting again as it is overwritten when the tab index is 0
'  (No filtering required on dates for OK messages)
   If SSTab1.Tab = 0 Then
      cmtFlag = frmFilter.MessageFlag(0)
   End If
   
   Set tvid = tvMsg(SSTab1.Tab)
   tvid.Nodes.Clear
   If RS.RecordCount = 0 Then
      Set nd = tvid.Nodes.Add(, _
                              , _
                              mCtrl.NewNodeKey(0, _
                                               CStr(cmtFlag), _
                                               "None"), _
                              "<No entries found>", _
                              1)
   Else
      Do Until RS.EOF
         Set nd = tvid.Nodes.Add(, _
                                 , _
                                 mCtrl.NewNodeKey(RS!Processed, _
                                                  CStr(cmtFlag), _
                                                  IIf(SSTab1.Tab = 0, "GetPractices", "GetFiles"), _
                                                  CStr(msgFlag)), _
                                 Format(RS!Processed, "dd/mm/yyyy"), _
                                 1)
         tvid.Nodes.Add nd, _
                        tvwChild, _
                        mCtrl.NewNodeKey(RS!Processed, _
                                        "Tmp", _
                                        IIf(SSTab1.Tab = 0, "GetPractices", "GetFiles"), _
                                        CStr(msgFlag)), _
                        "Please wait...", _
                        1
         
         RS.MoveNext
      Loop
   End If
   
'   wb.NavigateTo "Initial", False
   
   RS.Close
   Set RS = Nothing
   Set iceCmd = Nothing
End Sub

Private Sub GetFiles(NodeId As Node)
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim tvid As TreeView
   Dim nd As Node
   Dim nCol As Long
   
'   rtfData.Text = ""
'   rtfComments.Text = ""
   
   vData = objTV.ReadNodeData(NodeId)
   Set tvid = tvMsg(SSTab1.Tab)
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("Date", adVarChar, adParamInput, 10, vData(0))
      .Parameters.Append .CreateParameter("cStat", adInteger, adParamInput, , vData(1))
      .Parameters.Append .CreateParameter("Status", adInteger, adParamInput, , 0)
      
      Select Case SSTab1.Tab
         Case 0
            .Parameters(1).value = &HFFFFFFFF
            .Parameters("Status").value = vData(4)
            If InStr(1, NodeId.Text, " ") = 0 Then
               .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 6, Null)
            Else
               .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 6, Left(NodeId.Text, (InStr(1, NodeId.Text, " ") - 1)))
            End If
            
            .CommandText = "ICECONFIG_Logs_HeaderOK"
         
         Case 1
            .Parameters("Status").value = (MS_PARSE_FAIL Or MS_NO_OUTPUT Or MS_DATA_INTEGRITY)
            .CommandText = "ICECONFIG_Logs_HeaderErrors"
         
         Case 2
            .Parameters("Status").value = (MS_DTS_FAIL Or MS_CRYPT_FAIL)
            .CommandText = "ICECONFIG_Logs_HeaderErrors"
      
         Case 3
            .Parameters("Status").value = (MS_ACK_REJECT_PART Or MS_ACK_REJECT_ALL Or MS_ACK_FAIL)
            .CommandText = "ICECONFIG_Logs_HeaderErrors"
         
      End Select
      Set RS = .Execute
   End With
   
   Do Until RS.EOF
      If RS!EDI_LTS_Index = LTS_Index Or LTS_Index = 0 Then
         Set nd = tvid.Nodes.Add(NodeId, _
                                 tvwChild, _
                                 mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                                  RS!ImpExp_File, _
                                                  IIf(IsNull(RS!EDI_LTS_Index), "GetComments", "GetReports"), _
                                                  CStr(vData(1))), _
                                 Trim(fs.GetFileName(RS!ImpExp_File)), _
                                 1)
         
         If RS!Header_Status > 0 Then
            If (RS!Header_Status And MS_REQUEUE) = MS_REQUEUE Then
               nd.ForeColor = BPBLUE
            Else
               nd.ForeColor = BPRED
            End If
         Else
            If (RS!Header_Status And (MS_AWAIT_ACK Or MS_ACK_RECEIVED)) = MS_AWAIT_ACK Then
               nd.ForeColor = BPBLUE
            Else
               nd.ForeColor = BPGREEN
            End If
         End If
         
         If Not NodeId Is Nothing Then
            nd.Checked = NodeId.Checked
         End If
         
         If SSTab1.Tab = 0 Then
            If ((RS!Comment_Status And &HFFFFFCFF) And vData(1)) > 0 Then
               nd.Image = 4
            End If
         End If
         
         If Not IsNull(RS!EDI_LTS_Index) Then
            tvid.Nodes.Add nd, _
                           tvwChild, _
                           mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                            RS!ImpExp_File, _
                                            "GetReports", _
                                                  CStr(vData(1))), _
                           "Please wait...", _
                           1
         End If
      End If
      
      RS.MoveNext
   Loop
   
   RS.Close
   
   Set RS = Nothing
   Set iceCmd = Nothing
End Sub

Private Sub GetPractices(NodeId As Node)
   Dim iceCmd As New ADODB.Command
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim nd As Node
   Dim tvid As TreeView
   Dim natCode As String
   
   vData = objTV.ReadNodeData(NodeId)
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_Logs_Practice"
      .Parameters.Append .CreateParameter("Date", adVarChar, adParamInput, 10, vData(0))
      .Parameters.Append .CreateParameter("Status", adInteger, adParamInput, , vData(4))
      Set RS = .Execute
   End With
   
   Set tvid = tvMsg(0)
   Do Until RS.EOF
      If IsNull(RS!EDI_NatCode) Then
         natCode = "No_Trader_Code_Match"
      Else
         natCode = RS!EDI_NatCode & " (" & RS!EDI_Name & ")"
      End If
      Set RS2 = Nothing
      
      Set nd = tvid.Nodes.Add(NodeId, _
                              tvwChild, _
                              mCtrl.NewNodeKey(CStr(vData(0)), _
                                               CStr(vData(1)), _
                                               "GetFiles", _
                                               CStr(vData(4))), _
                              natCode, _
                              1)
      
      If Not NodeId Is Nothing Then
         nd.Checked = NodeId.Checked
      End If
      
      If ((RS!Comment_Flag And &HFFFFFCFF) And vData(1)) > 0 Then
         nd.Image = 4
      End If
      
      tvid.Nodes.Add nd, _
                     tvwChild, _
                     mCtrl.NewNodeKey(CStr(vData(0)), _
                                           CStr(vData(1)), _
                                           "GetFiles", _
                                           CStr(vData(4))), _
                     "Please wait...", _
                     1
      RS.MoveNext
   Loop
   
   RS.Close
   
   Set RS = Nothing
   Set iceCmd = Nothing
End Sub

Private Sub GetReports(NodeId As Node)
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim tvid As TreeView
   Dim nd As Node
   Dim fBuf As String
   Dim repDesc As String
   
   vData = objTV.ReadNodeData(NodeId)
   Set tvid = tvMsg(SSTab1.Tab)
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_Logs_Reports"
      .Parameters.Append .CreateParameter("ImpExp", adInteger, adParamInput, , vData(0))
      .Parameters.Append .CreateParameter("cStat", adInteger, adParamInput, , vData(4))
      Set RS = .Execute
   End With
   
   If RS.RecordCount > 0 Then
      repLTS_Index = RS!EDI_LTS_Index
   End If
   
   Do Until RS.EOF
      repDesc = Trim(RS!Service_Id) & " (" & Trim(RS!Patient_Name) & ")"
      
      If (RS!Message_Status And MS_REQUEUE) Then
         repDesc = "+ " & repDesc
      End If
      
      Set nd = tvid.Nodes.Add(NodeId, _
                              tvwChild, _
                              mCtrl.NewNodeKey(CStr(vData(0)), _
                                               RS!Service_ImpExp_Message_ID, _
                                               "ShowReport", _
                                               RS!Service_Report_Index), _
                              repDesc, _
                              1)
'      If chkStatus > 0 Then
         nd.Checked = NodeId.Checked
'      End If
      
'      If (RS!Message_Status And cmtFlag) > 0 Then
'         nd.ForeColor = BPRED
'      End If
      
      If SSTab1.Tab > 0 Then
         If (RS!Message_Status And frmFilter.MessageFlag(SSTab1.Tab)) > 0 Then
            nd.ForeColor = BPRED
         End If
      End If
      
      
      If Not IsNull(RS!Service_ImpExp_Msg_Id) Then
         nd.Image = 4
      End If
      
      RS.MoveNext
   Loop
   RS.Close
   
   Set RS = Nothing
   Set iceCmd = Nothing
End Sub

Public Property Let MessageFilter(lngNewValue As Long)
   cmtFlag = lngNewValue
End Property

Private Sub NodeClick(NodeId As Node)
   On Error GoTo procEH
   Dim tvid As TreeView
   Dim i As Integer
   
'   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   Set tvid = tvMsg(SSTab1.Tab)
   
   frmNewLogView.AutoRedraw = False
'   tvid.Visible = False
   vData = objTV.ReadNodeData(NodeId)
   eClass.FurtherInfo = NodeId.Text & " (" & vData(2) & ")"
   
   blnKeyVisible = True
   cmdRCKeys.Visible = False
   cmdRCKeys_Click
   curRepId = 0
   
   rtfComments.Text = ""
   Select Case vData(2)
      Case "GetReports"
         ShowFile NodeId
         ShowComments NodeId
'         ShowDTSTracking NodeId
         
      Case "GetComments"
         If fs.FileExists(vData(1)) Then
            ShowFile NodeId
         End If
         ShowComments NodeId
         
      Case "ShowReport"
         ShowReport NodeId
         ShowComments NodeId
         
      Case "GetPractices"
         ShowLog NodeId
         
      Case "GetFiles"
         ShowPractice NodeId
'         wb.NavigateTo "Initial", False
   
   End Select
   frmNewLogView.AutoRedraw = True
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
'   MsgBox Err.Number & ": " & Err.Description
   eClass.CurrentProcedure = "frmNewLogView.NodeClick"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub NodeExpand(NodeId As Node)
   Dim tvid As TreeView
   Dim i As Integer
   
   Set tvid = tvMsg(SSTab1.Tab)
   
   vData = objTV.ReadNodeData(NodeId)
   frmNewLogView.AutoRedraw = False
   
   For i = 1 To NodeId.Children
      tvid.Nodes.Remove NodeId.child.Index
   Next i
   
   Select Case vData(2)
      Case "GetFiles"
         GetFiles NodeId
         
      Case "GetPractices"
         GetPractices NodeId
         
      Case "GetReports"
         GetReports NodeId
         
   End Select
   frmNewLogView.AutoRedraw = True
'   tvid.Visible = True
End Sub

Private Sub ShowComments(NodeId As Node)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim tBuf As String
   Dim rData As clsDTSRequest
   Dim strRes As String
   
   vData = objTV.ReadNodeData(NodeId)
   cmtFlag = frmFilter.MessageFlag(SSTab1.Tab)
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      
      If vData(2) = "ShowReport" Then
         .Parameters.Append .CreateParameter("ImpExp", adInteger, adParamInput, , vData(1))
         .CommandText = "ICECONFIG_Logs_Comments_By_Report"
      Else
         .Parameters.Append .CreateParameter("ImpExp", adInteger, adParamInput, , vData(0))
         .CommandText = "ICECONFIG_Logs_Comments"
      End If
      .Parameters.Append .CreateParameter("cStat", adInteger, adParamInput, , cmtFlag)
      
      Set RS = .Execute
   End With
   
   rtfComments.Text = ""
   sBuf.Clear
   Do Until RS.EOF
      sBuf.Append RS!Date_Added & ": " & RS!Service_ImpExp_Comment & vbCrLf
      RS.MoveNext
   Loop
   
   rtfComments.Text = sBuf.value
      
   strRes = ""
   
   If vData(2) = "GetReports" And blnUseDTS And SSTab1.Tab = 0 Then
      DoEvents
      RS.MoveFirst
      RS.Filter = "LocalId <> ''"
      
      If RS.EOF = False Then
         Set rData = dtsAPI.ReadDTSResponse(RS!LocalId)
         
         lastDTSId = RS!LocalId
         
         With rData
            wb.DTSData .MsgId, _
                       .LocalId, _
                       .Recipient, _
                       .FromSMTPAddress, _
                       .ToSMTPAddress, _
                       .TransferredOn, _
                       .CurrentStatus, _
                       .SentOn
               
         End With
      Else
         wb.DTSData "", "", "", "", "", "N/A", "Awaiting Sending via DTS", ""
      End If
      
      RS.Close
   End If
   
   Set RS = Nothing
   Set iceCmd = Nothing
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.ShowComments"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub ShowFile(NodeId As Node)
   On Error GoTo procEH
   Dim fBuf As String
   Dim bodyHdr As String
   Dim qt As String
   Dim lc As Long
   
   qt = Chr(34)
   
   If blnUseDTS And SSTab1.Tab = 0 Then
      wb.NavigationBar = 3
   Else
      wb.NavigationBar = 0
   End If
   
   wb.InfoBGColour = "#FFFFCC"
   
   vData = objTV.ReadNodeData(NodeId)
   sBuf.Clear
   
'   rtfData.Tag = vData(1)
   wb.Tag = vData(1)
   wb.LocationTitle = fs.GetFileName(wb.Tag)
   wb.Style = "files"
   
   If UCase(fs.GetExtensionName(wb.Tag)) <> "XML" Then
      Open vData(1) For Input As #1
      While Not EOF(1)
         Line Input #1, fBuf
         sBuf.Append fBuf
         If Len(fBuf) > 0 Then
            sBuf.Append vbCrLf
         End If
         lc = lc + 1
      Wend
      Close #1
   
      fBuf = Replace(sBuf.ActualValue, "?'", "#apos#")
      
      If lc < 3 Then
         fBuf = Replace(fBuf, "'", "</td></tr><tr><td>" & vbCrLf)
      Else
         fBuf = Replace(fBuf, vbCrLf, "</td></tr><tr><td>" & vbCrLf)
      End If
      
      fBuf = Replace(fBuf, "#apos#", "'")
      
      bodyHdr = "<div NOWRAP>" & vbCrLf & "<Table Width=100%><tr><td>"
      wb.NavigateTo bodyHdr & fBuf & "</TR></table>"
   Else
      wb.NavigateTo wb.Tag, True
   End If
   Exit Sub
   
procEH:
'   If eClass.Behaviour = -1 Then
'      Stop
'      Resume
'   End If
   If Err.Number = 76 Or Err.Number = 53 Or Err.Number = 71 Then
      wb.NavigateTo "<div align=" & Chr(34) & "center" & Chr(34) & "><font COLOR=" & Chr(34) & "#ff0000" & Chr(34) & ">" & _
                    "<p>&lt&lt Errors encountered attempting to display file &gt&gt</font></p>" & _
                     "<p>" & vData(1) & "</p>" & vbCrLf & _
                     "<p><font COLOR=" & Chr(34) & "#0000ff" & Chr(34) & ">" & "Error " & Err.Number & "<br>" & _
                     "<br>(" & Err.Description & ")</font></p>"
   Else
      eClass.CurrentProcedure = "frmNewLogView.ShowFile"
      eClass.Add Err.Number, Err.Description, Err.Source
   End If
End Sub

Public Sub ShowPractice(NodeId As Node)
   On Error GoTo procEH
   Dim tableHdr As String
   Dim htmlLST As String
   Dim htmlKeyHdr As String
   Dim htmlKey1 As String
   Dim htmlSpecHdr As String
   Dim htmlSpec As String
   Dim buf As String
   Dim sBuf As New StringBuffer
   Dim strSQL As String
   Dim ltRS As New ADODB.Recordset
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   
   tableHdr = "<table cellspacing='0' width='100%'>" & vbCrLf
   
   htmlLST = "<tr class=lts>" & vbCrLf & _
             " <td align='center' colspan='2'>Datastream: </td>" & vbCrLf & _
             " <td>(#LTS#)</td>" & vbCrLf & _
             "</tr>" & vbCrLf
             
   htmlKeyHdr = "<tr class=hdr>" & vbCrLf & _
                " <td colspan='3'>Local Practice Keys</td>" & vbCrLf & _
                "</tr>" & vbCrLf & _
                "<tr class=hdr>" & vbCrLf & _
                " <td>Local Key</td>" & vbCrLf & _
                " <td>No. of key 3's</td>" & vbCrLf & _
                " <td>&nbsp;</td>" & vbCrLf & _
                "</tr>" & vbCrLf
                
   htmlKey1 = "<tr class=key>" & vbCrLf & _
              "   <td align=center>#KEY1#</td>" & vbCrLf & _
              "   <td align=center>#KEY3#</td>" & vbCrLf & _
              "   <td>&nbsp;</td>" & vbCrLf & _
              "</tr>" & vbCrLf
              
   htmlSpecHdr = "<tr class=hdr>" & vbCrLf & _
                 " <td colspan='3'>Specialties</td>" & vbCrLf & _
                 "</tr>" & vbCrLf & _
                 "<tr class=hdr>" & vbCrLf & _
                 "   <td>Code</td>" & vbCrLf & _
                 "   <td >Description</td>" & vbCrLf & _
                 "   <td>Sent via</td align='left'>" & vbCrLf & _
                 "</tr>" & vbCrLf
                 
   htmlSpec = "<tr class=spec>" & vbCrLf & _
              "   <td align='center'>#KCODE#</td>" & vbCrLf & _
              "   <td>#KDESC#</td>" & vbCrLf & _
              "   <td align='center'>#MSG#</td>" & vbCrLf & _
              "</tr>" & vbCrLf
              
   strSQL = "SELECT EDI_LTS_Index, EDI_Msg_Type, EDI_OrgCode " & _
            "FROM EDI_Local_Trader_Settings"
            
   ltRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   strSQL = "SELECT DISTINCT " & _
               "Case " & _
                  "When EDI_Local_Key1 is null Then 'No key Set' " & _
                  "Else EDI_Local_Key1 " & _
               "End As EDI_Local_Key1, " & _
               "EDI_LTS_Index " & _
            "FROM EDI_Local_Trader_Settings " & _
               "LEFT JOIN EDI_Matching em " & _
                  "INNER JOIN EDI_Recipient_Individuals er " & _
                  "ON em.Individual_Index = er.Individual_Index " & _
               "ON (EDI_LTS_Index = EDI_Local_Key2 AND EDI_Org_NatCode = '" & Left(NodeId.Text, 6) & "') " & _
            "ORDER BY EDI_LTS_Index"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly

   strSQL = "SELECT Case " & _
                  "When EDI_Specialty is Null Then 'None set up' " & _
                  "Else EDI_Specialty " & _
               "End as EDI_Specialty, " & _
               "case " & _
                  "When EDI_Korner_Code is Null Then 'N/A' " & _
                  "Else Convert(varchar(3),EDI_Korner_Code) " & _
               "End as EDI_Korner_Code, " & _
               "Case " & _
                  "When EDI_Msg_Format is Null Then 'Not sent' " & _
                  "Else EDI_Msg_Format " & _
               "End as EDI_Msg_Format, " & _
               "lt.EDI_LTS_Index " & _
            "FROM EDI_Local_Trader_Settings lt " & _
               "LEFT JOIN EDI_Loc_Specialties es " & _
               "ON (lt.EDI_LTS_Index = es.EDI_LTS_Index AND es.EDI_Nat_Code = '" & Left(NodeId.Text, 6) & "') " & _
            "ORDER BY lt.EDI_LTS_Index"
            
   RS2.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   Do Until ltRS.EOF
      sBuf.Append tableHdr & Replace(htmlLST, "#LTS#", ltRS!EDI_Msg_Type & " - " & ltRS!EDI_OrgCode)
      sBuf.Append htmlKeyHdr
      RS.Filter = "EDI_LTS_Index = " & ltRS!EDI_LTS_Index
      Do Until RS.EOF
         buf = Replace(htmlKey1, "#KEY1#", RS!EDI_Local_Key1)
         strSQL = "SELECT COUNT(*) " & _
                  "FROM EDI_Matching " & _
                  "WHERE EDI_Local_Key1 = '" & RS!EDI_Local_Key1 & "' " & _
                     "AND EDI_Local_Key2 = " & ltRS!EDI_LTS_Index
         RS3.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         sBuf.Append Replace(buf, "#KEY3#", RS3(0))
         RS3.Close
         RS.MoveNext
      Loop
      RS.Filter = ""
      RS.MoveFirst
      
      sBuf.Append htmlSpecHdr
      RS2.Filter = "EDI_LTS_Index = '" & ltRS!EDI_LTS_Index & "'"
      Do Until RS2.EOF
         buf = Replace(htmlSpec, "#KCODE#", RS2!EDI_Korner_Code)
         buf = Replace(buf, "#KDESC#", RS2!EDI_Specialty)
         sBuf.Append Replace(buf, "#MSG#", Mid(RS2!EDI_Msg_Format, InStr(RS2!EDI_Msg_Format, ",") + 1))
         RS2.MoveNext
      Loop
      RS2.Filter = ""
      RS2.MoveFirst
      
      sBuf.Append "</table>" & vbCrLf
      ltRS.MoveNext
   Loop
   
   With wb
      .LocationTitle = ""
      .LocationToolTip = ""
      .Style = "practice"
      .NavigateTo sBuf.ActualValue
   End With
   
   ltRS.Close
   RS.Close
   RS2.Close
   Set ltRS = Nothing
   Set RS = Nothing
   Set RS2 = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.ShowPractice"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub


Public Sub ShowReport(NodeId As Node)
   On Error GoTo procEH
   Dim natCode As String
   Dim repDateTime As String
   Dim repType As String
   Dim RepId As Long
   Dim blnComment As Boolean
   Dim K As Integer
   Dim maxRange As Integer
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim iceCmd As New ADODB.Command
   Dim iBuf As New StringBuffer
   Dim rBuf As New StringBuffer
   Dim resDisp As Long
   Dim strComment As String
   Dim sampData As String
   Dim repHeader As String
   Dim repBuf As String
   Dim repTitle As String
   Dim hpData As New clsHealthParties
   
   wb.NavigationBar = 2
   wb.BrowserToolTip = ""
   wb.LocationToolTip = "Double click to View file"
   
   vData = objTV.ReadNodeData(NodeId.Parent)
   wb.Tag = vData(1)
   
   vData = objTV.ReadNodeData(NodeId)
   RepId = vData(4)
   curRepId = RepId
   
   Set collRTF = Nothing
   Set collRTF = New Collection
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_ReportLocalTrader"
      .Parameters.Append .CreateParameter("ReportId", adInteger, adParamInput, , RepId)
      Set RS = .Execute
      
      repLTS_Index = RS!EDI_LTS_Index
      intUseRCIndex = RS!UseLabReadCodes
      
      RS.Close
   End With
         
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Patient"
      Set RS = .Execute
      natCode = RS!EDI_NatCode & ""
      repDateTime = RS!DateTime_Of_Report
      repType = RS!Service_Report_Type
      blnComment = RS!Comment_Marker
   End With
   
   repHeader = Replace(repHTMLHeader, "#Surname#", Trim(RS!Surname))
   repHeader = Replace(repHeader, "#Forename#", Trim(RS!Forename))
   repHeader = Replace(repHeader, "#DOB#", RS!Date_Of_Birth & "")
   repHeader = Replace(repHeader, "#Sex#", IIf(RS!Sex = 1, "Male", IIf(RS!Sex = 2, "Female", "N/S")))
   repHeader = Replace(repHeader, "#NHSNo#", RS!New_Nhs_No & "")
   repHeader = Replace(repHeader, "#HospNo#", RS!Hospital_Number & "")
   repHeader = Replace(repHeader, "#ServiceId#", RS!Service_Report_Id)
   repHeader = Replace(repHeader, "#RepDate#", repDateTime)
   repHeader = Replace(repHeader, "#RepSpec#", repType)
'   repHeader = Replace(repHeader, "#RepStatus#", RS!Status)
   
   repTitle = Trim(RS!Surname) & " " & Trim(RS!Forename)
'   wb.InfoTitle = Trim(RS!Surname) & " " & Trim(RS!Forename)
   wb.InfoDescription = RS!Service_Report_Id & " (Index: " & RepId & ")"
   wb.LocationTitle = NodeId.Parent.Text
'   sBuf.Clear
'   sBuf.Append vbCrLf & Trim(RS!Surname) & ", " & Trim(RS!Forename) & vbTab & RS!Date_Of_Birth & "  " & RS!Sex
'   sBuf.Append vbCrLf & "NHS No.: " & vbTab & RS!New_NHS_No
'   sBuf.Append vbCrLf & "Hosp. No.: " & vbTab & RS!Hospital_Number
'   sBuf.Append vbTab & "Service ID:  " & RS!Service_Report_Id
   
   hpData.IndividualIndex = -1
   hpData.DownloadRequest = RS!GP_Download
   RS.Close
   
   hpData.Read RepId, True
   
   
'   With iceCmd
'      .CommandText = "ICELABCOMM_Report_HealthParties"
'      Set RS = .Execute
'   End With
   
'   RS.Find "EDI_HP_Type = '902'"
   
   repHeader = Replace(repHeader, "#Clinician#", hpData.HP902Name)
   repHeader = Replace(repHeader, "#NatCode#", natCode)
   repHeader = Replace(repHeader, "#Specialty#", hpData.ClinicianSpeciality)
   repHeader = Replace(repHeader, "#Recp#", hpData.HP902Name & " (" & hpData.HP902Code & ")")    'IIf(IsNull(RS!EDI_OP_Name), "Not Matched", RS!EDI_OP_Name))
   
'   RS.Find "EDI_HP_Type='906'"
   repHeader = Replace(repHeader, "#Reqst#", hpData.HP906Name & " (" & hpData.HP906Code & ")") 'IIf(IsNull(RS!EDI_OP_Name), "Not Matched", RS!EDI_OP_Name))
   
'   sBuf.Append vbCrLf & "Clinician/Specialty: " & Trim(RS!Clinician_Surname) & "  " & RS!Clinician_Speciality_Code
'   sBuf.Append vbCrLf & "Destination: " & natCode
'   sBuf.Append vbCrLf & "Report Date: " & repDateTime
'   sBuf.Append vbTab & "Type: " & repType & vbCrLf
'   RS.Close
   
   strSQL = "SELECT Colour_Code " & _
            "FROM Service_Tubes_Colours " & _
            "WHERE Colour_Name IN (" & _
               "SELECT Report_Colour " & _
               "FROM Service_Reports_Colours " & _
               "WHERE Report_Type IN (" & _
                  "SELECT Specialty_Code " & _
                  "FROM Specialty " & _
                  "WHERE Specialty LIKE '" & repType & "'))"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
   If RS.RecordCount > 0 Then
      wb.InfoBGColour = ConvertToRGB(RS!Colour_Code)
   Else
      wb.InfoBGColour = "#FFFFCC"
'      repHeader = Replace(repHeader, "##FFFFCC", ConvertToRGB(RS!Colour_Code))
   End If
      
   RS.Close
      
   sBuf.Clear
   sBuf.Append repHeader
   
   If blnComment Then
      sBuf.Append repHTMLHR & repHTMLTableStart
      With iceCmd
         .CommandText = "ICELABCOMM_Report_Comments"
         Set RS = .Execute
      End With
      
      Do Until RS.EOF
         strComment = repHTMLComment
         strComment = Replace(strComment, "#Comment#", RS!Service_Report_Comment)
'         sBuf.Append vbCrLf & vbTab & "{" & RS!Service_Report_Comment & "}"
         sBuf.Append strComment
         RS.MoveNext
      Loop
      
      RS.Close
      sBuf.Append repHTMLTableEnd
   End If
   
   iceCmd.CommandText = "ICELABCOMM_Report_Sample"
   Set RS = iceCmd.Execute
   
   rBuf.Clear
   
'   For i = 0 To 5
'      tStop(i) = 0
'   Next i
   
   sampCnt = 1
   
   Do Until RS.EOF
      sampText = RS!Sample_Text
      
      sampData = Replace(sampHTML, "#SampRef#", sampCnt)
      sampData = Replace(sampData, "#SampDets#", RS!Sample_Text & " (" & RS!Sample_Code & ")")
      sampData = Replace(sampData, "#ColDate#", RS!Collection_DateTime & "")
      sampData = Replace(sampData, "#RecdDate#", RS!Collection_DateTimeReceived & "")
      sBuf.Append sampData
      
'      sBuf.Append vbCrLf & vbCrLf & "Sample Key: #" & sampCnt & vbCrLf
'      sBuf.Append "Details: " & RS!Sample_Text & " (" & RS!Sample_Code & ")" & vbCrLf
'      sBuf.Append vbTab & "Collected: " & RS!Collection_DateTime
'      sBuf.Append vbTab & "Received: " & RS!Collection_DateTimeReceived & vbCrLf
      ShowInvestigation RepId, rBuf, RS!Sample_Index
'      rBuf.Append invHTMLEnd
      RS.MoveNext
      sampCnt = sampCnt + 1
   Loop
   
   RS.Close
   
   iceCmd.CommandText = "ICELABCOMM_LocalTrader"
   Set RS = iceCmd.Execute
   
   wb.InfoTitle = repTitle & " (" & repStatus & ")"
   
   wb.ToolTipText = "Datastream: " & RS!EDI_OrgCode & " - " & RS!EDI_Msg_Type
'   rBuf.Append repHTMLTrailer
   RS.Close
   
'   rBuf.Append flagStr
   resDisp = sBuf.Length + 2
   
   Open App.Path & "\bptext.html" For Output As #1
   Print #1, sBuf.value & rBuf.value
   Close #1
   
   repBuf = sBuf.ActualValue & rBuf.ActualValue
   With wb
      .Script = "report"
      .Style = "report"
      .NavigateTo repBuf, False
   End With
'   wb.NavigateTo App.Path & "\bptext.html", True

'   With rtfData
'      '.ToolTipText = "'N/A' = Read Code not applicable 'NONE' = No read code for test"
''      .Visible = False
'      .TextRTF = sBuf.ActualValue & vbCrLf & rBuf.ActualValue & " "
'      If collRTF.Count > 0 Then
''         For i = 0 To 4
''            Debug.Print "Tab " & i & " = " & tStop(i)
''         Next i
'
'         For i = 1 To collRTF.Count&
'            Set objRTF = collRTF.item(i)
'            .SelStart = objRTF.TextPosition + resDisp
'            .SelLength = objRTF.TextLength
'
'            Select Case objRTF.TextType
'               Case "INV"
'                  If objRTF.TextStyle = "TAB" Then
'                     .SelTabCount = 4
'                     .SelTabs(0) = invStop(0) + 1
'                     Debug.Print .SelText
'                     For K = 1 To 3
'                        Debug.Print invStop(K)
'                        .SelTabs(K) = .SelTabs(K - 1) + invStop(K - 1) + 1
'                     Next K
'                  End If
'
''                  .SelTabs(1) = 1000
'
'               Case "RES"
'                  If objRTF.TextStyle = "TAB" Then
'                     .SelTabCount = 6
'                     .SelTabs(0) = 1
'                     For K = 1 To 5
'                        .SelTabs(K) = .SelTabs(K - 1) + tStop(K - 1) + 2
'                     Next K
'                  End If
'
'               Case "COLOUR"
'                  .SelColor = objRTF.TextColour
'
'            End Select
'
'            If objRTF.TextStyle = "BOLD" Then
'               .SelBold = True
'            End If
'
'            If objRTF.TextStyle = "ULINE" Then
'               .SelUnderline = True
'            End If
'
'            If objRTF.TextStyle = "STRIKE" Then
'               .SelStrikeThru = True
'            End If
'         Next i
'
'         .SelStart = Len(.Text) - 1
'         .SelLength = 1
'         .SelColor = vbBlack
'      End If
      Set objRTF = Nothing
      Set collRTF = Nothing
      Set collRTF = New Collection
'      .Visible = True
'   End With
   
   Set RS = Nothing
   cmdRCKeys.Caption = "Show Read Code Key"
   cmdRCKeys.Visible = True
   Exit Sub
    
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.DisplayReport"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Function ShowInvestigation(RepId As Long, _
                                   ByRef sBuf As StringBuffer, _
                                   sampId As Long) As String
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim iceCmd As New ADODB.Command
   Dim objRTF As New rtfSelection
   Dim vData As Variant
   Dim vFields As Variant
   Dim iposn As Integer
   Dim txtPos As Long
   Dim txtStart As Long
   Dim invStr(4) As String
   Dim thisInv As String
   Dim rBuf As New StringBuffer
   Dim i As Integer
   Dim strComment As String
   Dim invData As String
   Dim rcStr As String
   Dim invCls As String
'   Dim resTable As String
   
   txtPos = sBuf.Length
   txtStart = txtPos
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      If intUseRCIndex > 0 Then
         .CommandText = "ICELABCOMM_Report_Invest_By_Index"
      Else
         .CommandText = "ICELABCOMM_Report_Invest_By_Code"
      End If
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RepId)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repLTS_Index) 'frmFilter.CurrentLTSIndex)
      .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , sampId)
      .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
      Set RS = .Execute
   End With
   
   If iceCmd.Parameters("Abnormal").value Then
      sBuf.Append outOfRange
'      Set objRTF = New rtfSelection
'      objRTF.TextType = "COLOUR"
'      objRTF.TextPosition = txtPos + 2
'      objRTF.TextLength = 29
'      objRTF.TextColour = BPRED
'      sBuf.Append vbCrLf & "CONTAINS OUT OF RANGE RESULTS" & vbCrLf
'      collRTF.Add objRTF
'      Set objRTF = Nothing
'      txtStart = txtPos + 33
   End If
   
   With iceCmd
      .Parameters.Delete 3
      .Parameters.Delete 2
      .Parameters.Delete 1
   End With
   
'   If IsNull(RS!EDI_RC_Index) Then
'      vFields = Array(9, 10, 12, 13, 18, 19, 20, 21, 22, 23, 8)
'   Else
'      vFields = Array(14, 15, 16, 17, 24, 25, 26, 27, 28, 29, 8)
'   End If
'
'   iposn = 0
'   vData = RS.GetRows(-1, adBookmarkFirst, vFields)
   RS.MoveFirst
   
   Do Until RS.EOF
      invStr(0) = "#" & sampCnt
      invStr(1) = " "
      rcStr = ""
      
      If IsNull(RS!Result_Index) Or RS!Result_Recs > 1 Then
         If IsNull(RS!Read_V2RC) Then
            rcStr = "&nbsp;NONE"
         Else
            rcStr = "&nbsp;" & RS!Read_V2RC
         End If
               
         If RS!Read_Status = "D" Then
'         If vData(8, iposn) = "D" Then
            invCls = "class=" & Chr(34) & "deleted" & Chr(34) & " " '   Flagged as deleted - Blue + Line through
         ElseIf RS!EDI_OP_Suppress Then
            invCls = "class=" & Chr(34) & "suppressed" & Chr(34) & " " ' Suppressed - Red + Line through
         ElseIf RS!EDI_Op_Active = False Then
            invCls = "class=" & Chr(34) & "inact" & Chr(34) & " " '   Inactive - Red
         ElseIf RS!Read_Battery = "F" Then
            invCls = "class=" & Chr(34) & "inapp" & Chr(34) & " " ' Inapplicable - Blue
         Else
            invCls = "class=" & Chr(34) & "ir" & Chr(34) & " "
         End If
         
         
         invData = Replace(invHTML, "#InvRC#", rcStr)
         invData = Replace(invData, "#IRCClass#", invCls)
'         If IsNull(vData(4, iposn)) Then
'            invData = Replace(invHTML, "#InvRC#", ">[NONE ]")
''            invStr(2) = "[NONE ]"
'         Else
'            invData = Replace(invHTML, "#InvRC#", ">[" & vData(4, iposn) & "]")
''            invStr(2) = "[" & vData(4, iposn) & "]"
'         End If
               
      Else
         invData = Replace(invHTML, "#InvRC#", "&nbsp;N/A")
'         invStr(2) = "[ N/A ]"
      End If
      invData = Replace(invData, "#Inv#", Trim(RS!Investigation_Requested) & " (" & Trim(RS!investigation_Code) & ")")
      invData = Replace(invData, "#ITId#", sampCnt & "_" & RS.AbsolutePosition)
      
      sBuf.Append invData
      
      If RS!Comment_Marker Then
         sBuf.Append repHTMLTableStart
         With iceCmd
            .CommandText = "ICELABCOMM_Report_InvestComments"
            .Parameters(0).value = RS!Investigation_Index
            Set RS2 = .Execute
         End With
         
         Do Until RS2.EOF
            strComment = repHTMLComment
            strComment = Replace(strComment, "#Comment#", RS2!Service_Investigation_Comment)
            sBuf.Append strComment
            RS2.MoveNext
         Loop
         
         RS2.Close
         sBuf.Append repHTMLTableEnd
      End If
            
'      For i = 0 To 3
'         invStr(i) = ""
'      Next i
'      thisInv = ""
      
'      resTable = Replace(htmlpindent, "#RTId#", RS.AbsolutePosition)
'      resTable = Replace(resTable, "#ClassId#", "")
'      sBuf.Append repHTMLPIndent
'      resTable = Replace(repHTMLTableStart, "border=" & Chr(34) & "0", "border=" & Chr(34) & "1")
'      resTable = Replace(resTable, "100%", "98%")
      resTable = Replace(resHTMLTableStart, "#RTId#", sampCnt & "_" & RS.AbsolutePosition)
      sBuf.Append repHTMLPIndent & resTable
      
      ShowResult sBuf, RS!Investigation_Index
      sBuf.Append repHTMLTableEnd
      sBuf.Append repHTMLPEnd
      
      RS.MoveNext
      iposn = RS.AbsolutePosition - 1
'      sBuf.Append vbCrLf
   Loop
   RS.Close
   
   Set RS2 = Nothing
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.ShowInvestigation2"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Sub ShowLog(NodeId As Node)
   On Error GoTo procEH
   Dim LogView As New LogToHTML
   Dim logData As String
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "SELECT Connection_LogDirs " & _
            "FROM Connections " & _
               "Inner Join Connect_Modules " & _
               "ON Connections.Connection_Name = Connect_Modules.Connection_Name " & _
            "WHERE Module_Name = 'IceMsg.exe' " & _
               "AND Connection_InFlightMapping = 'EDIRECIPLIST'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   logData = fs.BuildPath(RS!Connection_LogDirs, "IceMsg" & Format(NodeId.Text, "_yyyymmdd") & ".xml")
   RS.Close
   Set RS = Nothing
'   LogView.ReadLogFile logData
   
   Me.MousePointer = vbHourglass
   With wb
      .Tag = logData
      .Style = "log"
      .NavigationBar = 0
      .LocationTitle = logData
      .BrowserToolTip = "IceMsg Log for " & NodeId.Text
      .NavigateTo logData, True
   End With
   Me.MousePointer = vbNormal
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.ShowLog"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Function ShowResult(ByRef sBuf As StringBuffer, _
                            InvId As Long) As Long
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim resStr(6) As String
   Dim resRange As String
   Dim thisResult As String
   Dim txtPos As Long
   Dim txtStart As Long
   Dim objRTF As New rtfSelection
   Dim vData As Variant
   Dim vFields As Variant
   Dim rPosn As Integer
   Dim i As Integer
   Dim strFlag As String
   Dim blnRemoved As Boolean
   Dim strComment As String
   Dim resData As String
   Dim resRC As String
   Dim endTag As String
'   Dim resTable As String
   
   txtPos = sBuf.Length
   txtStart = txtPos
   
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      If intUseRCIndex > 0 Then
         .CommandText = "ICELABCOMM_Report_Results_By_Index"
      Else
         .CommandText = "ICELABCOMM_Report_Results_By_Code"
      End If
      .Parameters.Append .CreateParameter("InvId", adInteger, adParamInput, , InvId)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repLTS_Index) 'frmFilter.CurrentLTSIndex)
      Set RS = .Execute
      .Parameters.Delete ("LTSIndex")
   End With
   
'   If IsNull(RS!EDI_RC_Index) Then
'      vFields = Array(16, 17, 19, 20, 21, 28, 29, 30, 31, 32, 33, 34, 5)
'   Else
'      vFields = Array(22, 23, 24, 25, 26, 35, 36, 37, 38, 39, 40, 41, 5)
'   End If
'
'   vData = RS.GetRows(-1, adBookmarkFirst, vFields)
   RS.MoveFirst
'   rPosn = 0
   
   Do Until RS.EOF
      If RS!status = "FR" Or RS!status = "" Then
         repStatus = "Final"
      ElseIf RS!status = "SR" Then
         repStatus = "Supplementary"
      Else
         repStatus = "Interim"
      End If
      
      blnRemoved = False
      If RS!Abnormal_Flag Then
         resData = Replace(resHTML, "#ClassId#", "oor")
         resData = Replace(resData, "#ResValue#", "*" & RS!Result)
      Else
         resData = Replace(resHTML, "#ClassId#", "ir")
         resData = Replace(resData, "#ResValue#", RS!Result)
      End If
      
      If RS!Read_Ratio <> "T" Then
         If IsNumeric(RS!Result) Then
            If RS!UOM_Code = "" And _
               RS!EDI_OP_UOM = "" Then
               resStr(0) = "!"
               blnRemoved = True
            End If
         End If
      End If
   
      endTag = ""
      
      If Not IsNull(RS!Read_V2RC) Then
         If RS!EDI_Op_Active = False Then
            resRC = "class=" & Chr(34) & "inact" & Chr(34) & " " '  Inactive - Red
         ElseIf RS!EDI_OP_Suppress Then
            resRC = "class=" & Chr(34) & "suppressed" & Chr(34) & " " '  Suppressed - Red + line-through
         
         ElseIf RS!Read_Test = "F" Then
            resRC = "class=" & Chr(34) & "inapp" & Chr(34) & " " '  Inapplicable - Blue
         ElseIf RS!Read_Status = "D" Then
            resRC = "class=" & Chr(34) & "deleted" & Chr(34) & " " '  Flagged as deleted - Blue + line-through
         ElseIf blnRemoved Then
            resRC = "class=" & Chr(34) & "removed" & Chr(34) & " " '  No UOM - Blue + underline
         Else
            resRC = "class=" & Chr(34) & "ir" & Chr(34) & " "
         End If
         
         resData = Replace(resData, "#ResRC#", RS!Read_V2RC)
         resData = Replace(resData, "#RCClass#", resRC)
      ElseIf RS!EDI_RC_Index = 0 Then
         resData = Replace(resData, "#ResRC#", "&nbsp;XXXX")
      Else
         resRC = "class=" & Chr(34) & "suppressed" & Chr(34) & " " '  Suppressed - Red + line-through
         resData = Replace(resData, "#ResRC#", "&nbsp;NONE")
         resData = Replace(resData, "#RCClass#", resRC)
      End If
      
      resData = Replace(resData, "#ResRubric#", Trim(RS!Result_Rubric))
      resData = Replace(resData, "#UOM#", RS!UOM_Code)
      
      If Trim(RS!Lower_Range & "") = "" And Trim(RS!Upper_Range & "") = "" Then
         resRange = ""
      ElseIf Trim(RS!Lower_Range) = "" Then
         resRange = "<" & Trim(RS!Upper_Range)
      ElseIf Trim(RS!Upper_Range) = "" Then
         resRange = ">" & Trim(RS!Lower_Range)
      Else
         resRange = Trim(RS!Lower_Range) & " to " & Trim(RS!Upper_Range)
      End If
      
      resData = Replace(resData, "#Range#", resRange)
      resData = Replace(resData, "#RTId#", sampCnt & "_" & RS.AbsolutePosition)
      
      sBuf.Append resData
      
      If RS!Comment_Marker Then
         sBuf.Append repHTMLTableEnd & repHTMLTableStart
         With iceCmd
            .CommandText = "ICELABCOMM_Report_ResultComment"
            .Parameters(0).value = RS!Result_Index
            Set RS2 = .Execute
         End With
         
         Do Until RS2.EOF
            strComment = repHTMLComment
            strComment = Replace(strComment, "#Comment#", Replace(Trim(RS2!Service_Result_Comment), "  ", "&nbsp;"))
            sBuf.Append strComment
            RS2.MoveNext
         Loop
         
         RS2.Close
         sBuf.Append repHTMLTableEnd & resTable
      End If
      
      RS.MoveNext
      rPosn = RS.AbsolutePosition - 1
   Loop
   RS.Close
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.ShowResult"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Sub RecordTabStop(ByRef resBuf As String, TabNo As Integer)
   If (Me.TextWidth(resBuf) - (TabNo * Me.TextWidth(Chr(255)))) > tStop(TabNo) Then
      tStop(TabNo) = Me.TextWidth(resBuf)
   End If
   resBuf = resBuf & Chr(255)
   Debug.Print resBuf & " width(" & TabNo; " ) = " & tStop(TabNo)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload frmFilter
   Set dtsAPI = Nothing
End Sub

Private Sub mnuLTS_Click(Index As Integer)
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   If mnuLTS(Index).Caption = "All datastreams" Then
      LTS_Index = 0
      LTS_DataStream = "All datastreams"
      LTS_OrgCode = "All"
   Else
      strSQL = "SELECT EDI_LTS_Index, EDI_Msg_Type, EDI_OrgCode " & _
               "FROM EDI_Local_Trader_Settings " & _
               "WHERE Organisation = '" & frmMain.cboTrust.Text & "' " & _
                  "AND EDI_OrgCode = '" & Trim(Left(mnuLTS(Index).Caption, InStr(1, mnuLTS(Index).Caption, "-") - 2)) & "' " & _
                  "AND EDI_Msg_Type = '" & Trim(Mid(mnuLTS(Index).Caption, InStr(1, mnuLTS(Index).Caption, "-") + 1)) & "'"
      RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      LTS_DataStream = RS!EDI_Msg_Type
      LTS_Index = RS!EDI_LTS_Index
      LTS_OrgCode = RS!EDI_OrgCode
      RS.Close
   End If
   
   Set RS = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   
   Me.MousePointer = vbHourglass
   If SSTab1.Tab = 0 Then
      With frmFilter
         blnRCInactive = (.optOKLevel(0).value = 1)
         blnRCInapplicable = (.optOKLevel(1).value = 1)
         blnRCDelete = (.optOKLevel(2).value = 1)
         blnRCNoUOM = (.optOKLevel(3).value = 1)
         blnRCSuppBatt = (.optOKLevel(4).value = 1)
         blnRCSuppTest = (.optOKLevel(5).value = 1)
      End With
   Else
      blnRCInactive = False
      blnRCInapplicable = True
      blnRCDelete = True
      blnRCNoUOM = True
      blnRCSuppBatt = True
      blnRCSuppTest = True
   End If
   tvMsg(SSTab1.Tab).Visible = False
   GetDates
   Me.MousePointer = vbNormal
   tvMsg(SSTab1.Tab).Visible = True
End Sub

Private Sub tvMsg_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)
   If Node.Tag = "" Then
      NodeExpand Node
      Node.Tag = "X"
   End If
End Sub

Private Sub tvMsg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim tNode As Node
   Dim i As Integer
   Dim nText As String
   
   If Button = vbRightButton Then
      If mnuLTS.UBound > 0 Then
         PopupMenu mnuLocalTrader
         
         frmNewLogView.MousePointer = vbHourglass
         DoEvents
         GetDates
         
         lblLTS.Caption = "Datastream: " & LTS_OrgCode & " (" & LTS_DataStream & ") only"
            
         lblLTS.Visible = (LTS_Index > 0)
         
         frmNewLogView.MousePointer = vbNormal
      End If
   End If
End Sub

Private Sub tvMsg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not NodeToClear Is Nothing Then
      NodeToClear.Checked = False
      Set NodeToClear = Nothing
   End If
End Sub

Private Sub tvMsg_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)
   nodeCheck Node
End Sub

Private Sub tvMsg_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
   NodeClick Node
End Sub

Private Sub nodeFiles(NodeId As Node, parentChecked As Boolean)
   On Error GoTo procEH
   Dim cNode As Node
   Dim iceCmd As New ADODB.Command
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim mStat As Long
   
   If NodeId.Children > 0 Then
      vData = objTV.ReadNodeData(NodeId)
      If vData(2) <> "GetReports" Then
         Set cNode = NodeId.child
         nodeFiles cNode, NodeId.Checked
         
         Do Until cNode.Next Is Nothing
            Set cNode = cNode.Next
            nodeFiles cNode, cNode.Checked
         Loop
      Else
         If NodeId.Checked Then
            rCtrl.RequeueItem = vData(0)
            rTot = rTot + 1
         End If
      End If
   Else
      If NodeId.Checked Then
         vData = objTV.ReadNodeData(NodeId)
         
         mStat = frmFilter.MessageFlag(SSTab1.Tab)
         
         If SSTab1.Tab = 0 Then
            ' Do we include All messages or just not acked? The 'And' removes unwanted flags
            If (mStat And IS_NO_ACK) = IS_NO_ACK Then
               mStat = (MS_MSGOK Or MS_AWAIT_ACK Or MS_ACK_RECEIVED)
            Else
               mStat = MS_MSGOK
            End If
         End If
         
         strSQL = "SELECT DISTINCT Service_ImpExp_Id " & _
                  "FROM ImpExp_View "
            
         If SSTab1.Tab = 0 Then
            strSQL = strSQL & _
                  "WHERE Convert(varchar(10), Date_Added, 103) = '" & vData(0) & "' " & _
                     "AND Ack_Status & 0x" & Hex(mStat) & " = 0 "
         Else
            strSQL = strSQL & _
                  "WHERE Convert(varchar(10), Last_Comment_Added, 103) = '" & vData(0) & "' " & _
                     "AND Header_Status & 0x" & Hex(mStat) & " <> 0 "
         End If
         
         Select Case vData(2)
            Case "GetPractices"
               If SSTab1.Tab = 0 Then
                  strSQL = strSQL & _
                     "AND Header_Status < 0 "
               End If
               
            Case "GetFiles"
               If SSTab1.Tab = 0 Then
                  If InStr(1, NodeId.Parent.Text, " ") = 0 Then
                     strSQL = strSQL & _
                              "AND EDI_NatCode Is Null"
                  Else
                     strSQL = strSQL & _
                              "AND EDI_NatCode =  '" & Left(NodeId.Parent.Text, (InStr(1, NodeId.Parent.Text, " ") - 1)) & "' "
                  End If
               End If
            
            Case Else
               strSQL = ""
               
         End Select
            
         If strSQL <> "" Then
            RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
            Do Until RS.EOF
               rCtrl.RequeueItem = RS!Service_ImpExp_Id
               RS.MoveNext
               rTot = rTot + 1
            Loop
         End If
      End If
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
       Stop
      Resume
   End If
   eClass.FurtherInfo = "frmNewLogView.NodeFiles"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub nodeRequeue(NodeId As Node, Optional parentChecked As Boolean = False)
   On Error GoTo procEH
   Dim cNode As Node
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim mStat As Long
   
   If NodeId.Children > 0 Then
      Set cNode = NodeId.child
      nodeRequeue cNode, NodeId.Checked
      
      Do Until cNode.Next Is Nothing
         Set cNode = cNode.Next
         nodeRequeue cNode, cNode.Checked
      Loop
   Else
      If NodeId.Checked Then
         vData = objTV.ReadNodeData(NodeId)
         
         mStat = frmFilter.MessageFlag(SSTab1.Tab)
         
         If SSTab1.Tab = 0 Then
            ' Do we include All messages or just not acked? The 'And' removes unwanted flags
            If (mStat And IS_NO_ACK) = IS_NO_ACK Then
               mStat = (MS_MSGOK Or MS_AWAIT_ACK Or MS_ACK_RECEIVED)
            Else
               mStat = MS_MSGOK
            End If
         End If
            
         strSQL = "SELECT DISTINCT Service_ImpExp_Message_Id " & _
                  "FROM Service_ImpExp_Messages SM " & _
                     "INNER JOIN ImpExp_View iv " & _
                     "ON SM.Service_ImpExp_Id = iv.Service_ImpExp_Id "
         
         If SSTab1.Tab = 0 Then
            strSQL = strSQL & _
                  "WHERE Convert(varchar(10), iv.Date_Added, 103) = '" & vData(0) & "' " & _
                     "AND Ack_Status & 0x" & Hex(mStat) & " = 0 "
         Else
            strSQL = strSQL & _
                  "WHERE Convert(varchar(10), iv.Last_Comment_Added, 103) = '" & vData(0) & "' " & _
                     "AND Header_Status & 0x" & Hex(mStat) & " <> 0 "
         End If
         
         Select Case vData(2)
            Case "GetPractices"
               If SSTab1.Tab = 0 Then
                  strSQL = strSQL & _
                     "AND Header_Status < 0 "
               End If
               
            Case "GetFiles"
               If SSTab1.Tab = 0 Then
                  If InStr(1, NodeId.Parent.Text, " ") = 0 Then
                     strSQL = strSQL & _
                              "AND iv.EDI_NatCode Is Null"
                  Else
                     strSQL = strSQL & _
                              "AND iv.EDI_NatCode =  '" & Left(NodeId.Parent.Text, (InStr(1, NodeId.Parent.Text, " ") - 1)) & "' "
                  End If
               End If
            
            Case "GetReports"
               strSQL = "SELECT DISTINCT Service_ImpExp_Message_Id " & _
                        "FROM Service_ImpExp_Messages SM " & _
                           "INNER JOIN ImpExp_View iv " & _
                           "ON SM.Service_ImpExp_Id = iv.Service_ImpExp_Id " & _
                        "WHERE iv.Service_ImpExp_Id = " & vData(0)
            
            Case "ShowReport"
               strSQL = ""
               rCtrl.RequeueItem = vData(1)
               rTot = rTot + 1
         
            Case Else
               strSQL = ""
               
         End Select
         
         If strSQL <> "" Then
            RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
            Do Until RS.EOF
               rCtrl.RequeueItem = RS!Service_ImpExp_Message_ID
               RS.MoveNext
               rTot = rTot + 1
            Loop
         End If
      End If
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmNewLogView.nodeRequeue2"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub nodeCheck(NodeId As Node)
   Dim cNode As Node
   
   Set NodeToClear = Nothing
   If NodeId.Children > 0 Then
      Set cNode = NodeId.child
      cNode.Checked = NodeId.Checked
      nodeCheck cNode
      
      Do Until cNode.Next Is Nothing
         Set cNode = cNode.Next
         cNode.Checked = NodeId.Checked
         nodeCheck cNode
      Loop
   Else
      If NodeId.Parent Is Nothing Then
         Set NodeToClear = NodeId
      Else
         NodeId.Checked = NodeId.Parent.Checked
      End If
   End If
End Sub



