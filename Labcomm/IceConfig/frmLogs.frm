VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogs 
   BackColor       =   &H8000000C&
   Caption         =   "Log Search Options"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11880
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      Caption         =   "Show..."
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   10440
      TabIndex        =   31
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optShow 
         BackColor       =   &H8000000C&
         Caption         =   "Files"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optShow 
         BackColor       =   &H8000000C&
         Caption         =   "Reports"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraResults 
      BackColor       =   &H8000000C&
      Caption         =   "Search Results"
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   240
      TabIndex        =   25
      Top             =   3720
      Width           =   11535
      Begin MSComctlLib.ProgressBar pbReq 
         Height          =   375
         Left            =   600
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear All"
         Height          =   255
         Left            =   9360
         TabIndex        =   19
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   7920
         TabIndex        =   18
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdRequeue 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Requeue"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3480
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DGResults 
         Height          =   3015
         Left            =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.Label lblReqInfo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Requeue in progress"
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
         Height          =   375
         Left            =   600
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   10455
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   10440
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      Caption         =   "Standard Options"
      ForeColor       =   &H8000000E&
      Height          =   3495
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   3975
      Begin VB.ComboBox comboLTI 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "frmLogs.frx":0000
         Top             =   1920
         Width           =   3735
      End
      Begin VB.ComboBox comboSvcType 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPStart 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   51773441
         CurrentDate     =   37572
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51773441
         CurrentDate     =   37572
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000C&
         Caption         =   "Datastream"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000C&
         Caption         =   "Up to and including:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Search From:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Caption         =   "Report Type"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Advanced"
      ForeColor       =   &H8000000E&
      Height          =   3495
      Left            =   4440
      TabIndex        =   20
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox cboKorner 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2760
         Width           =   2460
      End
      Begin VB.CheckBox chkSearchOpt 
         BackColor       =   &H8000000C&
         Caption         =   "Discipline"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtRepId 
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox comboClin 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   3495
      End
      Begin VB.CheckBox chkSearchOpt 
         BackColor       =   &H8000000C&
         Caption         =   "Interchange Nos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox chkSearchOpt 
         BackColor       =   &H8000000C&
         Caption         =   "Report Id"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkSearchOpt 
         BackColor       =   &H8000000C&
         Caption         =   "Patient Details"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkSearchOpt 
         BackColor       =   &H8000000C&
         Caption         =   "Clinician"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkSearchOpt 
         BackColor       =   &H8000000C&
         Caption         =   "Practice"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtForename 
         Height          =   285
         Left            =   4080
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtRefEnd 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtRefStart 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox comboPractice 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lblPatFore 
         BackColor       =   &H8000000C&
         Caption         =   "Forename"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblPatId 
         BackColor       =   &H8000000C&
         Caption         =   "NHS/Hosp/Surname"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblInter 
         BackColor       =   &H8000000C&
         Caption         =   "to"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   2280
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private reportSQL As String
Private fileSQL As String
Private addSQL(6) As String
Private strDefault As String
Private ltIndex As Integer
Private svcType As Integer

Private iceCmd As ADODB.Command

Private strArray() As String
Private rqRS As New ADODB.Recordset
Private repFile As String
Private logFile As String
Private sortCol As String

Private Function CheckRequeueCriteria(repIndex As Long, _
                                      natCode As String, _
                                      LTSIndex As Long) As Integer
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim rVal As Integer
   
   If LTSIndex = -1 Then
      rVal = 1
   Else
      strSQL = "SELECT * " & _
               "FROM EDI_Recipients " & _
               "WHERE EDI_NatCode = '" & natCode & "' " & _
                  "AND EDI_Active = 1"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      If RS.RecordCount = 0 Then
         rVal = 2
      Else
         RS.Close
         strSQL = "SELECT Service_Report_Id " & _
                  "FROM Service_Reports " & _
                  "WHERE Service_Report_Index = " & repIndex
         RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
         If RS.RecordCount = 0 Then
            rVal = 3
         Else
            RS.Close
            strSQL = "SELECT * " & _
                     "FROM EDI_Rep_List " & _
                     "WHERE EDI_Report_Index = " & repIndex
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            If RS.EOF Then
               rVal = 0
            Else
               rVal = 4
            End If
         End If
      End If
   End If
   RS.Close
   Set RS = Nothing
   CheckRequeueCriteria = rVal
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmLogs.CheckRequeueCriteria"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Sub chkSearchOpt_Click(Index As Integer)
   On Error GoTo procEH
   Dim blnVisible As Boolean
   Dim RS As New ADODB.Recordset
   
   blnVisible = (chkSearchOpt(Index).value = 1)
   Select Case Index
      Case 0
         txtSurname.Visible = blnVisible
         txtForename.Visible = blnVisible
         lblPatId.Visible = blnVisible
         lblPatFore.Visible = blnVisible
         
      Case 1
         comboPractice.Visible = blnVisible
   
      Case 2
         comboClin.Visible = blnVisible
         
      Case 3
         txtRepId.Visible = blnVisible
         
      Case 4
         txtRefStart.Visible = blnVisible
         lblInter.Visible = blnVisible
         txtRefEnd.Visible = blnVisible
         
      Case 5
         cboKorner.Visible = blnVisible
'         txtDiscipline.Visible = blnVisible
         
   End Select
   
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmLogs.chlSearchOpt_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub cmdClear_Click()
   Dim i As Integer
   
   With DGResults
      For i = 0 To .SelBookmarks.Count - 1
         .SelBookmarks.Remove 0
      Next i
   End With
End Sub

Private Sub cmdClose_Click()
   Set iceCmd = Nothing
   Unload Me
End Sub

Private Sub cmdGo_Click()
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nullStr As String
   Dim capStr As String
   Dim natCode As String
   Dim sqlAND As String
   Dim strTmp As String
   Dim i As Integer
   Dim max(4) As Long
   Dim exWidth As Integer
   Dim recsToAnalyse As Long
   Dim blnContinue As Boolean
   
   strSQL = "SELECT count(*) " & _
            "FROM Service_ImpExp_Headers sh WITH (Index (IX_Service_ImpExp_Headers)) " & _
            "WHERE sh.Date_Added > '" & Format(dtpStart.value, "yyyymmdd 00:00:00") & "'and sh.Date_Added < '" & Format(dtpEnd.value, "yyyymmdd 23:59:59") & "' " & _
               "AND sh.EDI_LTS_Index = " & ltIndex & _
               "AND Service_Type = " & svcType
   
   rqRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   recsToAnalyse = rqRS(0)
   rqRS.Close
   
   blnContinue = True
   iceCon.CommandTimeout = 20
   
   If recsToAnalyse > 35000 Then
      blnContinue = MsgBox("This involves analysing over " & recsToAnalyse & " records." & _
                          "The search will possibly take longer than 5 minutes. Are you " & _
                          "sure you wish to continue?", vbYesNo, "Data warning") = vbYes
      If blnContinue Then
         iceCon.CommandTimeout = 600
      End If
   End If
   
   If blnContinue Then
      fraResults.Caption = "Searching..."
      DGResults.Caption = "Searching..."
      frmLogs.MousePointer = vbHourglass
      DGResults.Visible = False
      frmLogs.Refresh
                  
      If optShow(0).value = True Then
         strSQL = Replace(reportSQL, "<P1>", Format(dtpStart.value, "yyyymmdd"))
      Else
         strSQL = Replace(fileSQL, "<P1>", Format(dtpStart.value, "yyyymmdd"))
      End If
      
      strSQL = Replace(strSQL, "<P2>", Format(dtpEnd.value, "yyyymmdd 23:59:59"))
      strSQL = Replace(strSQL, "<P3>", ltIndex)
      strSQL = Replace(strSQL, "<P4>", svcType)
      
      For i = 0 To 5
         sqlAND = ""
         If chkSearchOpt(i).value = 1 Then
            Select Case i
               Case 0
                  If optShow(0).value = True Then
                     addSQL(0) = Replace(addSQL(0), "<P1>", Replace(txtSurname.Text, "'", "''"))
                     sqlAND = Replace(addSQL(0), "<P2>", Replace(txtForename.Text, "'", "''"))
                  End If
               
               Case 1
                  If Trim(comboPractice.Text <> "") Then
                     sqlAND = Replace(addSQL(1), "<P>", Left(comboPractice.Text, InStr(1, comboPractice.Text, " - ") - 1))
                  End If
                  
               Case 2
                  If optShow(0).value = True Then
                     If Trim(comboClin.Text <> "") Then
                        strTmp = Trim(Left(comboClin.Text, InStr(1, comboClin.Text, "[") - 1))
                        sqlAND = Replace(addSQL(2), "<P>", Replace(strTmp, "'", "''"))
                     End If
                  End If
               
               Case 3
                  If optShow(0).value = True Then
                     If Trim(txtRepId.Text) <> "" Then
                        sqlAND = Replace(addSQL(3), "<P>", txtRepId.Text)
                     End If
                  End If
               
               Case 4
                  If Trim(txtRefStart.Text) <> "" Then
                     sqlAND = Replace(addSQL(4), "<P1>", txtRefStart.Text)
                     sqlAND = Replace(sqlAND, "<P2>", txtRefEnd.Text)
                  End If
               
               Case 5
   '               If Trim(txtDiscipline.Text) <> "" Then
                     sqlAND = Replace(addSQL(5), "<P>", Left(cboKorner.Text, 3)) 'txtDiscipline.Text)
   '               End If
               
            End Select
            strSQL = strSQL & sqlAND
            
         End If
      Next i
         
      rqRS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      iceCon.CommandTimeout = 10
      Set DGResults.DataSource = rqRS
      
   '   For i = 0 To 3
   '      max(i) = 0
   '   Next i
      
   '   Do Until rqrs.EOF
   '      If frmLogs.TextWidth(Trim(RS!Report)) > max(0) Then
   '         max(0) = frmLogs.TextWidth(Trim(RS!Report))
   '      End If
   '      If frmLogs.TextWidth(Trim(RS!Patient)) > max(1) Then
   '         max(1) = frmLogs.TextWidth(Trim(RS!Patient))
   '      End If
   '      If frmLogs.TextWidth(Trim(RS!practice)) > max(2) Then
   '         max(2) = frmLogs.TextWidth(Trim(RS!practice))
   '      End If
   '      If frmLogs.TextWidth(Trim(RS!Date_Added)) > max(3) Then
   '         max(3) = frmLogs.TextWidth(Trim(RS!Date_Added))
   '      End If
   '      rqrs.MoveNext
   '   Loop
      
      If rqRS.RecordCount > 0 Then
         rqRS.MoveFirst
         cmdRequeue.Enabled = True
         With DGResults
            .Columns(0).Width = 1800   '  max(0) + exWidth
            .Columns(1).Width = 2500   '  max(1) + exWidth
            .Columns(2).Width = 900 '  max(2) + exWidth
            .Columns(5).Width = 1900   '  max(3) + exWidth
         
            If comboSvcType.ListIndex = 0 Then
               fraResults.Caption = " Laboratory Report(s) found"
               .BackColor = &HC0FFC0
               .ForeColor = &HFF&
               cmdRequeue.BackColor = &HC0FFC0
            ElseIf comboSvcType.ListIndex = 1 Then
               fraResults.Caption = " EDI Report(s) found"
               .BackColor = &HC0FFFF
               .ForeColor = &HFF0000
               cmdRequeue.BackColor = &HC0FFFF
            Else
               fraResults.Caption = " Non-standard Report(s) found"
               .BackColor = &HFFFFFF
               .ForeColor = &H0&
            End If
            .Tag = rqRS.RecordCount - 1
   '         exWidth = 150
   '         .Columns(0).Width = max(0) + exWidth
   '         .Columns(1).Width = max(1) + exWidth
   '         .Columns(2).Width = max(2) + exWidth
   '         .Columns(5).Width = max(3) + exWidth
            .Visible = True
            capStr = Mid(comboSvcType.List(comboSvcType.ListIndex), InStr(1, comboSvcType.List(comboSvcType.ListIndex), " - ") + 3)
            fraResults.Caption = rqRS.RecordCount & " " & capStr & " found"
         End With
      Else
         cmdRequeue.Enabled = False
         cmdRequeue.BackColor = &H8000000F
         fraResults.Caption = "No records found"
      End If
      
      frmLogs.MousePointer = vbNormal
   End If
   
   iceCon.CommandTimeout = 10
   
   Set rqRS = Nothing
   Exit Sub
   
procEH:
      If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmLogs.cmdGo_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub cmdRequeue_Click()
   On Error GoTo procEH
   Dim bkm As Variant
   Dim i As Integer
   Dim pbStep As Long
'   Dim rCtrl As New requeueControl
   Dim colId As Integer
   Dim rqCount As Integer
   Dim lastRepId As Long
   Dim dCount As Integer
   Dim bLen As Integer
   Dim mbPrompt As String
   Dim strReject As String
   Dim strSQL As String
   Dim strRepList As String
   Dim strMsgData As String
   Dim tBuf As New StringBuffer
   Dim tBufRep As New StringBuffer
   Dim tBufMsg As New StringBuffer
   Dim iceCmd As New ADODB.Command
   Dim fSep As String
   
'  0  Service_Id AS Report,
'  1  Patient_Name AS Patient,
'  2  EDI_Org_NatCode AS Practice,
'  3  Clinician_National_Code AS Clinician,
'  4  Control_Ref AS Interchange,
'  5  Discipline,
'  6  Service_ImpExp_Messages.Date_Added,
'  7  ImpExp_File,
'  8  Service_Report_Index AS Id,
'  9  Service_Type AS Type,
'  10 EDI_Individual_Index_To AS Clinician,
'  11 Service_ImpExp_Messages.EDI_LTS_Index AS Trader
'  12 Destination
'  13 Messages EDI_LTS_Index
'  14 Service_ImpExp_Id
   
   strReject = ""
   rqCount = 0
   
   rqCount = DGResults.SelBookmarks.Count
   If rqCount > 0 Then
      pbStep = rqCount / 10
      pbStep = (pbStep - (pbStep Mod 8)) / 8
      
      DGResults.Visible = False
      
      With lblReqInfo
         .Caption = "Requeuing " & rqCount & " records"
         .Visible = True
      End With
      
      With pbReq
         .value = 0
         .Min = 0
         .max = rqCount + (pbStep * 8)
         .Visible = True
      End With
      
      With rCtrl
         If optShow(1).value = True Then
            colId = 5
            .RequeueReport = False
         Else
            colId = 10
            .RequeueReport = True
         End If
         
         .CallerForm = Me
         .UseProgressBar pbReq, pbStep
      End With
      
      Unload frmEDIRequeue
      
      Me.Caption = "Processing please wait..."
      Me.MousePointer = vbHourglass
      Me.Refresh
      
      With DGResults
         For Each bkm In .SelBookmarks
            rCtrl.RequeueItem = .Columns(colId).CellText(bkm)
         Next
      End With
            
      rCtrl.CallerForm = Me
      rCtrl.RequeueData
   
   Else
      MsgBox "No reports selected!", vbInformation, "Requeue Reports"
   End If
   
   lblReqInfo.Visible = False
   pbReq.Visible = False
   DGResults.Visible = True
   Me.Caption = "Log Search Options"
   Me.MousePointer = vbNormal
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmLogs.cmdRequeue_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub cmdSelect_Click()
   Dim i As Integer
   Dim intRel As Integer
   Dim xx As Variant
   Dim RS As New ADODB.Recordset
   
   Me.MousePointer = vbHourglass
   With DGResults
      .Visible = False
      Set RS = .DataSource
      RS.MoveFirst
      .Row = 0
      Do Until RS.EOF
         .SelBookmarks.Add .RowBookmark(.Row)
         RS.MoveNext
      Loop
      Set RS = Nothing
      .Visible = True
   End With
   Me.MousePointer = vbNormal
End Sub

Private Sub comboClin_Change()
   With comboClin
'      .ToolTipText = .ItemData(.ListIndex)
   End With
End Sub

Private Sub comboLTI_Click()
   ltIndex = Left(comboLTI.Text, 2)
End Sub

Private Sub comboSvcType_Click()
   With comboSvcType
      If .ListIndex = 0 Then
         .BackColor = &HC0FFC0
         .ForeColor = &HFF&
      ElseIf .ListIndex = 1 Then
         .BackColor = &HC0FFFF
         .ForeColor = &HFF0000
      End If
   End With
   
   svcType = Left(comboSvcType, 2)
End Sub

Private Sub comboSvcType_DropDown()
   With comboSvcType
      .BackColor = &HFFFFFF
      .ForeColor = &H0&
   End With
End Sub

'Private Sub DGResults_HeadClick(ByVal colIndex As Integer)
'   sortCol = DGResults.Columns(colIndex).dataField
'   RS.Sort = DGResults.Columns(colIndex).dataField
'End Sub

Private Sub dtpEnd_CloseUp()
   If dtpStart.value > dtpEnd.value Then
      dtpStart.value = dtpEnd.value
   End If
End Sub

Private Sub dtpStart_CloseUp()
   If dtpEnd.value < dtpStart.value Then
      dtpEnd.value = dtpStart.value
   End If
End Sub

Private Sub Form_Load()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim bodySQL As String
   
   dtpEnd.value = Format(Now(), "dd/mm/yyyy")
'   DTPEnd.value = "12/11/2002"
   dtpStart.value = DateAdd("d", -14, dtpEnd.value)
   
   strSQL = "SELECT * " & _
            "FROM Service_Types "
   
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   bodySQL = "SELECT TOP 1 Service_Type, description FROM Service_ImpExp_Headers " & _
             "INNER JOIN Service_Types ON Service_Type = Type_Index " & _
             "WHERE Service_Type="
   
   strSQL = ""
   
   Do Until RS.EOF
      If strSQL <> "" Then
         strSQL = strSQL & " UNION " & bodySQL
      Else
         strSQL = bodySQL
      End If
      
      strSQL = strSQL & RS!Type_Index
      RS.MoveNext
   Loop
   
   RS.Close
   
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      comboSvcType.AddItem RS!Service_Type & " - " & RS!Description
      RS.MoveNext
   Loop
   RS.Close
   
   With comboSvcType
      .Text = comboSvcType.List(0)
      .ListIndex = 0
      svcType = Left(.Text, 2)
   End With

   strSQL = "SELECT EDI_Msg_Type, EDI_LTS_Index FROM EDI_Local_Trader_Settings ORDER BY EDI_LTS_Index"
   RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
   Do Until RS.EOF
      comboLTI.AddItem RS!EDI_LTS_Index & " - " & RS!EDI_Msg_Type
      RS.MoveNext
   Loop
   RS.Close
   
   comboLTI.ListIndex = 0
   ltIndex = Left(comboLTI.Text, 2)

   strSQL = "SELECT EDI_NatCode, EDI_Name " & _
            "FROM EDI_Recipients"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      comboPractice.AddItem RS!EDI_NatCode & " - " & RS!EDI_Name
      RS.MoveNext
   Loop
   RS.Close
   
   strSQL = "SELECT DISTINCT c.Clinician_National_Code, " & _
               "Clinician_Local_Code AS LocalCode, " & _
               "Case " & _
                  "When EDI_GP_Name is Null then 'N/A' " & _
                  "When EDI_OP_Name is null then EDI_GP_Name " & _
                  "Else EDI_OP_Name " & _
               "End as Individual, " & _
               "Case " & _
                  "When EDI_Org_NatCode is null then 'N/A' " & _
                  "Else EDI_Org_NatCode " & _
               "End As Practice, " & _
               "Clinician_Surname " & _
            "FROM Clinician c " & _
               "INNER JOIN Clinician_Local_Id cl " & _
               "ON c.Clinician_Index = cl.Clinician_Index " & _
               "LEFT JOIN EDI_Matching em " & _
                  "INNER JOIN EDI_Recipient_Individuals er ON " & _
                  "em.Individual_Index = er.Individual_Index " & _
               "ON Clinician_Local_Code = EDI_Local_Key3 " & _
            "ORDER BY Practice, c.Clinician_Surname"
   
'   strSQL = "SELECT Clinician_National_Code, Clinician_Surname " & _
'            "FROM Clinician " & _
'            "ORDER BY Clinician_National_Code"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
'      comboClin.AddItem RS!Clinician_Surname & vbTab & RS!Clinician_National_Code & vbTab & " (" & RS!practice & ")"
      comboClin.AddItem Left(RS!Clinician_National_Code & "         ", 9) & "[" & RS!practice & "] " & RS!Clinician_Surname
      RS.MoveNext
   Loop
   
   RS.Close
   
   For i = 0 To chkSearchOpt.Count - 1
      chkSearchOpt(i).value = 0
   Next i
   
   strSQL = "SELECT * " & _
            "FROM CRIR_Specialty " & _
            "where Specialty_code between '800' and '899' or Specialty_code = '502'"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Do Until RS.EOF
      cboKorner.AddItem RS!Specialty_Code & " - " & RS!Specialty
      RS.MoveNext
   Loop
   RS.Close
   
   txtSurname.Visible = False
   lblPatId.Visible = False
   txtSurname.Text = ""
   txtForename.Visible = False
   lblPatFore.Visible = False
   txtForename.Text = ""
   comboPractice.Visible = False
   comboClin.Visible = False
   txtRepId.Visible = False
   txtRepId.Text = ""
   txtRefStart.Visible = False
   txtRefStart.Text = ""
   lblInter.Visible = False
   txtRefEnd.Visible = False
   txtRefEnd.Text = ""
   cboKorner.Visible = False
'   txtDiscipline.Visible = False
'   txtDiscipline.Text = ""
   sortCol = ""
   
   optShow(0).value = True
   
   repFile = fs.BuildPath(App.Path, "RepListFile.ice")
   logFile = fs.BuildPath(App.Path, "MsgFile.ice")
   
   'reportSQL = "SELECT * " & _
               "FROM Requeue_View " & _
               "WHERE Datediff(d, '<P1>', Date_Added) >=0 " & _
                  "AND DateDiff(d, '<P2>', Date_Added) <=0 " & _
                  "AND Type = <P3>"
   
   'fileSQL = "SELECT DISTINCT Date_Added, " & _
             "Practice, " & _
             "ImpExp_File, " & _
             "Interchange, " & _
             "Discipline, " & _
             "Type, " & _
             "ImpExpRef " & _
             "FROM Requeue_View " & _
               "WHERE Datediff(d, '<P1>', Date_Added) >=0 " & _
                  "AND DateDiff(d, '<P2>', Date_Added) <=0 " & _
                  "AND Type = <P3>"
   fileSQL = "SELECT DISTINCT sh.Date_Added, ImpExp_File, Control_Ref, " & _
                  "Discipline , Service_Type, sh.Service_ImpExp_Id " & _
               "FROM Service_ImpExp_Headers sh WITH (Index (IX_Service_ImpExp_Headers)) " & _
                  "INNER JOIN Service_ImpExp_Messages sm " & _
                  "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
                  "INNER JOIN EDI_Health_Parties " & _
                  "ON sm.Service_Report_Index = EDI_Report_Index " & _
                     "AND EDI_HP_Type='901' " & _
                  "INNER JOIN EDI_Matching em " & _
                     "INNER JOIN EDI_Recipient_Individuals ei " & _
                     "ON em.Individual_Index = ei.Individual_Index " & _
                  "ON EDI_HP_Nat_Code = EDI_Local_Key1 " & _
               "WHERE sh.Date_Added > '<P1>' and sh.Date_Added < '<P2>' " & _
                  "AND sh.EDI_LTS_Index = <P3> " & _
                  "AND Service_Type = <P4>"

   
   reportSQL = "SELECT DISTINCT rTrim(left(Service_Id,16)) As Report, Patient_Name, Patient_Id_Key, " & _
                  "EDI_Org_NatCode as Practice, ImpExp_File, Control_Ref, " & _
                  "Discipline, sh.Date_Added, sm.Service_Report_Index, Service_Type, " & _
                  "Service_ImpExp_Message_Id , sh.Service_ImpExp_Id " & _
             "FROM Service_ImpExp_Headers sh WITH (Index (IX_Service_ImpExp_Headers)) " & _
                  "INNER JOIN Service_ImpExp_Messages sm " & _
                     "INNER JOIN Patient_Local_Ids pl " & _
                     "ON pl.Patient_Local_Id = sm.Patient_Local_Id " & _
                        "AND Retired = 0 " & _
                  "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
                  "INNER JOIN EDI_Health_Parties " & _
                     "INNER JOIN EDI_Matching em " & _
                        "INNER JOIN EDI_Recipient_Individuals ei " & _
                        "ON em.Individual_Index = ei.Individual_Index " & _
                     "ON EDI_HP_Nat_Code = EDI_Local_Key1 " & _
                  "ON sm.Service_Report_Index = EDI_Report_Index " & _
                     "AND EDI_HP_Type='901' " & _
               "WHERE sh.Date_Added > '<P1>'and sh.Date_Added < '<P2>' " & _
                  "AND sh.EDI_LTS_Index = <P3> " & _
                  "AND Service_Type = <P4>"
   
   addSQL(0) = " AND Patient_Id_Key in (Select PatKey from ICECONFIG_Requeue_Patient('<P1>','<P2>'))"
   addSQL(1) = " AND EDI_Org_NatCode = '<P>'"
   addSQL(2) = " AND Service_Report_Index in (SELECT RepIndex FROM ICECONFIG_Requeue_clinician('<P>'))"
   addSQL(3) = " AND Report Like '<P>' + '%'"
   addSQL(4) = " AND Control_Ref between <P1> and <P2>"
   addSQL(5) = " AND Discipline = '<P>'"
      
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmLogs.FormLoad"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   rqRS.Close
   Set rqRS = Nothing
End Sub

Private Sub optShow_Click(Index As Integer)
   If Index = 0 Then
      comboSvcType.Enabled = True
      chkSearchOpt(0).Visible = True
      chkSearchOpt(2).Visible = True
      chkSearchOpt(3).Visible = True
   Else
      comboSvcType.ListIndex = 1
      comboSvcType.Enabled = False
      chkSearchOpt(0).value = 0
      chkSearchOpt(0).Visible = False
      chkSearchOpt(2).value = 0
      chkSearchOpt(2).Visible = False
      chkSearchOpt(3).value = 0
      chkSearchOpt(3).Visible = False
   End If
End Sub

Private Sub txtRefEnd_GotFocus()
   txtRefEnd.SelStart = 0
   txtRefEnd.SelLength = Len(txtRefEnd.Text)
End Sub

Private Sub txtRefStart_Change()
   If txtRefStart.Text = "" Then
      txtRefEnd.Text = ""
      txtRefEnd.Enabled = False
   Else
      txtRefEnd.Enabled = True
      txtRefEnd.Text = txtRefStart.Text
   End If
End Sub

Private Sub txtRefStart_GotFocus()
   txtRefStart.SelStart = 0
   txtRefStart.SelLength = Len(txtRefEnd.Text)
End Sub
