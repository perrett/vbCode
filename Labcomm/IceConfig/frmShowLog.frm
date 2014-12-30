VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmShowLog 
   Caption         =   "File Name: "
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   7935
      Left            =   135
      TabIndex        =   11
      Top             =   120
      Width           =   9375
      ExtentX         =   16536
      ExtentY         =   13996
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter"
      Height          =   300
      Left            =   6840
      TabIndex        =   5
      Top             =   8280
      Width           =   705
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter Options"
      Enabled         =   0   'False
      Height          =   1365
      Left            =   1200
      TabIndex        =   1
      Top             =   8160
      Width           =   4155
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "All"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "Errors only"
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Go"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3240
         TabIndex        =   4
         Top             =   480
         Width           =   675
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   225
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51707906
         CurrentDate     =   37516
      End
      Begin VB.Label Label2 
         Caption         =   "View"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Show From"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   8880
      Width           =   1650
   End
   Begin RichTextLib.RichTextBox rtbLogs 
      Height          =   7950
      Left            =   150
      TabIndex        =   10
      Top             =   75
      Visible         =   0   'False
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   14023
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmShowLog.frx":0000
   End
End
Attribute VB_Name = "frmShowLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private logFile As String

Private Sub AnalyseLog()
   On Error GoTo procEH
   Dim fNum As Integer
   Dim fBuf As String * 1024
   Dim logSTr As String
   Dim buf As String
   Dim dBuf As New StringBuffer
   Dim i As Integer
   Dim dataField(7) As String
   Dim dPos As Long
   Dim dLen As Long
   Dim fPos As Long
   Dim ePos As Long
   Dim pad As Integer
   Dim blnShow As Boolean
   Dim strDisp(2, 3) As String
   Dim maxLen As Integer
   Dim maxWidth(3) As Integer
   Dim j As Integer
   
   fNum = FreeFile
   Open logFile For Binary As #fNum
   
   Do Until EOF(fNum)
      Get #1, , fBuf
      dBuf.Append fBuf
   Loop
   
   buf = Left(dBuf.ActualValue, LOF(fNum))
   Close #fNum
   dBuf.Clear
   
   fPos = 1
   dPos = 1
   Do Until dPos >= Len(buf)
      For i = 0 To 6
         If Asc(Mid(buf, dPos + 1, 1)) > 0 Then
            dLen = (Asc(Mid(buf, dPos + 1, 1)) * 255) + Asc(Mid(buf, dPos, 1)) + 1
            pad = 2
         Else
            dLen = Asc(Mid(buf, dPos, 1))
            pad = 2
         End If
         
         If Mid(buf, dPos + pad, dLen) = "." Then
            logSTr = "<None>"
         Else
            logSTr = Mid(buf, dPos + pad, dLen)
         End If
         dataField(i) = logSTr ' & vbTab ' & vbCrLf
         dPos = dPos + dLen + 2
      Next i
      
      dPos = dPos + 2
      
      If chkFilter.value = 1 Then
         blnShow = (CDate(Mid(dataField(0), 2, Len(dataField(0)) - 3)) > dtpTime.value)
      Else
         blnShow = True
      End If
         
      If blnShow Then
         strDisp(0, 0) = "Date/Time: " & dataField(0) & vbTab
         strDisp(0, 1) = "Procedure: " & dataField(1) & vbTab
         strDisp(0, 2) = "Operation: " & dataField(2) & vbCrLf
         
         strDisp(1, 0) = "Error Number: " & dataField(4) & vbTab
         strDisp(1, 1) = "Error Source: " & dataField(5)
         strDisp(1, 2) = vbCrLf
         
'         strDisp(2, 0) = "Description: " & dataField(3) & vbTab
         
         logSTr = ""
         For i = 0 To 1
            For j = 0 To 2
               maxLen = Me.TextWidth(strDisp(i, j))
               If maxLen > maxWidth(2) Then
                  maxWidth(j) = maxLen
               End If
               logSTr = logSTr & strDisp(i, j)
            Next j
            dBuf.Append logSTr
            logSTr = ""
         Next i
         dBuf.Append "Description: " & dataField(3) & vbCrLf & "Information: " & dataField(6) & vbCrLf & vbCrLf
'
'         maxLen(0) = frmShowLog.TextWidth(logSTr)
'         dBuf.Append "Date/Time: " & dataField(0) & vbTab & "Procedure: " & dataField(1) & vbTab & "Operation: " & dataField(2) & vbCrLf
'         dBuf.Append "Error Number: " & dataField(4) & vbTab & "Description: " & dataField(3) & vbTab & "Error Source: " & dataField(5) & vbCrLf
'         dBuf.Append "Information: " & dataField(6)
'         dBuf.Append "                   <<<   >>>" & vbCrLf
      End If
   Loop
   With rtbLogs
      .Text = dBuf.ActualValue
      .SelStart = 1
      .SelLength = dBuf.Length
      .SelTabCount = 2
      .SelTabs(0) = maxWidth(0) + 200
      .SelTabs(1) = maxWidth(0) + maxWidth(1) + 500
      .SelLength = 0
   End With
   Exit Sub
   
procEH:
   rtbLogs.Text = ">>> Unable to View Log File - Incorrect format" & vbCrLf & _
                  "    Error " & Err.Number & " (" & Err.Description & ")"
End Sub

Private Sub chkFilter_Click()
   If chkFilter.value = 1 Then
      fraFilter.Enabled = True
      dtpTime.Enabled = True
      cmdFilter.Enabled = True
   Else
      fraFilter.Enabled = False
      dtpTime.Enabled = False
      cmdFilter.Enabled = False
   End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFilter_Click()
   AnalyseLog
End Sub

Private Sub Form_Load()
   dtpTime.value = DateAdd("N", -10, Now())
'   Me.ScaleMode = 4
   wb.Navigate logFile
   
'   AnalyseLog
'   Me.ScaleMode = 1
End Sub

Public Property Let LogFileName(strNewValue As String)
   logFile = strNewValue
   frmShowLog.Caption = "File: " & logFile
End Property

Private Sub Form_Resize()
   rtbLogs.Width = frmShowLog.Width - 350
End Sub
