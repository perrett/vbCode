VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraNav 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Navigation"
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Close"
         Height          =   375
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Up"
         Height          =   375
         Left            =   3481
         TabIndex        =   3
         ToolTipText     =   "Up one level in the hierarchy"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdHome 
         BackColor       =   &H0080FF80&
         Caption         =   "Home"
         Height          =   375
         Left            =   1527
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Return to the navigation start point"
         Top             =   120
         Width           =   855
      End
   End
   Begin SHDocVwCtl.WebBrowser wbFolder 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7850
      ExtentX         =   13847
      ExtentY         =   8705
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal HWnd As Long) As Long

Private strArray() As String
Private filtArr() As String
Private testStr As String
Private WithEvents tt As Shell32.ShellFolderView
Attribute tt.VB_VarHelpID = -1
Private homeDir As String
Private blnBack As Boolean

Private WithEvents hd As MSHTMLCtl.HTMLDocument
Attribute hd.VB_VarHelpID = -1
Private tr As MSHTMLCtl.IHTMLTxtRange

Private Sub ReplaceIt()
   Dim msgStr As String
   Dim i As Integer
   Dim j As Integer
   Dim soFar As String
   Dim rChar As String
   Dim strTemp As String
   
   strArray = Split(testStr, "'")
   filtArr = Filter(strArray, "?")
   soFar = "?"
   
   For i = 0 To UBound(filtArr)
      rChar = Right(filtArr(i), 1)
      If rChar = "?" Then
         If Mid(filtArr(i), Len(filtArr(i)) - 1, 1) = "?" Then
            rChar = ""
         Else
            rChar = "'"
         End If
      End If
      If InStr(1, soFar, rChar) = 0 Then
         soFar = soFar & rChar
      End If
   Next i
   
   strTemp = testStr
   For i = 1 To Len(soFar)
      strTemp = Replace(strTemp, "?" & Mid(soFar, i, 1), Chr(i))
   Next i
   
   strArray = Split(strTemp, "'")
   
   For i = 0 To UBound(strArray)
      Select Case Left(strArray(i), 3)
         Case "UNA"
         
         Case "UNB"
         
         Case "UNH"
         
         Case "BGM"
         
         Case "DTM"
         
         Case "NIR"
            
         Case "UNT"
         
      End Select
      
      For j = 1 To Len(soFar)
         strArray(i) = Replace(strArray(i), Chr(j), Mid(soFar, j, 1))
      Next j
      msgStr = msgStr & strArray(i) & vbCrLf
   Next i
   Debug.Print msgStr
End Sub

Public Property Let HomeDirectory(strNewValue As String)
   homeDir = strNewValue
End Property

Private Sub cmdCopy_Click()
   Dim hDoc As MSHTMLCtl.HTMLDocument
   Dim tRange As MSHTMLCtl.IHTMLTxtRange
   
   Set hDoc = wbFolder.document
   
   If hDoc.selection <> "" Then
      Set tRange = hDoc.selection.createRange
      MsgBox tRange.Text
   End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdHome_Click()
   bCtrl.ProcUnBindFromBrowser
   bCtrl.NavigateTo homeDir
   bCtrl.ProcBindToBrowser
End Sub

Private Sub cmdUp_Click()
   blnBack = True
   bCtrl.ProcUnBindFromBrowser
   wbFolder.GoBack
   bCtrl.ProcBindToBrowser
End Sub

Private Sub Form_Load()
   Dim fBuf As String * 8192
   Dim buf As String
   Dim fName As String
   Dim i As Integer
   
   homeDir = "C:\Ice\Labcomm\Intray"
   bCtrl.BrowserType = "Classic"
   Me.Show
   
   wbFolder.Visible = True
   bCtrl.Register wbFolder
   
   bCtrl.NavigateTo homeDir
   
   Set tt = wbFolder.document
End Sub

Private Sub Form_Resize()
   If frmMain.Height < 2500 Then
      frmMain.Height = 2500
   End If
   
   If frmMain.Width < 4000 Then
      frmMain.Width = 4000
   End If
   
   fraNav.Width = frmMain.Width - 105
   cmdHome.Left = (fraNav.Width / 4) - 427
   cmdUp.Left = (2 * (fraNav.Width / 4)) - 427
   cmdClose.Left = (3 * (fraNav.Width / 4)) - 427
   wbFolder.Height = frmMain.Height - 495
   wbFolder.Width = frmMain.Width - 70
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bCtrl.ProcUnBindFromBrowser
   bCtrl.BrowserType = "Standard"
End Sub

Private Sub hd_onselectionchange()
   Set tr = hd.selection.createRange
   If tr.Text = "" Then
      cmdCopy.Enabled = False
   Else
      cmdCopy.Enabled = True
   End If
End Sub

Private Function tt_DefaultVerbInvoked() As Boolean
   Dim fBuf As String * 8192
   Dim buf As String
   Dim fName As String
   Dim i As Integer
   Dim aMsg As New DecodeAck
   
   If blnBack Then
      blnBack = False
'      Set tt = wbFolder.document
   Else
      Debug.Print "tt.SelectionChanged (Selected items = " & tt.SelectedItems.count & ")"
      bCtrl.ProcUnBindFromBrowser
      
      If tt.SelectedItems.Item(0).IsFolder = False Then
         If tt.SelectedItems.count > 0 Then
            ackFile = tt.SelectedItems.Item(0).Path
            bCtrl.ProcBindToBrowser
            
            aMsg.ReadAckFile ackFile
            Set tt = Nothing
            aMsg.ShowAck
            Set hd = wbFolder.document
         End If
      End If
   End If

   tt_DefaultVerbInvoked = True
End Function

Private Sub Xtt_SelectionChanged()
   Dim fBuf As String * 8192
   Dim buf As String
   Dim fName As String
   Dim i As Integer
   Dim aMsg As New DecodeAck
   
   If blnBack Then
      blnBack = False
'      Set tt = wbFolder.document
   Else
      Debug.Print "tt.SelectionChanged (Selected items = " & tt.SelectedItems.count & ")"
      bCtrl.ProcUnBindFromBrowser
      
      If tt.SelectedItems.count > 0 Then
         ackFile = tt.SelectedItems.Item(0).Path
         bCtrl.ProcBindToBrowser
         
         aMsg.ReadAckFile ackFile
         Set tt = Nothing
         aMsg.ShowAck
         Set hd = wbFolder.document
      End If
   End If
End Sub

Private Function tt_VerbInvoked() As Boolean
   Dim i As Integer
   
   tt_VerbInvoked = False
End Function

Private Sub wbFolder_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
   bCtrl.ProcUnBindFromBrowser
End Sub

Private Sub wbFolder_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
   bCtrl.ProcBindToBrowser
   If blnBack Then
'      blnBack = False
      Set tt = wbFolder.document
   End If
   Debug.Print bCtrl.BrowserClassType & " (tt set to: " & TypeName(tt) & ")"
'   If bCtrl.BrowserClassType = "FolderView" Then
'      Set tt = wbFolder.document
'   End If
End Sub
