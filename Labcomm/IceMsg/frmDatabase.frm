VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Selection"
   ClientHeight    =   3060
   ClientLeft      =   5940
   ClientTop       =   4920
   ClientWidth     =   2835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUDL 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   120
         Picture         =   "frmDatabase.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton cmdDB 
         Caption         =   " "
         Default         =   -1  'True
         Height          =   450
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Please select the database you wish to connect to:"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   2400
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Amend login details"
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cboDB 
         Height          =   315
         ItemData        =   "frmDatabase.frx":030A
         Left            =   120
         List            =   "frmDatabase.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "#"
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Username"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lvDB 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private multiUDL As Boolean
Private UDLPath As String
Private UDLConnectData As Variant
Private UDLFile As String
Private blnEditUDL As Boolean
Private dbButton As Integer

Private iceCon As New ADODB.Connection
Private sqlServer As New SQLDMO.sqlServer
Private sqlDb As SQLDMO.Database
Private UDLServer As Variant
Private UDLDatabase As String
Private dbUser As String
Private dbPass As String

Private Sub BuildUDLButtons()
   Dim curFile As String
   Dim UDLCount As Integer
      
   curFile = Dir(fs.BuildPath(UDLPath, "*.udl"))
   Do Until curFile = ""
      If UDLCount >= cmdDB.Count Then
         Load cmdDB(UDLCount)
         
         frmDatabase.Height = frmDatabase.Height + 500
         Line1.Y1 = Line1.Y1 + 500
         Line1.Y2 = Line1.Y2 + 500
         cmdExit.Top = cmdExit.Top + 500
         
         With cmdDB(UDLCount)
            .Top = 840 + (500 * UDLCount)
            .Left = 120
            .Visible = True
         End With
      End If
      
      ReadUDL fs.BuildPath(UDLPath, curFile)
      
      With cmdDB(UDLCount)
         .Caption = UDLDatabase & " (" & UDLServer & ")"
         .ToolTipText = curFile
         .Tag = UDLConnectData
      End With
      
      UDLCount = UDLCount + 1
      curFile = Dir
   Loop
End Sub

Public Property Let ConfigPath(strNewValue As String)
   UDLPath = strNewValue
End Property


Private Sub cboDB_DropDown()
   ShowAvailableDatabases
End Sub

'Private Sub cboDB_Click()
'   Dim dblist As SQLDMO.NameList
'   Dim i As Integer
'
'   sqlServer.Connect UDLServer, txtUser.Text, txtPass.Text
'   Set dblist = sqlServer.Databases
'   For i = 1 To dblist.Count
'      cboDB.AddItem dblist(i)
'   Next i
'End Sub

'Private Sub cboDB_DropDown()
'   On Error GoTo procEH
'   Dim i As Integer
'
'   cboDB.Clear
'   With sqlServer
'      If sqlServer.VerifyConnection(SQLDMOConn_CurrentState) = False Then
'         .Connect CStr(lvDB.SelectedItem), txtUser.Text, txtPass.Text
'      End If
'
'      For i = 1 To .Databases.Count
'         cboDB.AddItem .Databases(i).Name
'      Next i
'      cboDB.ListIndex = 0
'   End With
'   Exit Sub
'
'procEH:
'   MsgBox Err.Description, vbExclamation, "Login Incorrect"
'   Exit Sub
'End Sub

Private Sub cmdCancel_Click()
   Dim curFile As String
   
   With cmdDB(dbButton)
      curFile = fs.BuildPath(UDLPath, .ToolTipText)
      UDLConnectData = ReadUDL(curFile)
      .Caption = UDLDatabase & " (" & UDLServer & ")"
'      .ToolTipText = curFile
      .Tag = UDLConnectData
   End With
   lvDB.Visible = False
   fraLogin.Visible = False
   fraUDL.Visible = True
End Sub

Private Sub cmdDB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim connectData As Variant
   Dim curFile As String
   
   dbButton = Index
   connectData = cmdDB(Index).Tag
   If RunningInIDE Then
      UDLFile = fs.BuildPath(UDLPath, "ice.udl")
'      DB_UDL_FILE = UDLFile
      
      If Button = vbRightButton Then
         frmDatabase.MousePointer = vbHourglass
         fraUDL.Visible = False
         ShowAvailableServers
      Else
         Me.Hide
      End If
   
   Else
      UDLFile = fs.BuildPath(UDLPath, cmdDB(Index).ToolTipText)
'      DB_UDL_FILE = UDLFile
      Me.Hide
   End If
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub cmdLogin_Click()
   On Error GoTo procEH
   Dim curFile As String
   
   UDLServer = lvDB.SelectedItem.Text
   If cboDB.Text <> "" Then
      UDLDatabase = cboDB.Text
   End If
   
   If lvDB.SelectedItem.Text = "(local)" Then
      dbUser = "sa"
      dbPass = txtPass.Text
   Else
      dbUser = txtUser.Text
      dbPass = txtPass.Text
   End If
   
   If sqlServer.VerifyConnection(SQLDMOConn_CurrentState) = False Then
      sqlServer.Connect UDLServer, dbUser, dbPass
   End If
   
   Set sqlDb = sqlServer.Databases(UDLDatabase)
   
   WriteUDL ' connectData
   Me.Caption = UDLDatabase
   With cmdDB(dbButton)
      curFile = fs.BuildPath(UDLPath, .ToolTipText)
      ReadUDL curFile
      .Caption = UDLDatabase & " (" & UDLServer & ")"
'      .ToolTipText = curFile
      .Tag = UDLConnectData
   End With
   lvDB.Visible = False
   fraLogin.Visible = False
   fraUDL.Visible = True
AfterError:
   Exit Sub
   
procEH:
   If Err.Number = -2147199728 Then
      MsgBox "Unable to find database " & cboDB.Text, vbExclamation, "Specified server " & sqlServer.Name
      Resume AfterError
   End If
End Sub

Public Function ConnectionDetails(Optional ByRef ServerId As SQLDMO.sqlServer = Nothing, _
                                  Optional ByRef dbId As SQLDMO.Database = Nothing, _
                                  Optional ByRef sqlUser As String = "", _
                                  Optional ByRef sqlPass As String = "", _
                                  Optional ByRef UDLsvrId As Variant = Nothing, _
                                  Optional ByRef UDLdb As String = "") As ADODB.Connection
   
   On Error GoTo procEH
   Dim dbCon As New ADODB.Connection
   Dim dbTime As String
   Dim timeOut As Long
   
   UDLPath = Read_Ini_Var("General", "ConfigPath", INIFile)
   
   If UDLPath = "" Then
      UDLPath = App.Path
   ElseIf UCase(fs.GetExtensionName(UDLPath)) = "UDL" Then
      UDLPath = fs.GetParentFolderName(UDLPath)
   End If
   
   UDLFile = fs.BuildPath(UDLPath, "ice.udl")
        
   If fs.FileExists(UDLFile) Then
      dbTime = Read_Ini_Var("General", "dbTimeOut", INIFile)
      If dbTime = "" Then
         timeOut = 10
      Else
         timeOut = CLng(dbTime)
      End If
      
      Initialize
         
      If Not ServerId Is Nothing Then
         If sqlServer.VerifyConnection(SQLDMOConn_CurrentState) = False Then
            sqlServer.Connect UDLServer, dbUser, dbPass
         End If
         
         Set sqlDb = sqlServer.Databases(UDLDatabase)
         
         Set ServerId = sqlServer
         Set dbId = sqlDb
         sqlUser = dbUser
         sqlPass = dbPass
         UDLsvrId = UDLServer
         UDLdb = UDLDatabase
      End If
      
      With dbCon
         .CursorLocation = adUseClient
         .Mode = adModeReadWrite
         .CommandTimeout = timeOut
         .ConnectionTimeout = timeOut
         .Open "FILE NAME=" & UDLFile
      End With
      
   Else
      MsgBox "Specified UDL file in title bar does not exist", vbCritical, UDLFile
      End
   End If
   
   Set ConnectionDetails = dbCon
   Unload Me
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmDatabase.ConnectionDetails"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Sub Initialize()
   Dim curFile As String
   Dim defUDL As String
   Dim UDLCount As Integer
   Dim dbName As String
   Dim strArray() As String
   
   sqlServer.ApplicationName = "IceConfig"
   sqlServer.EnableBcp = True
      
   If (Read_Ini_Var("UDLData", "MultiUDL", INIFile) = 1) Or _
      RunningInIDE Then
      BuildUDLButtons
      Me.Show 1
   Else
      ReadUDL UDLFile
   End If
End Sub

Private Sub Form_Load()
   
   
'   curFile = Dir(fs.BuildPath(UDLPath, "*.udl"))
'   Do Until curFile = ""
'      If udlCount >= cmdDB.Count Then
'         Load cmdDB(udlCount)
'
'         frmDatabase.Height = frmDatabase.Height + 500
'         Line1.Y1 = Line1.Y1 + 500
'         Line1.Y2 = Line1.Y2 + 500
'         cmdExit.Top = cmdExit.Top + 500
'
'         With cmdDB(udlCount)
'            .Top = 840 + (500 * udlCount)
'            .Left = 120
'            .Visible = True
'         End With
'      End If
'
'      ReadUDL fs.BuildPath(UDLPath, curFile)
'
'      With cmdDB(udlCount)
'         .Caption = UDLDatabase & " (" & UDLServer & ")"
'         .ToolTipText = curFile
'         .Tag = UDLConnectData
'      End With
'
'      udlCount = udlCount + 1
'      curFile = Dir
'   Loop
End Sub

Private Function AmendUDL(connectData As Variant) As Variant
   Dim tmpCon As New ADODB.Connection
   Dim dblink As New MSDASC.DataLinks

   tmpCon.ConnectionString = connectData
   dblink.PromptEdit tmpCon
   AmendUDL = tmpCon.ConnectionString
   Set tmpCon = Nothing
   Set dblink = Nothing
End Function

'
'Public Function ReadUDL(UDLFile As String, _
'                         Optional SetConnection As Boolean = True) As String
'   Dim pTag As String
'   Dim pos As Long
'   Dim clen As Integer
'   Dim buf As String
'   Dim constr As Variant
'
''   conStr = ""
'   pTag = "P" & Chr(0) & _
'          "r" & Chr(0) & _
'          "o" & Chr(0) & _
'          "v" & Chr(0) & _
'          "i" & Chr(0) & _
'          "d" & Chr(0) & _
'          "e" & Chr(0) & _
'          "r"
''   UniConStr = ""
'   If fs.FileExists(UDLFile) Then
'      buf = Space(1500)
'      Open UDLFile For Binary As #1
'      Get #1, , buf
'      Close #1
'
'      pos = InStr(1, buf, pTag, vbBinaryCompare)
'      UDLConnectData = Replace(Mid(buf, pos), Chr(0), "")
'      constr = UDLConnectData
'
'      pos = InStr(1, constr, "Source=") + 7
'      clen = InStr(pos, constr, vbCrLf) - pos
'      UDLServer = Mid(constr, pos, clen)
'
''      pos = InStr(1, constr, "Source=") + 7
''      clen = InStr(pos, constr, vbCrLf) - pos
''      UDLDatabase = Mid(constr, pos, clen)
'
'      pos = InStr(1, constr, "User ID=") + 8
'      clen = InStr(pos, constr, ";") - pos
'      dbUser = Mid(constr, pos, clen)
'
'      pos = InStr(1, constr, "Password=") + 9
'      If pos > 9 Then
'         clen = InStr(pos, constr, ";") - pos
'         dbPass = Mid(constr, pos, clen)
'         If dbPass = Chr(34) & Chr(34) Then
'            dbPass = ""
'         End If
'      Else
'         dbPass = ""
'      End If
'
'      pos = InStr(1, constr, "Initial Catalog") + 16
'      clen = InStr(pos, constr, ";") - pos
'      UDLDatabase = Mid(constr, pos, clen)
'      cboDB.AddItem UDLDatabase
'      cboDB.ListIndex = 0
'   Else
'      MsgBox "Specified UDL (" & UDLFile & ") not found", _
'             vbExclamation, "Invalid UDL"
'   End If
'
'   ReadUDL = constr
'End Function

Public Function ReadUDL(UDLFile As String, _
                        Optional FromfrmDatabase As Boolean = True) As String
   Dim pTag As String
   Dim pos As Long
   Dim clen As Integer
   Dim buf As String
   Dim constr As Variant
   Dim UDLConnectData As String
   
'   conStr = ""
   pTag = "P" & Chr(0) & _
          "r" & Chr(0) & _
          "o" & Chr(0) & _
          "v" & Chr(0) & _
          "i" & Chr(0) & _
          "d" & Chr(0) & _
          "e" & Chr(0) & _
          "r"
'   UniConStr = ""
   If fs.FileExists(UDLFile) Then
      buf = Space(1500)
      Open UDLFile For Binary As #1
      Get #1, , buf
      Close #1
      
      pos = InStr(1, buf, pTag, vbBinaryCompare)
      UDLConnectData = Replace(Mid(buf, pos), Chr(0), "")
      constr = UDLConnectData
      
      pos = InStr(1, constr, "Source=") + 7
      clen = InStr(pos, constr, vbCrLf) - pos
      UDLServer = Mid(constr, pos, clen)
   
'      pos = InStr(1, constr, "Source=") + 7
'      clen = InStr(pos, constr, vbCrLf) - pos
'      UDLDatabase = Mid(constr, pos, clen)
      
      pos = InStr(1, constr, "User ID=") + 8
      clen = InStr(pos, constr, ";") - pos
      dbUser = Mid(constr, pos, clen)
      
      pos = InStr(1, constr, "Password=") + 9
      If pos > 9 Then
         clen = InStr(pos, constr, ";") - pos
         dbPass = Mid(constr, pos, clen)
         If dbPass = Chr(34) & Chr(34) Then
            dbPass = ""
         End If
      Else
         dbPass = ""
      End If
   
      pos = InStr(1, constr, "Initial Catalog") + 16
      clen = InStr(pos, constr, ";") - pos
      UDLDatabase = Mid(constr, pos, clen)
      If FromfrmDatabase Then
         frmDatabase.cboDB.AddItem UDLDatabase
'         frmDatabase.cboDB.ListIndex = 0
      End If
   Else
      MsgBox "Specified UDL (" & UDLFile & ") not found", _
             vbExclamation, "Invalid UDL"
   End If
   
   ReadUDL = constr
End Function

Private Sub ShowAvailableServers()
   Dim dmoDB As New SQLDMO.Application
   Dim svrNames As SQLDMO.NameList
   Dim I As Integer
   Dim lKey As String
   
   Set svrNames = dmoDB.ListAvailableSQLServers
   lvDB.ListItems.Clear
   
   For I = 1 To svrNames.Count
      lKey = "Svr_" & I
      lvDB.ListItems.Add , lKey, svrNames.Item(I)
   Next I
   
   frmDatabase.Caption = UDLDatabase
   frmDatabase.MousePointer = vbNormal
   
   lvDB.Visible = True
   txtUser.Text = dbUser
   txtPass = dbPass
End Sub
'
'Public Property Get UDLFile() As String
'   UDLFile = UDLFile
'End Property

Public Sub ShowAvailableDatabases()
   On Error GoTo procEH
   Dim I As Integer
   
   cboDB.Clear
   With sqlServer
      If sqlServer.VerifyConnection(SQLDMOConn_CurrentState) = False Then
         .Connect CStr(lvDB.SelectedItem), txtUser.Text, txtPass.Text
      End If
      
      For I = 1 To .Databases.Count
         cboDB.AddItem .Databases(I).Name
         If .Databases(I).Name = UDLDatabase Then
            cboDB.ListIndex = I - 1
         End If
      Next I
      If cboDB.ListIndex = -1 Then
         cboDB.ListIndex = 0
      End If
   End With
   Exit Sub
      
procEH:
   Stop
   MsgBox Err.Description, vbExclamation, "Login Incorrect"
End Sub

Private Function WriteUDL()
   Dim iceFile As String
   Dim buf As String
   Dim fileHdr As String
   Dim fileTrailer As Variant
   Dim connectData As String
   
   connectData = "Provider=SQLOLEDB.1;Password=<PASS>;Persist Security Info=True;User ID=<USER>;Initial Catalog=<DB>;Data Source=<SVR>"
   If dbPass = "" Then
      connectData = Replace(connectData, "<PASS>", Chr(34) & Chr(34))
   Else
      connectData = Replace(connectData, "<PASS>", dbPass)
   End If
   
   connectData = Replace(connectData, "<USER>", dbUser)
   connectData = Replace(connectData, "<DB>", sqlDb.Name)
   If sqlServer.Name = "(local)" Then
      connectData = Replace(connectData, "<SVR>", sqlServer.TrueName)
   Else
      connectData = Replace(connectData, "<SVR>", sqlServer.Name)
   End If
   
   fileTrailer = Chr(13) & Chr(0) & Chr(10) & Chr(0)
   iceFile = fs.BuildPath(UDLPath, "ice.udl")
   
   buf = Space(1500)
   Open iceFile For Binary As #1
   Get #1, , buf
   Close #1
   
   fileHdr = Left(buf, 128) & StrConv(connectData, vbUnicode) & fileTrailer
   fs.DeleteFile iceFile
   Open iceFile For Binary As #1
   Put #1, , fileHdr
   Close #1
End Function

Private Sub lstDB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   UDLServer = lstDB.List(lstDB.ListIndex)

   If Button = vbLeftButton Then
      cmdLogin_Click
   Else
 '     lstDB.Visible = False
'      fraLogin.Visible = True
   End If
End Sub

Private Sub lvDB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo procEH
   If Button = vbLeftButton Then
      cmdLogin_Click
   Else
      lvDB.Visible = False
'      ShowAvailableDatabases
      fraLogin.Visible = True
   End If
   
AfterError:
   Exit Sub
   
procEH:
   Resume AfterError
End Sub
