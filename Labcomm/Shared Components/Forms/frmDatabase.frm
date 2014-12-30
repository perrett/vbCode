VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Selection"
   ClientHeight    =   3105
   ClientLeft      =   5940
   ClientTop       =   4920
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUDL 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   105
      TabIndex        =   6
      Top             =   105
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
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   " "
         Height          =   450
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
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
      Caption         =   "Login details"
      Height          =   2940
      Left            =   3030
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1605
         TabIndex        =   11
         Top             =   2145
         Width           =   855
      End
      Begin VB.ComboBox cboDB 
         Height          =   315
         ItemData        =   "frmDatabase.frx":030A
         Left            =   255
         List            =   "frmDatabase.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1485
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   255
         TabIndex        =   5
         Top             =   2145
         Width           =   855
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
      Begin VB.Label lblSearch 
         Alignment       =   2  'Center
         Caption         =   "Searching for databases..."
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   30
         TabIndex        =   13
         Top             =   1515
         Visible         =   0   'False
         Width           =   2595
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
      Left            =   90
      TabIndex        =   12
      Top             =   135
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Searching for available databases..."
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   75
      TabIndex        =   14
      Top             =   990
      Width           =   2790
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
Private blnUseDMO As Boolean
Private dbServer As SQLDMO.sqlServer
Private conErrStr As String

Private iceCon As New ADODB.Connection

Private Sub BuildUDLButtons()
   Dim curFile As String
   Dim UDLCount As Integer
   Dim addHeight As Long
   Dim blnAddButton As Boolean
      
   curFile = Dir(fs.BuildPath(UDLPath, "*.udl"))
   Do Until curFile = ""
      UDLConnectData = curFile & "|" & ReadUDL(fs.BuildPath(UDLPath, curFile))
      
      If UDLDatabase = "ICEHL7History" Then
         If Read_Ini_Var("GENERAL", "UseHL7db", INIFile) = 1 Then
            blnAddButton = True
         Else
            blnAddButton = False
         End If
      Else
         blnAddButton = True
      End If
      
      If blnAddButton Then
         If UDLCount >= cmdDB.Count Then
            Load cmdDB(UDLCount)
            
            addHeight = addHeight + cmdDB(UDLCount).Height + 80
            
            With cmdDB(UDLCount)
               .Top = 960 + addHeight
               .Left = 120
               .Visible = True
            End With
         End If
         
         With cmdDB(UDLCount)
            .Caption = UDLDatabase & " (" & UDLServer & ")"
            .ToolTipText = curFile
            .Tag = UDLConnectData
            
            If UDLDatabase = "ICEHL7History" Then
               .BackColor = &HFFF9CC
               .Default = False
            Else
               .BackColor = &H8000000F
               .Default = True
               .ZOrder (1)
            End If
         End With
         
         UDLCount = UDLCount + 1
      End If
      
      curFile = Dir
   Loop
   
   frmDatabase.Height = frmDatabase.Height + addHeight
   fraUDL.Height = fraUDL.Height + addHeight
   Line1.Y1 = Line1.Y1 + addHeight
   Line1.Y2 = Line1.Y2 + addHeight
   cmdExit.Top = frmDatabase.Height - 1665
   
End Sub

Public Property Let ConfigPath(strNewValue As String)
   UDLPath = strNewValue
End Property

Private Sub cboDB_Click()
   Dim pos As Integer
   
   pos = InStr(1, Me.Caption, "(Using") - 1
   If pos > 0 Then
      Me.Caption = Left(Me.Caption, pos) & "(Using " & cboDB.Text & ")"
   End If
End Sub

Private Sub xcboDB_DropDown()
   ShowAvailableDatabases
End Sub

Private Sub cmdCancel_Click()
   Dim curFile As String
   
   With cmdDB(dbButton)
      curFile = fs.BuildPath(UDLPath, .ToolTipText)
      UDLConnectData = ReadUDL(curFile)
      .Caption = UDLDatabase & " (" & UDLServer & ")"
      .Tag = UDLConnectData
   End With
   
   ResizeForm False
End Sub

Private Sub cmdDB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim connectData As Variant
   Dim curFile() As String
   
   dbButton = Index
   curFile = Split(cmdDB(Index).Tag, "|")
   UDLFile = fs.BuildPath(UDLPath, curFile(0))
   AnalyseConnectionString curFile(1)
      
   If RunningInIDE Then
      
      If Button = vbRightButton Then
         Me.Caption = "Select Server"
         Me.MousePointer = vbHourglass
         fraUDL.Visible = False
         
         DoEvents
         ShowAvailableServers
         DoEvents
         cmdLogin.Default = True
         Me.MousePointer = vbNormal
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
   
   If cmdLogin.Caption = "Login" Then
      dbServer.EnableBcp = True
      Set sqlServer = dbServer
      dbUser = txtUser.Text
      dbPass = txtPass.Text
      UDLServer = lvDB.SelectedItem.Text
'      lblSearch.Top = 1500
      cmdLogin.Enabled = False
'      cmdLogin.Top = 2200
      cmdCancel.Enabled = False
'      cmdCancel.Top = 2200
      
      lblSearch.Caption = "Establishing connection to: " & UDLServer
      
      DoEvents
      
      If sqlServer.Name <> UDLServer Then
         sqlServer.Disconnect
      End If
      
      If sqlServer.VerifyConnection(SQLDMOConn_CurrentState) Then
         txtPass.Text = "supplied"
      Else
         sqlServer.Connect UDLServer, dbUser, dbPass
      End If
            
      frmDatabase.Caption = frmDatabase.Caption & " - Connected"
      lblSearch.Caption = "Listing available databases..."
      
      DoEvents
      
      ShowAvailableDatabases
      
      lblSearch.Caption = "Select required database"
      cboDB.Visible = True
      cmdLogin.Caption = "Select"
      
   Else
      ResizeForm False
      UDLServer = lvDB.SelectedItem.Text
      If cboDB.Text <> "" Then
         UDLDatabase = cboDB.Text
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
      cmdLogin.Caption = "Login"
      cmdDB(0).Default = True
      
   End If
   
AfterError:
   cmdLogin.Enabled = True
   cmdCancel.Enabled = True
   Exit Sub
   
procEH:
   MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Unable to access database"
   Resume AfterError
End Sub

Public Function ConnectionDetails(Optional ByRef ServerId As SQLDMO.sqlServer = Nothing, _
                                  Optional ByRef dbId As SQLDMO.Database = Nothing, _
                                  Optional ByRef sqlUser As String = "", _
                                  Optional ByRef sqlPass As String = "", _
                                  Optional ByRef UDLsvrId As Variant = Nothing, _
                                  Optional ByRef UDLdb As String = "", _
                                  Optional UDLToUse As String = "ice.udl") As ADODB.Connection
   
   On Error GoTo procEH
   Dim dbCon As New ADODB.Connection
   Dim iceCmd As New ADODB.Command
   Dim dbTime As String
   Dim iniDB As String
   Dim timeOut As Long
   Dim cVers As Integer
   Dim dbVers As Integer
   Dim rVal As Integer
      
   UDLPath = Read_Ini_Var("General", "ConfigPath", INIFile)
   
   If UDLPath = "" Then
      UDLPath = App.Path
   ElseIf UCase(fs.GetExtensionName(UDLPath)) = "UDL" Then
      UDLPath = fs.GetParentFolderName(UDLPath)
   End If

   UDLFile = fs.BuildPath(UDLPath, UDLToUse)
   
   If fs.FileExists(UDLFile) Then
      dbTime = Read_Ini_Var("General", "dbTimeOut", INIFile)
      If dbTime = "" Then
         timeOut = 10
      Else
         timeOut = CLng(dbTime)
      End If
   
      Initialize
            
      With dbCon
         .CursorLocation = adUseClient
         .Mode = adModeReadWrite
         .CommandTimeout = timeOut
         .ConnectionTimeout = timeOut
         .Open "FILE NAME=" & UDLFile
      End With
      
      If Not dbServer Is Nothing Then
         Set ServerId = dbServer
         ServerId.ApplicationName = "IceLabcommSuite"
         ServerId.EnableBcp = True
         
         If ServerId.VerifyConnection(SQLDMOConn_CurrentState) = False Then
            ServerId.Connect UDLServer, dbUser, dbPass
         End If
         
         Set dbId = ServerId.Databases(UDLDatabase)
'            Set ServerId = sqlServer
      
'            Set dbId = sqlDb
'            sqlUser = dbUser
'            sqlPass = dbPass
'            UDLsvrId = UDLServer
'            UDLdb = UDLDatabase
      End If
      
      cVers = App.Minor
      
      iniDB = Read_Ini_Var("General", "dbVersion", INIFile)
      If iniDB <> "" Then
         If IsNumeric(iniDB) Then
            If cVers < Val(iniDB) Then
               cVers = Val(iniDB)
            End If
         End If
      End If
      
      With iceCmd
         .ActiveConnection = dbCon
         .CommandType = adCmdStoredProc
         .CommandText = "IceLabcomm_Check_dbVersion"
         .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
         .Parameters.Append .CreateParameter("codeVersion", adInteger, adParamInput, , cVers)
         .Parameters.Append .CreateParameter("dbVersion", adInteger, adParamOutput, , dbVers)
         .Execute
         
         DBVERSION = .Parameters("dbVersion").Value
         dbVers = .Parameters("Return").Value
      End With
      
      If dbVers > 0 Then
         MsgBox "Code version: " & cVers & " will not run against Database version " & dbVers, vbCritical, "Database Version Check"
         End
      End If
      
   Else
      MsgBox "Specified UDL file in title bar does not exist", vbCritical, UDLFile
      End
   End If
   
   Set ConnectionDetails = dbCon
   Unload Me
   Exit Function
   
procEH:
   If Err.Number = -2147217900 Then
      If App.Minor < 340 Then
         Set ConnectionDetails = dbCon
      End If
   Else
      conErrStr = Err.Number & ": " & Err.Description & " ( " & Err.Source & ")"
      Set ConnectionDetails = Nothing
   End If
End Function

Public Property Get ConnectionError() As String
   ConnectionError = conErrStr
End Property

Public Function DMOAvailable() As SQLDMO.sqlServer
   On Error GoTo procEH
   
   Set DMOAvailable = CreateObject("SQLDMO.sqlServer")
   
   Exit Function
   
procEH:
   Set DMOAvailable = Nothing
End Function

Public Property Get DMOStatus() As Boolean
   DMOStatus = Not (dbServer Is Nothing)
End Property

Private Sub Initialize()
   Dim curFile As String
   Dim defUDL As String
   Dim UDLCount As Integer
   Dim dbName As String
   Dim strArray() As String
   
   If (Read_Ini_Var("UDLData", "MultiUDL", INIFile) = 1) Or _
      RunningInIDE Then
      BuildUDLButtons
      Me.Show 1
   Else
      'ReadUDL UDLFile
      AnalyseConnectionString ReadUDL(UDLFile)
   End If
End Sub

Private Sub Form_Load()
   cboDB.Visible = False
   ResizeForm False
   Set dbServer = DMOAvailable
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

Private Sub AnalyseConnectionString(constr As String)
   Dim pTag As String
   Dim pos As Long
   Dim clen As Integer
   Dim buf As String
   Dim UDLConnectData As String
   
   pos = InStr(1, constr, "Source=") + 7
   clen = InStr(pos, constr, vbCrLf) - pos
   UDLServer = Mid(constr, pos, clen)
      
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
   'UDLDatabase = Mid(constr, pos, clen)
End Sub

Public Function ReadUDL(UDLFile As String, _
                        Optional FromfrmDatabase As Boolean = True) As String
   Dim pTag As String
   Dim pos As Long
   Dim clen As Integer
   Dim buf As String
   Dim constr As Variant
   Dim UDLConnectData As String
   
   pTag = "P" & Chr(0) & _
          "r" & Chr(0) & _
          "o" & Chr(0) & _
          "v" & Chr(0) & _
          "i" & Chr(0) & _
          "d" & Chr(0) & _
          "e" & Chr(0) & _
          "r"
   
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
      
      pos = InStr(1, constr, "Initial Catalog") + 16
      clen = InStr(pos, constr, ";") - pos
      UDLDatabase = Mid(constr, pos, clen)
      If FromfrmDatabase Then
         frmDatabase.cboDB.AddItem UDLDatabase
      End If
   Else
      MsgBox "Specified UDL (" & UDLFile & ") not found", _
             vbExclamation, "Invalid UDL"
   End If
   
   ReadUDL = constr
End Function

Private Sub ResizeForm(ShowLogin As Boolean)
   If ShowLogin Then
      fraUDL.Visible = False
      frmDatabase.Width = 5895
      cboDB.Visible = False
      
      lblSearch.Caption = "Please log in..."
      lblSearch.Visible = True
'      cmdLogin.Top = 1530
'      cmdCancel.Top = 1530
      fraLogin.Visible = True
   Else
      frmDatabase.Width = 2925
'      lblSearch.Top = 2340
      lblSearch.Caption = "Searching for servers..."
'      cmdLogin.Top = 2145
'      cmdCancel.Top = 2145
      cboDB.Visible = False
      lvDB.Visible = False
      fraLogin.Visible = False
      fraUDL.Visible = True
   End If
End Sub

Private Sub ShowAvailableServers()
   Dim dmoDB As SQLDMO.Application
   Dim svrNames As SQLDMO.NameList
   Dim I As Integer
   Dim lKey As String
   
   Set dmoDB = CreateObject("SQLDMO.Application")
   
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
   Exit Sub
End Sub

Public Sub ShowAvailableDatabases()
   On Error GoTo procEH
   Dim I As Integer
   Dim defDb As Integer
   
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
         If .Databases(I).Name = .Logins(txtUser.Text).Database Then
            defDb = I - 1
         End If
      Next I
      If cboDB.ListIndex = -1 Then
         cboDB.ListIndex = defDb
      End If
   End With
   
   frmDatabase.Caption = frmDatabase.Caption & " (Using " & cboDB.Text & ")"
   Exit Sub
      
procEH:
   Stop
   MsgBox Err.Description, vbExclamation, "Login Incorrect"
End Sub

Private Function WriteUDL()
   Dim iceFile As String
   Dim buf As String
   Dim fileHdr As String
   Dim FileTrailer As Variant
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
   
   FileTrailer = Chr(13) & Chr(0) & Chr(10) & Chr(0)
   'iceFile = fs.BuildPath(UDLPath, "ice.udl")
   
   buf = Space(1500)
   Open UDLFile For Binary As #1
   Get #1, , buf
   Close #1
   
   fileHdr = Left(buf, 128) & StrConv(connectData, vbUnicode) & FileTrailer
   fs.DeleteFile UDLFile
   Open UDLFile For Binary As #1
   Put #1, , fileHdr
   Close #1
End Function

Private Sub lvDB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo procEH
   If lvDB.SelectedItem.Text = "Peterboro" Then
      txtUser.Text = "anglia"
   Else
      txtUser = "sa"
   End If
   
   txtPass.Text = ""

   frmDatabase.Caption = "Server: " & lvDB.SelectedItem.Text
   ResizeForm True
   cmdLogin.Caption = "Login"
   txtPass.SetFocus

AfterError:
   Exit Sub
   
procEH:
   Resume AfterError
End Sub
