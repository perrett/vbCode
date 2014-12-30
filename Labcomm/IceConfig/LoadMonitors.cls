VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadMonitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private nd(2) As Node
Private tv1 As TreeView
Private vData As Variant
Private mCtrl As New ManageControls
'Private Activity() As String
'Private ActivityDirs() As String
'Private ActivityDirPattern() As String
Private curParam  As String
Private statCaption As String
Private curNodeType As String
Private CurrentModule As String
Private CurrentInFlight As String
Private CurrentViewOption As Integer
Private IceCmdText As String
Private rstFile As String
Private LastLVItem As MSComctlLib.ListItem
Private dirStr As String
Private WithEvents iceAsyncCon As ADODB.Connection
Attribute iceAsyncCon.VB_VarHelpID = -1
Private iceAsyncCmd As ADODB.Command

Private blnExecuteComplete As Boolean

Private commRS As New ADODB.Recordset
Private invRS As New ADODB.Recordset
Private resRS As New ADODB.Recordset
Private pracRS As New ADODB.Recordset
Private fxdRS As New ADODB.Recordset
Attribute fxdRS.VB_VarHelpID = -1
Private spRS As New ADODB.Recordset
Private wBrowse As New HTMLConvert
Private blob As New clsDBBlob
Private iconId As Integer
'Private wbEv As New browserControl

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim iniData As String
   
   Set tv1 = frmMain.TreeView1
   'fView.FrameToShow = Fra_FILEDETAILS
   frmMain.DTStats.value = Now()
   frmMain.DTStatEnd.MaxDate = Now()
   frmMain.DTStatEnd.value = frmMain.DTStats.value - 7
   'frmMain.optStatPeriod(1).value = True
   
   iniData = Read_Ini_Var("General", "dbTimeOut", iniFile)
   
   If iniData = "" Or IsNumeric(iniData) = False Then
      iniData = "10"
   End If
   
   Set iceAsyncCon = New ADODB.Connection
   iceAsyncCon.CommandTimeout = Val(iniData)
   iceAsyncCon.ConnectionTimeout = Val(iniData)
   iceAsyncCon.ConnectionString = iceCon.ConnectionString
   iceAsyncCon.Open
   
   blnExecuteComplete = True
   
   With frmWait
      .Label1.Caption = "Please wait - preparing Statistics package"
      .ProgressBar1.Visible = True
      .ProgressBar1.max = 3
      .ProgressBar1.value = 0
      .Show
      .Refresh
   End With
   
   With frmMain
      .Refresh
      .MousePointer = vbHourglass
   End With
   
   strSQL = "SELECT EDI_NatCode, EDI_Name " & _
            "FROM EDI_Recipients"

   pracRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   With frmMain.cboStatsPractice
      .Clear
      .AddItem "All practices"
      
      Do Until pracRS.EOF
         .AddItem pracRS!EDI_NatCode
         pracRS.MoveNext
      Loop
      .ListIndex = 0
   End With
   
   frmWait.ProgressBar1.value = 2
   Set pracRS.ActiveConnection = Nothing
   blob.DBConnection = iceCon
   Unload frmWait
   frmMain.MousePointer = vbNormal
   
   With frmMain.wb
      .DBConnection = iceCon
      .BrowserToolTip = ""
      .LocationToolTip = ""
   End With
   iconId = 12
   
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMonitors.Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Activities(NodeId As Node)
   On Error GoTo procEH
   Dim dStr As String
   
   eClass.FurtherInfo = "System Monitor"
'   fView.FrameToShow = Fra_TESTDETAILS
   vData = objTV.ReadNodeData(NodeId)
   Set LastLVItem = Nothing
   
   With frmMain.wb
      If vData(0) = "VERSION" Then
         .NavigationBar = 0
         .BrowserToolTip = ""
         .InfoBGColour = "#ffffcc"
         .Script = blob.Read("NotSupported", TT_Script)
         .Style = blob.Read("NotSupported", TT_Style)
         .NavigateTo blob.Read("NotSupported", TT_HTML)
      ElseIf vData(0) = "FILE" Then
         .InfoCallBack = Me
         .LocationTitle = vData(1)
         
         If vData(0) = "FILE" Then
            frmMain.wb.NavigateTo vData(1)
'            frmMain.wb.ShowDirInfo CStr(vData(1))
         Else
            frmMain.wb.ShowDirInfo CStr(vData(1))
         End If
      End If
   End With
   frmMain.wb.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMonitors.Activities"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function CallBack(Action As String, Optional Params As Variant) As String
   Dim fBuf As String
   Dim sBuf As New StringBuffer
   Dim bodyHdr As String
   Dim qt As String
   
   qt = Chr(34)
   Select Case Action
      Case "FileToNotepad"
         OpenInNotepad CStr(Params)
      
      Case "ReadFile"
         frmMain.wb.LocationTitle = fs.GetFileName(Params)
         frmMain.wb.Tag = Params
         
         Open Params For Input As #1
         While Not EOF(1)
            Line Input #1, fBuf
            sBuf.Append fBuf
            If Len(fBuf) > 0 Then
               sBuf.Append vbCrLf
            End If
         Wend
         Close #1
         fBuf = Replace(sBuf.ActualValue, "?'", "#apos#")
         fBuf = Replace(fBuf, "'", "</TR><TR>")
         fBuf = Replace(fBuf, "#apos#", "'")
         
         bodyHdr = "<div NOWRAP>" & vbCrLf & "<font face=" & qt & "Courier New" & qt & " size=" & qt & "2" & qt & ">" & vbCrLf & "<Table Width=100%><TR>"
         CallBack = bodyHdr & fBuf & "</TR></table></Font></Div>"
         
      Case "LocationFileName"
         CallBack = frmMain.wb.Tag
         
   End Select
End Function

Public Sub FirstView()
   On Local Error GoTo procEH
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("ACTIVITY", _
                                              "NONE", _
                                              "Details"), _
                             "Labcomm folders", _
                             iconId, _
                             iconId)
   
   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", _
                                  "Tmp"), _
                 "Please wait...", _
                 iconId, _
                 iconId
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("STATS", _
                                              "NONE", _
                                              "Details"), _
                             "Statistics", _
                             iconId, _
                             iconId)
   
   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", _
                                  "Tmp"), _
                 "Please wait...", _
                 iconId, _
                 iconId
   
'   Set nd(1) = tv1.Nodes.Add(nd(0), _
                             tvwChild, _
                             mCtrl.NewNodeKey("INVS", _
                                              "NONE", _
                                              "Statistics"), _
                             "Investigation Statistics", _
                             iconid, _
                             iconid)
   
'   Set nd(1) = tv1.Nodes.Add(nd(0), _
                             tvwChild, _
                             mCtrl.NewNodeKey("RESS", _
                                              "RES", _
                                              "Statistics"), _
                             "Result Statistics", _
                             iconid, _
                             iconid)
   
'   Set nd(1) = tv1.Nodes.Add(nd(0), _
                             tvwChild, _
                             mCtrl.NewNodeKey("MESSAGE", _
                                              "MSG", _
                                              "Details"), _
                             "Message Statistics", _
                             iconid, _
                             iconid)
   
'   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", _
                                  "Tmp"), _
                 "Please wait...", _
                 iconid, _
                 iconid
   
   
'   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("INVS", _
                                              "INV", _
                                              "Statistics"), _
                             "Investigation Statistics", _
                             iconid, _
                             iconid)
   
'   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", _
                                  "Tmp"), _
                 "Please wait...", _
                 iconid, _
                 iconid
                 
'   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("RESS", _
                                              "RES", _
                                              "Statistics"), _
                             "Result Statistics", _
                             iconid, _
                             iconid)
   
'   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", _
                                  "Tmp"), _
                 "Please wait...", _
                 iconid, _
                 iconid
                 
'   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("MESSAGE", _
                                              "MSG", _
                                              "Details"), _
                             "Message Statistics", _
                             iconid, _
                             iconid)
   
'   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", _
                                  "Tmp"), _
                 "Please wait...", _
                 iconid, _
                 iconid
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("SEARCH", _
                                              "NONE", _
                                              "Details"), _
                             "Locate reports using codes", _
                             iconId, _
                             iconId)
   
   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", _
                                  "Tmp"), _
                 "Please wait...", _
                 iconId, _
                 iconId
   
   'fView.FrameToShow = fra_None
   
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMonitors.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
   Dim i As Integer
   Dim inDir As String
   Dim fileMask As String
   Dim curFile As String
   Dim strProc As String
   Dim nodeDesc As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
      
   vData = objTV.ReadNodeData(NodeId)
   Select Case vData(0)
      Case "ACTIVITY"
'         fView.FrameToShow = fra_None
'         fView.FrameToShow = Fra_FILEDETAILS
         strSQL = "SELECT Module_Name, " & _
                     "cn.Connection_Name, " & _
                     "Connection_InFlightMapping, " & _
                     "Connection_CollectHow, " & _
                     "Connection_TargetDirectory, " & _
                     "Connection_TargetFilemap, " & _
                     "Connection_HistoryDirs, " & _
                     "Connection_ErrorDirs, " & _
                     "Connection_LogDirs " & _
                  "FROM Connections Cn " & _
                     "INNER JOIN Connect_Modules CM " & _
                     "ON Cn.Connection_Name = CM.Connection_Name " & _
                  "WHERE (Connection_CollectHow is not null " & _
                        "AND rtrim(Connection_CollectHow) <> '') " & _
                     "AND Connection_TargetDirectory is not null " & _
                  "ORDER BY Module_Name"
         
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         Do While Not RS.EOF
            nodeDesc = ""
            Select Case UCase(RS!Module_Name)
               Case "ICEMSG.EXE"
                  If UCase(RS!Connection_InFlightMapping) = "ACKS" Then
                     nodeDesc = "Message Acknowledgments"
                  Else
                     nodeDesc = "Messaging"
                  End If
                  
               Case "ICEIMP.EXE"
                  nodeDesc = "Importer"
                  
            End Select
                  
            If nodeDesc <> "" Then
               Set nd(0) = tv1.Nodes.Add(NodeId, _
                                         tvwChild, _
                                         mCtrl.NewNodeKey(RS!Module_Name, _
                                                          RS!Connection_Name, _
                                                          "None"), _
                                         nodeDesc, _
                                         iconId, _
                                         iconId)
            
               If Trim(RS!Connection_CollectHow & "") <> "" Then
                  inDir = fs.GetParentFolderName(Trim(RS!Connection_CollectHow))
                  fileMask = fs.GetFileName(Trim(RS!Connection_CollectHow))
                  tv1.Nodes.Add nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey("FILE", _
                                                 inDir, _
                                                 "Activity", _
                                                 fileMask), _
                                "Input Directory - " & inDir, _
                                iconId, _
                                iconId
               End If
                  
               If Trim(RS!Connection_TargetDirectory & "") <> "" Then
                  tv1.Nodes.Add nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey("FILE", _
                                                 Trim(RS!Connection_TargetDirectory), _
                                                 "Activity", _
                                                 "*.*"), _
                                "Output Directory - " & Trim(RS!Connection_TargetDirectory), _
                                iconId, _
                                iconId
               End If
               
               If Trim(RS!Connection_ErrorDirs & "") <> "" Then
                  tv1.Nodes.Add nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey("FILE", _
                                                 Trim(RS!Connection_ErrorDirs), _
                                                 "Activity", _
                                                 "*.*"), _
                                "Error Directory - " & Trim(RS!Connection_ErrorDirs), _
                                iconId, _
                                iconId
               End If
               
               If Trim(RS!Connection_HistoryDirs & "") <> "" Then
                  Set nd(1) = tv1.Nodes.Add(nd(0), _
                                            tvwChild, _
                                            mCtrl.NewNodeKey("FOLDER", _
                                                             Trim(RS!Connection_HistoryDirs), _
                                                             "None"), _
                                            "History Directories - " & Trim(RS!Connection_HistoryDirs), _
                                            iconId, _
                                            iconId)
                  
                  curFile = Dir(fs.BuildPath(RS!Connection_HistoryDirs, "*.*"), vbDirectory)
                  
                  Do Until curFile = ""
                     If (curFile <> "." And curFile <> "..") And _
                        fs.FolderExists(fs.BuildPath(RS!Connection_HistoryDirs, curFile)) Then
                        tv1.Nodes.Add nd(1), _
                                      tvwChild, _
                                      mCtrl.NewNodeKey("FILE", _
                                                       fs.BuildPath(RS!Connection_HistoryDirs, curFile), _
                                                       "Activity", _
                                                       "*.*"), _
                                      curFile, _
                                      iconId
                     End If
                     curFile = Dir
                  Loop
               End If
               
               If Trim(RS!Connection_LogDirs & "") <> "" Then
                  tv1.Nodes.Add nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey("FILE", _
                                                 Trim(RS!Connection_LogDirs), _
                                                 "Activity", _
                                                 "*.log"), _
                                "Log Directory - " & Trim(RS!Connection_LogDirs), _
                                iconId, _
                                iconId
               End If
   '            RS2.Close
            End If
            RS.MoveNext
         Loop
      
         RS.Close
         
      Case "STATS"
         fView.Show Fra_TESTDETAILS
         
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey("INVS", _
                                                    "NONE", _
                                                    "Statistics"), _
                                   "Investigations Performed", _
                                   iconId, _
                                   iconId)
         
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey("RESS", _
                                                    "RES", _
                                                    "Statistics"), _
                                   "Result Dispatched", _
                                   iconId, _
                                   iconId)
         
         strSQL = "SELECT Description, Status_Flag " & _
                  "FROM EDI_Msg_Formats " & _
                  "WHERE Active = 1"
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey("MESSAGE", _
                                                    "MSG", _
                                                    "None"), _
                                   "Message Statistics", _
                                   iconId, _
                                   iconId)
                                      
         tv1.Nodes.Add nd(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey("MSGI", _
                                        "0", _
                                        "Statistics"), _
                       "Messages In", _
                       iconId, _
                       iconId
         
         Do Until RS.EOF
            tv1.Nodes.Add nd(0), _
                          tvwChild, _
                          mCtrl.NewNodeKey("MSGP", _
                                           RS!Status_Flag, _
                                           "Statistics"), _
                          RS!Description, _
                          iconId, _
                          iconId
            
            RS.MoveNext
         Loop
         RS.Close
      
         strSQL = "SELECT so.name, so.Id, count(sc.name) as NoParams " & _
                  "FROM sysobjects so " & _
                     "INNER JOIN syscolumns sc " & _
                     "ON so.id=sc.id " & _
                  "WHERE so.name like 'ICECONFIG_Statsx%' " & _
                  "GROUP BY so.name, so.Id"
         
         'strSQL = "SELECT name " & _
                  "FROM sysobjects " & _
                  "WHERE name like 'IceConfig_StatsX%'"
         
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         
         If RS.RecordCount > 0 Then
            Set nd(0) = tv1.Nodes.Add(NodeId, _
                                      tvwChild, _
                                      mCtrl.NewNodeKey("EXTRA", _
                                                       "EX", _
                                                       "None"), _
                                      "Extra statistics", _
                                      iconId, _
                                      iconId)
            Do Until RS.EOF
               strProc = Mid(RS(0), 18)
               
               Set nd(1) = tv1.Nodes.Add(nd(0), _
                                         tvwChild, _
                                         mCtrl.NewNodeKey(strProc, _
                                                          IIf(RS!NoParams > 2, RS!id, "0"), _
                                                          "CustomStats"), _
                                         strProc, _
                                         iconId, _
                                         iconId)
               RS.MoveNext
            Loop
            RS.Close
         End If
         
      Case "MESSAGE"
         strSQL = "SELECT Description, Status_Flag " & _
                  "FROM EDI_Msg_Formats " & _
                  "WHERE Active = 1"
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         
            Do Until RS.EOF
               tv1.Nodes.Add NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey("MSGP", _
                                              RS!Status_Flag, _
                                              "Statistics"), _
                             RS!Description, _
                             iconId, _
                             iconId
               
               RS.MoveNext
            Loop
            RS.Close
         
      Case "SEARCH"
         fView.Show Fra_TESTDETAILS
         
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey("TEST", _
                                                    "Find", _
                                                   "Statistics"), _
                                   "Find last test usage...", _
                                   iconId, _
                                   iconId)
                  
   
   End Select
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMonitors.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
'
'Private Sub Display_DirActivity()
'   On Error GoTo procEH
'   Dim i As Integer
'   Dim FL As Integer
'
'   eClass.FurtherInfo = "Displaying System Activity directory(ies)"
'   FL = 0
'   For i = 0 To UBound(ActivityDirs)
'      If (ActivityDirs(i) <> "") And (FL <= frmMain.FLActivity.UBound) Then
'         frmMain.FLlbl(FL).Caption = Activity(i)
'         frmMain.FLActivity(FL).Path = ActivityDirs(i)
'         frmMain.FLActivity(FL).Pattern = ActivityDirPattern(i)
'         If i <= frmMain.FLlbl.UBound Then
'            frmMain.FLlbl(FL).Visible = True
'         End If
'         frmMain.FLActivity(FL).Visible = True
'         FL = FL + 1
'      End If
'   Next i
'   frmMain.Timer2.Enabled = True
'   Exit Sub
'
'procEH:
'   If eClass.Behaviour = -1 Then
'      Stop
'      Resume
'   Else
'      eClass.CurrentProcedure = "frmMain.Display_DirActivity"
'      eClass.Add Err.Number, Err.Description, Err.Source
'   End If
'End Sub

Public Sub GraphData(RepData As Variant, _
                     AckData As Variant, _
                     ErrData As Variant)

   Dim vData() As Variant
'   With frmMain.chartStat
''      .ColumnCount = 3
''     .SeriesType = VtChSeriesType2dBar
'      If UBound(RepData) > 0 Then
'         .ChartData = RepData
'      End If
'   End With
   
End Sub

Public Sub ListViewItemClick(ListViewId As String, _
                             item As MSComctlLib.ListItem)
   On Error GoTo procEH
   Dim fileSpec As String
   Dim curFile As String
   Dim dFile As String
   Dim tFile As String
   Dim fStream As TextStream
   Dim Tstr As String
   Dim tBuf As New StringBuffer
   Dim iconId As Long
   Dim thisFile As File
   
   vData = objTV.ReadNodeData
   
   With frmMain
      Select Case ListViewId
         Case "Dirs"
'            If vData(0) = "LOG" Then
'               fileSpec = "*.log"
'            ElseIf vData(0) = "IN" Then
'               fileSpec = vData(4)
'            Else
'               fileSpec = "*.*"
'            End If
'
'            dirStr = fs.BuildPath(.lViewDirs.Tag, item.Text)
'            .lblDir.Caption = .lViewDirs.Tag
'            .lblFiles.Caption = dirStr
'            If Not LastLVItem Is Nothing Then
'               LastLVItem.Icon = 2
'            End If
'            Set LastLVItem = item
'            item.Icon = 1
'            curFile = Dir(fs.BuildPath(dirStr, fileSpec))
'            .lViewFiles.ListItems.Clear
'            Do Until curFile = ""
'               If curFile <> "." Or curFile <> ".." Then
''                  If fs.FolderExists(curFile) Then
'                     If UCase(fs.GetExtensionName(curFile)) = "LOG" Then
'                        iconId = 33
'                     Else
'                        iconId = 32
'                     End If
'                     .lViewFiles.ListItems.Add , , curFile, , iconId
'                     Set thisFile = fs.GetFile(fs.BuildPath(dirStr, curFile))
'                     .lViewFiles.ListItems(.lViewFiles.ListItems.Count).ListSubItems.Add , , Format(thisFile.DateCreated, "dd/mm/yyyy hh:nn:ss")
'                     .lViewFiles.ListItems(.lViewFiles.ListItems.Count).ListSubItems.Add , , thisFile.Size
''                  End If
'                  curFile = Dir
'               End If
'            Loop
         
         Case "Files"
            vData = objTV.ReadNodeData
            If vData(0) = "LOG" Then
               wBrowse.ReadLogFile fs.BuildPath(dirStr, item)
            Else
               wBrowse.PrepareOutput fs.BuildPath(dirStr, item)
            End If
            
            frmShowLog.LogFileName = wBrowse.HTMLfile
            frmShowLog.Show 1
            
'            If vData(0) = "LOG" Then
'               frmShowLog.LogFileName = fs.BuildPath(dirStr, item)
'               frmShowLog.Show 1
'            Else
'               .lblFiles.Caption = dirStr
'   '            tStr = fs.BuildPath(.lViewDirs.Tag, LastLVItem.Text)
'               dFile = fs.BuildPath(dirStr, item.Text)
'               tFile = fs.BuildPath(App.Path, fs.GetTempName)
'               Open dFile For Input As #1
'               Set fStream = fs.CreateTextFile(tFile)
'               fStream.WriteLine "Viewing file: " & dFile
'               fStream.WriteLine String(14 + Len(dFile), "=")
'               While Not EOF(1)
'                  Line Input #1, Tstr
'                  tBuf.Append Tstr
'                  If Len(Tstr) > 0 Then
'                     tBuf.Append vbCrLf
'                  End If
'               Wend
'               fStream.Write tBuf.value
'               fStream.Close
'               Close #1
'               Shell "Notepad.EXE " & tFile, vbNormalFocus '  fraPanel(5).Caption, vbNormalFocus
'            End If
         
      End Select
   End With
   Exit Sub
   
procEH:
   If Err.Number = 35605 Then
      Resume Next
   End If
   
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmMain.lViewDirs_ItemClick"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub PrepareHTML(SrcData As String, Optional param As String = "")
   On Error GoTo procEH
   Dim htmlHdr As String
   Dim htmlRow As String
   Dim htmlEnd As String
   Dim sBuf As New StringBuffer
   Dim i As Integer
   Dim fStart As Integer
   Dim fEnd As Integer
   Dim alignStr As String
   Dim alignFlag As Long
   Dim borderFlag As Long
   
   htmlEnd = "</table>" & vbCrLf
   
   Select Case SrcData
      Case "INVS"
         alignStr = "LLR"
         With frmMain.wb
            .Style = "stats_search"
         End With
      
      Case "RESS"
         alignStr = "LLRR"
         With frmMain.wb
            .Style = "stats_search"
         End With
      
      Case "MSGI"
         With frmMain.wb
            .Style = "stats_search"
         End With
         
      
      Case "TEST"
'         htmlHdr = "<table cellspacing=1 cellpadding=2 border-collapse:collapse align=left>" & vbCrLf & _
                   "  <tr class=hdr>" & vbCrLf & _
                   "    <td>Dated</td>" & vbCrLf & _
                   "    <td align=cente>Report Identifier</td>" & vbCrLf & _
                   "    <td align=centre>Sent to Practice</td>" & vbCrLf & _
                   "  </tr>" & vbCrLf
         
         borderFlag = 7
         With frmMain.wb
            .Style = "stats_search"
         End With
         
      Case "MSGP"
         htmlHdr = "<table cellspacing=1 cellpadding=2 border-collapse:collapse>" & vbCrLf & _
                   "  <tr class=hdr>" & vbCrLf & _
                   "    <td>&nbsp;</td>" & vbCrLf & _
                   "    <td align=center colspan=2>Total</td>" & vbCrLf & _
                   "    <td align=center colspan=2>Successful</td>" & vbCrLf & _
                   "    <td align=center colspan=2>Awaiting Acks</td>" & vbCrLf & _
                   "    <td align=center colspan=2>Rejected</td>" & vbCrLf & _
                   "    <td>Files</td>" & vbCrLf & _
                   "  </tr>" & vbCrLf & _
                   "  <tr class=hdr>" & vbCrLf & _
                   "    <td>Practice</td>" & vbCrLf & _
                   "    <td>Files</td>" & vbCrLf & _
                   "    <td>Reports</td>" & vbCrLf & _
                   "    <td>Files</td>" & vbCrLf & _
                   "    <td>Reports</td>" & vbCrLf & _
                   "    <td>Files</td>" & vbCrLf & _
                   "    <td>Reports</td>" & vbCrLf & _
                   "    <td>Files</td>" & vbCrLf & _
                   "    <td>Reports</td>" & vbCrLf & _
                   "    <td>Resent</td>" & vbCrLf & _
                   "  </tr>" & vbCrLf
         
         borderFlag = &HE
         With frmMain.wb
            .Style = "stats_msg"
         End With
   
   End Select

   With fxdRS
      If htmlHdr = "" Then
         htmlHdr = "<table cellspacing=0 cellpadding=2 border-collapse:collapse>" & vbCrLf & _
                   "  <tr class=hdr>" & vbCrLf
         For i = 0 To .Fields.Count - 1
            htmlHdr = htmlHdr & "      <td>" & .Fields(i).Name & "</td>" & vbCrLf
         Next i
         
         htmlHdr = htmlHdr & "   </tr>" & vbCrLf
      End If
      
      If alignStr = "" Then
         alignStr = String(.Fields.Count, "L")
      End If
      
      sBuf.Clear
      sBuf.Append htmlHdr
      
      Do Until .EOF
         For i = 0 To .Fields.Count - 1
            htmlRow = htmlRow & "    <td"
            htmlRow = htmlRow & " align="
            Select Case Mid(alignStr, i + 1, 1)
               Case "R"
                  htmlRow = htmlRow & "right"
               
               Case "C"
                  htmlRow = htmlRow & "center"
               
               Case Else
                  htmlRow = htmlRow & "left"
                  
            End Select
            
            If (borderFlag And (2 ^ i)) = (2 ^ i) Then
               htmlRow = htmlRow & " id=br"
            End If
            htmlRow = htmlRow & ">" & .Fields(i).value & "</td>" & vbCrLf
         Next i
         
         htmlRow = htmlRow & "  </tr>" & vbCrLf
         .MoveNext
         sBuf.Append htmlRow
         htmlRow = ""
      Loop
   End With
   
   With frmMain.wb
      .Visible = True
      .Style = "stats_msg"
      .NavigateTo sBuf.ActualValue & "</table>" & vbCrLf
      .Visible = True
   End With
   
   Exit Sub
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMonitors.PrepareHTML"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub ReadResultRS()
   If fs.FileExists(rstFile) = False Then
      strSQL = "SELECT DISTINCT Result_Code, Result_Rubric " & _
               "FROM Service_Results " & _
               "ORDER BY Result_Rubric"
      resRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      Set resRS.ActiveConnection = Nothing
      resRS.Save rstFile, adPersistADTG
   Else
      resRS.Open rstFile
   End If
End Sub

Public Function Refresh() As String
   objTV.RefreshNode = objTV.ActiveNode
End Function

Public Function RunStoredProc(InvOrSvc As String, _
                              Optional nCode As String = "") As ADODB.Recordset
   On Error GoTo procEH
   Dim icecmd As New ADODB.Command
   
'   vData = objTV.ReadNodeData(NodeId)
   With icecmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = IceCmdText
      .Parameters.Append .CreateParameter("fromDate", adVarChar, adParamInput, 10, Format(frmMain.DTStatEnd.value, "yyyymmdd"))
      .Parameters.Append .CreateParameter("toDate", adVarChar, adParamInput, 10, Format(frmMain.DTStats.value, "yyyymmdd"))
      
      Select Case IceCmdText
         Case "IceConfig_InvStats_By_Code"
            .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 10, nCode)
            .Parameters.Append .CreateParameter("InvCode", adVarChar, adParamInput, 10, InvOrSvc)
         
         Case "IceConfig_InvStats_By_Practice"
            .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 10, InvOrSvc)
            
         Case "IceConfig_ResultStats_By_Code"
            .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 10, nCode)
            .Parameters.Append .CreateParameter("ResCode", adVarChar, adParamInput, 10, InvOrSvc)
         
         Case "IceConfig_ResultStats_By_Practice"
            .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 10, InvOrSvc)
         
         Case "IceConfig_MsgStats_By_Day"
'            .Parameters.Append .CreateParameter("SvcType", adInteger, adParamInput, , InvOrSvc)
            
         Case "IceConfig_MsgTypeStats_By_Practice"
         
         Case "IceConfig_MsgCountStats_By_Status"
            .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 10, InvOrSvc)
'            .Parameters.Append .CreateParameter("SvcType", adInteger, adParamInput, , InvOrSvc)
         
      End Select
      Set RunStoredProc = .Execute
   End With
   Set icecmd = Nothing
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   Else
      eClass.CurrentProcedure = "LoadMonitors.RunStoredProc"
      eClass.Add Err.Number, Err.Description, Err.Source
   End If
End Function

Public Sub RunWhat(NodeId As Node, Optional Refresh As String = "")
   On Error GoTo procEH
   Dim Index As Integer
   
   If blnExecuteComplete Then
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "Details"
   '         fView.FrameToShow = Fra_HELP
            Details NodeId
            
         Case "FirstView"
            FirstView
            
         Case "Activity"
            Activities NodeId
            blnShowBrowser = True
            
         Case "Statistics"
            blnShowBrowser = True
            fView.FrameToShow = Fra_TESTDETAILS
            Statistics NodeId
            
         Case "CustomStats"
            blnShowBrowser = True
            fView.FrameToShow = Fra_TESTDETAILS
            CustomStats NodeId
            
   '      Case Else
   '         fView.FrameToShow = fra_None
            
      End Select
   Else
      Beep
      With frmWait
         .Label1.Caption = "Query still in progress!!!"
         .Show
      End With
   End If
   
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadMonitors.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
'
'Public Sub ShowPractices()
'   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
'
'   strSQL = "SELECT EDI_NatCode " & _
'            "FROM EDI_Recipients"
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'   Do Until RS.EOF
'      frmMain.lviewDisplay.ListItems.Add , , RS!EDI_NatCode
'      RS.MoveNext
'   Loop
'   RS.Close
'   Set RS = Nothing
'   Exit Sub
'
'procEH:
'   If eClass.Behaviour = -1 Then
'      Stop
'      Resume
'   End If
'   eClass.CurrentProcedure = "LoadMonitors.ShowPractices"
'   eClass.Add Err.Number, Err.Description, Err.Source
'End Sub
'
'Public Sub ShowInvestigations()
'   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
'
'   strSQL = "SELECT DISTINCT Investigation_Requested " & _
'            "FROM Service_Investigations"
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'   Do Until RS.EOF
'      frmMain.lviewDisplay.ListItems.Add , , RS!Investigation_Requested
'      RS.MoveNext
'   Loop
'   RS.Close
'   Set RS = Nothing
'   Exit Sub
'
'procEH:
'   If eClass.Behaviour = -1 Then
'      Stop
'      Resume
'   End If
'   eClass.CurrentProcedure = "LoadMonitors.ShowInvestigations"
'   eClass.Add Err.Number, Err.Description, Err.Source
'End Sub

Public Sub CustomStats(Optional NodeId As Node = Nothing)
   On Error GoTo procEH
   Dim strOut As String
   Dim strCaption As String
   Dim statFile As String
   Dim i As Integer
   Dim icecmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim pVal As String
   
   frmWait.Label1 = "Please wait - gathering data..."
   frmWait.Visible = True
   frmMain.Refresh
   frmWait.Refresh
   frmMain.MousePointer = vbHourglass

   vData = objTV.ReadNodeData(NodeId)
   
   With icecmd
      .ActiveConnection = iceCon
      .Prepared = True
      .CommandType = adCmdStoredProc
      
      If DateDiff("d", frmMain.DTStats.value, Now()) = 0 Then
         strCaption = Format(frmMain.DTStats.value, "yyyymmdd hh:nn:ss")
      Else
         strCaption = Format(frmMain.DTStats.value, "yyyymmdd 23:59:59")
      End If
   
      .CommandText = "ICECONFIG_StatsX_" & vData(0)
      
      .Parameters.Append .CreateParameter("fromDate", adVarChar, adParamInput, 20, Format(frmMain.DTStatEnd.value, "yyyymmdd 00:00:00"))
      .Parameters.Append .CreateParameter("toDate", adVarChar, adParamInput, 20, strCaption)
      
      If vData(1) <> "0" Then
         Load frmCustomStats
         frmCustomStats.StoredProc = "ICECONFIG_StatsX_" & vData(0)
'         strSQL = "SELECT name, xType, Length " & _
'                  "FROM syscolumns " & _
'                  "WHERE id = " & vData(1)
'         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
'         Do Until RS.EOF
'            If UCase(RS!Name) <> "@FROMDATE" And UCase(RS!Name) <> "@TODATE" Then
'               pVal = InputBox("Value for parameter " & Mid(RS!Name, 2), "Extra parameter required")
'
'               Select Case RS!xType
'                  Case 48
'                     .Parameters.Append .CreateParameter(Mid(RS!Name, 2), adTinyInt, adParamInput, , pVal)
'
'                  Case 52
'                     .Parameters.Append .CreateParameter(Mid(RS!Name, 2), adSmallInt, adParamInput, , pVal)
'
'                  Case 56
'                     .Parameters.Append .CreateParameter(Mid(RS!Name, 2), adInteger, adParamInput, , pVal)
'
'                  Case 61
'                     .Parameters.Append .CreateParameter(Mid(RS!Name, 2), adDate, adParamInput, , pVal)
'
'                  Case 167
'                     .Parameters.Append .CreateParameter(Mid(RS!Name, 2), adVarChar, adParamInput, RS!Length, Mid(pVal, 1, RS!Length))
'
'               End Select
'            End If
'            RS.MoveNext
'         Loop
      End If
'      RS.Close
      With frmCustomStats
         For i = 0 To frmCustomStats.lstParams.ListCount - 1
               icecmd.Parameters.Append icecmd.CreateParameter(Mid(.lstParams.List(i), 2), adVarChar, adParamInput, .lstPValue.ItemData(i), .lstParams.List(i))
         Next i
      End With
      Set RS = .Execute
      
   End With
        
   'Open App.Path & "\" & vData(0) & ".csv" For Output As #1
   statFile = Read_Ini_Var("GENERAL", "StatsPath", iniFile)
   If statFile = "" Then
      statFile = App.Path
   End If
   
   statFile = fs.BuildPath(statFile, vData(0) & ".csv")
   Open statFile For Output As #1
   
   If RS.EOF = False Then
      For i = 0 To RS.Fields.Count - 1
         strOut = strOut & RS.Fields(i).Name & ","
      Next i
      strOut = Left(strOut, Len(strOut) - 1)
      Print #1, strOut
      strOut = ""
   End If
   
   Do Until RS.EOF
      For i = 0 To RS.Fields.Count - 1
         strOut = strOut & RS(i) & ","
      Next i
      strOut = Left(strOut, Len(strOut) - 1)
      Print #1, strOut
      strOut = ""
      RS.MoveNext
   Loop
   RS.Close
   
   Close #1
   
   MsgBox vData(2) & " procedure complete." & vbCrLf & "Saved in CSV File: " & statFile, vbInformation, "Custom Statistics"
   
   Set RS = Nothing
   Unload frmWait
   frmMain.MousePointer = vbNormal
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMonitors.CustomStats"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Statistics(Optional NodeId As Node = Nothing)
   On Error GoTo procEH
   Dim icecmd As New ADODB.Command
   Dim i As Integer
   Dim cWidth As Integer
   Dim strCaption As String
   Dim strSQL As String
   
   If blnExecuteComplete Then
      blnExecuteComplete = False
      
      With frmWait
         .Label1.Caption = "Please wait. Compiling statistics..."
         .ProgressBar1.Visible = False
         .tmrBusy.Enabled = True
         .fraBusy.Visible = True
         .Show
         .Refresh
      End With
         
      DoEvents
      
      If Not (NodeId Is Nothing) Then
         vData = objTV.ReadNodeData(NodeId)
         curNodeType = vData(0)
         curParam = vData(1)
         
         Set iceAsyncCmd = New ADODB.Command
         
         With iceAsyncCmd
            .ActiveConnection = iceAsyncCon
            .Prepared = True
            .CommandType = adCmdStoredProc
            
            If DateDiff("d", frmMain.DTStats.value, Now()) = 0 Then
               strCaption = Format(frmMain.DTStats.value, "yyyymmdd hh:nn:ss")
            Else
               strCaption = Format(frmMain.DTStats.value, "yyyymmdd 23:59:59")
            End If
            
            .Parameters.Append .CreateParameter("fromDate", adVarChar, adParamInput, 20, Format(frmMain.DTStatEnd.value, "yyyymmdd 00:00:00"))
            .Parameters.Append .CreateParameter("toDate", adVarChar, adParamInput, 20, strCaption)
            
            Select Case vData(0)
               Case "INVS"
                  If frmMain.cboStatsPractice.ListIndex = 0 Then
                     .CommandText = "ICECONFIG_Stats_InvAll"
                     strCaption = "All investigations"
                  Else
                     .Parameters.Append .CreateParameter("natcode", adVarChar, adParamInput, 8, frmMain.cboStatsPractice.Text)
                     .CommandText = "ICECONFIG_Stats_InvByPractice"
                     strCaption = "Investigations for Practice: " & frmMain.cboStatsPractice.Text
                  End If
                  
               Case "RESS"
                  If frmMain.cboStatsPractice.ListIndex = 0 Then
                     .CommandText = "ICECONFIG_Stats_ResultsAll"
                     strCaption = "All results"
                  Else
                     .Parameters.Append .CreateParameter("natcode", adVarChar, adParamInput, 8, frmMain.cboStatsPractice.Text)
                     .CommandText = "ICECONFIG_Stats_ResultByPractice"
                     strCaption = "Results for Practice: " & frmMain.cboStatsPractice.Text
                  End If
                  
               Case "TEST"
                  strCaption = "Reports incorporating "
                  With frmResultMapFilter
                     .sTabRC.TabEnabled(0) = False
                     .Show 1
                     If .optFind(0).value Then
                        icecmd.CommandText = "ICECONFIG_Stats_ReportsForLocalCode"
                        strCaption = strCaption & "Local code: "
                     ElseIf .optFind(1).value Then
                        icecmd.CommandText = "ICECONFIG_Stats_ReportsForLocalDesc"
                        strCaption = strCaption & "Local description: "
                     Else
                        icecmd.CommandText = "ICECONFIG_Stats_ReportsForReadCode"
                        strCaption = strCaption & "Read Code: "
                     End If
                  End With
                  
                  .Parameters.Append .CreateParameter("LocCode", adVarChar, adParamInput, 35, frmResultMapFilter.txtFind.Text)
                  strCaption = strCaption & frmResultMapFilter.txtFind.Text
                  Unload frmResultMapFilter
                  
               Case "MSGP"
                  .CommandText = "IceConfig_Stats_MsgByPractice"
                  .Parameters.Append .CreateParameter("flag", adInteger, adParamInput, , vData(1))
                  strCaption = NodeId.Text & " messages"
                               
               Case "MSGI"
                  .CommandText = "IceConfig_Stats_MsgIn"
                  strCaption = "Messages In"
               
               Case Else
                  .CommandType = adCmdUnknown
                  .CommandText = "None"
               
            End Select
            
            If .CommandType = adCmdUnknown Then
               If fxdRS.State = adStateOpen Then
                  fxdRS.Close
               End If
               Set fxdRS = Nothing
            Else
               If frmMain.optStatPeriod(0).value Then
                  strCaption = strCaption & " on " & Format(frmMain.DTStats.value, "d/mm/yyyy")
               Else
                  strCaption = strCaption & " between " & Format(frmMain.DTStatEnd.value, "d/mm/yyyy") & _
                               " and " & Format(frmMain.DTStats.value, "d/mm/yyyy")
               End If
               
               frmMain.wb.LocationTitle = strCaption
               frmMain.Refresh
      
               Set fxdRS = .Execute(, , adAsyncExecute)
               frmWait.Refresh
               frmWait.MousePointer = vbHourglass
               frmMain.MousePointer = vbHourglass
               DoEvents
               'PrepareHTML CStr(vData(0))
               'fxdRS.Close
            End If
         End With
      End If
   Else
      Beep
      With frmWait
         .Label1.Caption = "Query still in progress!!!"
         .Show
      End With
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMonitors.Statistics"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
'
'Public Sub Timer()
'   Dim i As Integer
'   Dim FL As Integer
'   FL = 0
'   For i = 0 To UBound(ActivityDirs)
'      If (ActivityDirs(i) <> "") Then
'         frmMain.FLActivity(FL).Refresh
'         FL = FL + 1
'      End If
'   Next i
'End Sub

Public Property Let ViewingOption(intNewValue As Integer)
   CurrentViewOption = intNewValue
End Property

Private Sub Class_Terminate()
   Unload frmChart
   Unload frmWait
   With frmMain.wb
      .Top = 360
      .Left = 240
      .NavigateTo App.Path & "/logo.html"
   End With
   
   fView.FrameToShow = fra_None
   
'   invRS.Close
'   resRS.Close
   pracRS.Close
   iceAsyncCon.Close
   Set invRS = Nothing
   Set resRS = Nothing
   Set pracRS = Nothing
   Set iceAsyncCon = Nothing
End Sub

Private Sub iceAsyncCon_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   PrepareHTML curNodeType
   fxdRS.Close
   Set fxdRS = Nothing
   Set iceAsyncCmd = Nothing
   blnExecuteComplete = True
   frmWait.MousePointer = vbNormal
   Unload frmWait
   frmMain.MousePointer = vbNormal
End Sub

Private Sub iceAsyncCon_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   Debug.Print "WillExecute: " & adStatus
End Sub
