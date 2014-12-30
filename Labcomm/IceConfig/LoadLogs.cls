VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private tv1 As TreeView
Private tempTv As TreeView
Private mCtrl As New ManageControls
Private vData As Variant
Private strArray() As String
Private tvNode(4) As Node
Private practDest As String
Private defDateForm As String
Private myDateForm As String
Private Requeuestr As String
Private maxLen As Integer
Private StartDate As String
Private EndDate As String
Private PracticeId As String
Private logLTSIndex As Integer
Private LogFileName As String
Private blnNoFile As Boolean
Private objApp As outlook.Application
Private blnSearching As Boolean
Private blnHL7db As Boolean
Private hl7UDL As String
Private HL7Con As ADODB.Connection
Private WithEvents iceAsyncCon As ADODB.Connection
Attribute iceAsyncCon.VB_VarHelpID = -1
Private curNode As Node
Private execProc As String
Private asyncRS As ADODB.Recordset
Attribute asyncRS.VB_VarHelpID = -1
Private maxFilesPerNode As Integer
Private inProgress As Boolean
Private rtc As Integer

Private fileDate As String
Private filterNatCode As String
Private rsFile As String


Private Sub Class_Initialize()
   Dim strSQL As String
   Dim UDLFile As String
   Dim RS As New ADODB.Recordset
   Dim iniData As String
   
   iniData = Read_Ini_Var("General", "dbTimeOut", iniFile)
   
   If iniData = "" Or IsNumeric(iniData) = False Then
      iniData = "10"
   End If
   
   Set iceAsyncCon = New ADODB.Connection
   iceAsyncCon.CommandTimeout = Val(iniData)
   iceAsyncCon.ConnectionTimeout = Val(iniData)
   iceAsyncCon.ConnectionString = iceCon.ConnectionString
   iceAsyncCon.Open
   
'  Read the default date fromat for the database language, store, then set dateformat to dmy
'  This ensures the date restriction constructs work as intended
   
   strSQL = "SELECT dateformat " & _
            "FROM master.dbo.syslanguages " & _
            "WHERE langid = @@default_langid"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   defDateForm = ";SET dateformat " & RS!DateFormat
   RS.Close
   
   hl7UDL = Read_Ini_Var("GENERAL", "UseHL7DB", iniFile)
   If hl7UDL <> "" Then
   
   
   'If blnHL7db Then
      Set HL7Con = New ADODB.Connection
      UDLFile = Read_Ini_Var("GENERAL", "ConfigPath", iniFile)
      hl7UDL = fs.BuildPath(UDLFile, hl7UDL)
      
      If fs.FileExists(hl7UDL) Then
         blnHL7db = True
         With HL7Con
            .CursorLocation = adUseClient
            .Mode = adModeReadWrite
            .Open "FILE NAME=" & hl7UDL
         End With
      Else
         blnHL7db = False
         MsgBox hl7UDL & " specified in ini file does not exist - HL7 source unavailable"
      End If
   End If
   
   
   rsFile = Dir(fs.BuildPath(App.Path, "*.RS"))
   Do Until rsFile = ""
      rsFile = fs.BuildPath(App.Path, rsFile)
      fs.DeleteFile rsFile, True
      
      rsFile = Dir
   Loop
   
   iniData = Read_Ini_Var("General", "LogsMaxPerNode", iniFile)
   If IsNumeric(iniData) Then
      If Val(iniData) > 1000 Then
         maxFilesPerNode = 1000
      Else
         maxFilesPerNode = iniData
      End If
   Else
      maxFilesPerNode = 100
   End If
   
   myDateForm = "SET dateformat dmy;"
   Set RS = Nothing
   Set tv1 = frmMain.TreeView1
   
   fView.FrameToShow = Fra_LOGVIEW

End Sub

Public Sub Collapse(NodeId As Node)
   Dim i As Integer
   
'   If NodeId.Children > 2 Then
'      For i = 1 To NodeId.Children
'         tv1.Nodes.Remove NodeId.child.Index
'      Next i
'
'      tv1.Nodes.Add NodeId, _
'                    tvwChild, _
'                    mCtrl.NewNodeKey("Tmp", _
'                                     "Tmp"), _
'                                     "Please wait...", _
'                                     13, _
'                                     13
'   End If
End Sub

Public Property Get CurrentLogFile() As String
   CurrentLogFile = LogFileName
End Property
'

Private Sub SearchDisplayReport(NodeId As Node)
   vData = objTV.ReadNodeData(NodeId.child)
   logLTSIndex = vData(0)
   DisplayReport NodeId
End Sub

Public Property Let SearchStartDate(strNewValue As Date)
   StartDate = Format(strNewValue, "yyyymmdd") & " 00:00:00"
End Property

Public Property Let SearchEndDate(strNewValue As Date)
   EndDate = Format(strNewValue, "yyyymmdd") & " 23:59:59"
End Property

Public Property Let SearchInProgress(blnNewValue As Boolean)
   blnSearching = blnNewValue
End Property

Public Property Let Practice(strNewValue As String)
   PracticeId = strNewValue
End Property

Public Sub Dates(NodeId As Node) ', _
                 Optional startDate As Date = 0, _
                 Optional endDate As Date = 0, _
                 Optional practiceID As String = "")
   On Local Error GoTo procEH
   Dim i As Integer
   Dim RS As New ADODB.Recordset
   Dim thisPractice As String
   Dim iconId As Integer
   Dim callType As String
   Dim data1 As String
   Dim data2 As String
   Dim eDate As Date
   Dim sDate As Date
   Dim thisDate As Date
   Dim thisYear As String
   Dim thisMth As String
   Dim thisDay As String
   Dim blnNewSql As Boolean
   Dim tSql As String
   Dim closeSQL As String
   Dim iLoop As Integer
   Dim oLoop As Integer
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode)
   tv1.Visible = False
   
   For i = 0 To frmMain.OrgList.ListCount - 1
      If frmMain.OrgList.ItemData(i) = vData(0) Then
         frmMain.OrgList.ListIndex = i
         Exit For
      End If
   Next i
   
   vData = objTV.ReadNodeData(NodeId.Parent)
   logLTSIndex = vData(0)
   
   vData = objTV.ReadNodeData(NodeId)
   If vData(0) = 1 Then
      filterNatCode = ""
   End If
   
   'If NodeId.Tag <> "" Then
      rsFile = NodeId.Tag
   'End If
   
   If fs.FileExists(rsFile) Then
      'asyncRS.Open rsFile, , , , adCmdFile
      DatesForServiceType NodeId
   Else
      rsFile = "dates_" & logLTSIndex & "_" & vData(0)
      rsFile = fs.BuildPath(App.Path, rsFile & ".RS")
      NodeId.Tag = rsFile
   
      If PracticeId <> "" Then
         strSQL = "SELECT EDI_Org_NatCode " & _
                  "FROM EDI_Recipient_Individuals er " & _
                     "INNER JOIN EDI_Matching em " & _
                     "ON er.Individual_Index = em.Individual_Index " & _
                  "WHERE EDI_Local_Key1 = '" & Replace(PracticeId, "'", "''") & "'"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         filterNatCode = RS!EDI_Org_NatCode
         RS.Close
      Else
         filterNatCode = ""
      End If
      
      With frmWait
         .Label1.Caption = "Please wait - reading Logs..."
         .ProgressBar1.value = 0
         .ProgressBar1.max = 1
         .ProgressBar1.Visible = False
         .tmrBusy.Enabled = True
         .fraBusy.Visible = True
         .Show
         .Refresh
      End With
      DoEvents
         
      oLoop = 0
      iLoop = 0
      
      eClass.FurtherInfo = "Reading Service_ImpExp_HeadeRS"
   
      strSQL = "SELECT DISTINCT Max(Convert(int,Warning_Flag)) As Warnings, " & _
                  "Max(Error_Found) As ErroRS, " & _
                  "Max(Status_Flag) As Status_ErroRS, " & _
                  "Convert(varchar, sh.Date_Added,112) as Rep_Date_Added, " & _
                  "Convert(varchar, sh.Date_Added,103) as Date_Added " & _
               "FROM Service_ImpExp_HeadeRS sh " & _
                  "LEFT JOIN EDI_Recipients "
      If phoenix Then
         strSQL = strSQL & "  ON PatIndex('%' + EDI_NatCode + '%',ImpExp_File) > 0 "
      Else
         strSQL = strSQL & " er" & vbCrLf & _
               "    INNER JOIN EDI_Recipient_Ref err " & vbCrLf & _
               "    ON er.Ref_Index=err.Ref_Index " & vbCrLf & _
               "  ON Trader_Code = EDI_Trader_Account + EDI_Free_Part " & vbCrLf
      End If
   
      If vData(0) = 1 Then
         strSQL = strSQL & " OR Trader_Code Is Null "
      End If
      
      strSQL = strSQL & _
               "LEFT JOIN Service_ImpExp_Messages sm " & _
               "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id "
      
      If logLTSIndex = 0 Then
         strSQL = strSQL & " WHERE Service_Type = 2"
   
      Else
         strSQL = strSQL & " WHERE Service_Type = " & vData(0)
      End If
            
      strSQL = strSQL & " AND (Case " & _
                           "When sm.EDI_LTS_Index is null Then sh.EDI_LTS_Index " & _
                           "When sh.EDI_LTS_Index=0 then 0 " & _
                           "When sm.EDI_LTS_Index = 0 then sh.EDI_LTS_Index " & _
                           "When sm.EDI_LTS_Index <> sh.EDI_LTS_Index then sm.EDI_LTS_Index " & _
                           "Else sm.EDI_LTS_Index " & _
                           "End) = " & logLTSIndex
      
      If filterNatCode <> "" Then
         strSQL = strSQL & " AND EDI_NatCode = '" & filterNatCode & "' "
      End If
   
      If blnSearching Then
         iceAsyncCon.CommandTimeout = 30
         thisDate = frmMain.dtPFrom.value
         eDate = frmMain.dtPTo.value
      Else
         iceAsyncCon.CommandTimeout = 10
         thisDate = DateAdd("d", -logHist, Now())
         eDate = Now()
      End If
   
      thisYear = Year(thisDate)
      thisMth = Month(thisDate)
      thisDay = Day(thisDate)
   
      tSql = " AND (datepart(yyyy,sh.Date_Added) = '" & thisYear & "'" & _
                  "AND (datepart(m,sh.Date_Added) = '" & thisMth & "'" & _
                     "AND (datepart(d,sh.Date_Added) >= '" & thisDay & "'"
      closeSQL = "))"
      
      Do
         If thisYear <> Year(thisDate) Then
            thisYear = Year(thisDate)
            blnNewSql = True
         End If
   
         If thisMth <> Month(thisDate) Then
            tSql = tSql & "AND datepart(d,sh.Date_Added) <= '" & thisDay & "'" & closeSQL
   
            thisMth = Month(thisDate)
            thisDay = Day(thisDate)
            blnNewSql = True
         End If
         
         If blnNewSql Then
            tSql = tSql & _
                   "OR (datepart(yyyy,sh.Date_Added) = '" & thisYear & "'" & _
                     "AND (datepart(m,sh.Date_Added) = '" & thisMth & "'" & _
                        "AND (datepart(d,sh.Date_Added) >= '" & thisDay & "'"
            closeSQL = ")))"
            
            blnNewSql = False
         End If
         
         thisDay = Day(thisDate)
         thisDate = DateAdd("d", 1, thisDate)
         
      Loop Until thisDate > eDate
      
      tSql = tSql & "AND datepart(d,sh.Date_Added) <= '" & thisDay & "')" & closeSQL 'closeSQL '  "'))))"
            
      strSQL = strSQL & tSql
      
      
      strSQL = strSQL & _
               " GROUP BY Convert(varchar, sh.Date_Added,112), Convert(varchar, sh.Date_Added,103) " & _
               " ORDER BY Rep_Date_Added DESC"
      
      execProc = "Dates"
      Set curNode = NodeId
   
      Set asyncRS = New ADODB.Recordset
      asyncRS.Open strSQL, iceAsyncCon, adOpenKeyset, adLockReadOnly, adAsyncExecute
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
   Unload frmWait
   eClass.CurrentProcedure = "LoadLogs.Dates"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub DatesForServiceType(NodeId As Node)
   Dim RS As New ADODB.Recordset
   Dim iconId As Integer
   Dim iLoop As Integer
   Dim i As Integer
   Dim TempStr As String
   
   RS.Open rsFile, , , , adCmdFile
   'Set RS = asyncRS
   With frmWait
      .Label1.Caption = "Building tree View..."
      .ProgressBar1.max = frmWait.ProgressBar1.max + (RS.RecordCount)
      .tmrBusy.Enabled = False
      .fraBusy.Visible = False
      .ProgressBar1.Visible = True
      .Show
      
   End With
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   If filterNatCode <> "" Then
      If InStr(1, NodeId.Text, "[") = 0 Then
         NodeId.Text = NodeId.Text & " [For practice " & filterNatCode & "]"
      End If
   End If
   
   If RS.RecordCount > 0 Then
      frmWait.Label1.Caption = "Building tree view..."
      frmWait.Refresh
      DoEvents
      
      Do While Not RS.EOF
         iLoop = iLoop + 1
         iconId = 13
         If RS!warnings > 0 Then
            iconId = 25
         End If
         
         If RS!Errors > 0 Then
            iconId = 26
         End If
         
         If RS!Status_Errors > 0 Then
            iconId = 26
         End If
                  
         Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                       tvwChild, _
                                       mCtrl.NewNodeKey(RS!Date_Added, _
                                                        filterNatCode, _
                                                        "Files"), _
                                       Trim(Format(RS!Date_Added, "dd mmm yy")), _
                                       iconId, _
                                       iconId)
         
         Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
                                       tvwChild, _
                                       mCtrl.NewNodeKey("Tmp", _
                                                        "Tmp"), _
                                                        "Please wait...", _
                                                        13, _
                                                        13)
         frmWait.ProgressBar1.value = iLoop
         frmWait.Refresh
         DoEvents
         
         RS.MoveNext
      Loop
      
   Else
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", _
                                     "Tmp"), _
                    "<No data found>", _
                    13, _
                    13
   End If
   
   frmWait.Hide
   Unload frmWait
   tv1.Visible = True
   frmMain.MousePointer = vbNormal
   RS.Close
   Set RS = Nothing
   
   inProgress = False

End Sub

Public Function FileLocation(fPath As String, _
                             nodeSource As String, _
                             errorFile As Boolean) As String
   Dim RS As New ADODB.Recordset
   
   Dim iVar As String
   Dim newPath As String
   Dim cName As String
   Dim eDir As String
   Dim driveId As String
   
   cName = ""
      driveId = fs.GetDriveName(fPath)
      If fs.DriveExists(driveId) = False Then
         If Len(driveId) > 3 Then
            fPath = "C:" & Mid(fPath, InStr(3, fPath, "\"))
         End If
      End If
   
   If fs.FolderExists(fPath) Then
      newPath = fPath
   Else
      strSQL = "SELECT * " & _
               "FROM Service_Types " & _
               "WHERE Type_Index = " & nodeSource
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.BOF = False And RS.EOF = False Then
         strSQL = "SELECT Connection_Name " & _
                  "FROM Connect_Modules "
         Select Case UCase(Trim(RS!Description))
            Case "LABORATORY REPORTS"
               strSQL = strSQL & _
                        "WHERE Module_Name = 'ICEIMP.EXE')"
               
            Case "EDI REPORTS"
               strSQL = strSQL & _
                        "WHERE Module_Name = 'ICEMSG.EXE')"
                        
            Case "FHSA FILES"
               strSQL = strSQL & _
                        "WHERE Module_Name = 'ICEMFHSAEXE')"
               
         End Select
         
         If errorFile Then
            iVar = Left(cName, 3) & "Err"
         Else
            iVar = Left(cName, 3) & "Hist"
         End If
      End If
      RS.Close
      
      strSQL = "SELECT * " & _
               "FROM Connections " & _
               "WHERE Connection_Name = (" & strSQL
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.EOF = False Then
        If errorFile Then
           eDir = Mid(fPath, Len(Trim(RS!Connection_ErrorDirs & "")) + 2)
        Else
           eDir = Mid(fPath, Len(Trim(RS!Connection_HistoryDirs & "")) + 2)
        End If
      End If
      
      RS.Close
      
      newPath = Read_Ini_Var("Directory_Overrides", iVar, iniFile)
      If newPath = "" Then
         newPath = fPath
      Else
         newPath = fs.BuildPath(newPath, eDir)
      End If
   End If
   
   FileLocation = newPath
   
   Set RS = Nothing
End Function

Public Sub FilesForDay(NodeId As Node)
   Dim TempStr As String
   Dim RS As New ADODB.Recordset
   Dim fDir As String
   Dim lType As String
   Dim tRecs As Integer
   Dim iconId As Integer
   Dim iLoop As Integer
   Dim i As Integer
   
   rsFile = NodeId.Tag
   
   If logLevel = 2 Then
      eClass.LogMessage "SQL = " & strSQL & vbCrLf & " - Records returned = " & RS.RecordCount & " EOF = " & RS.EOF
   End If
   
   RS.Open rsFile, , , , adCmdFile
   
   tv1.Visible = False
   
   tRecs = RS.RecordCount
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   If tRecs > 0 Then
      
      With frmWait
         .fraBusy.Visible = False
         .ProgressBar1.Visible = True
         .ProgressBar1.max = frmWait.ProgressBar1.max + tRecs
         .Label1.Caption = "Building tree view..."
         .Show
         .Refresh
      End With
      
      DoEvents
      
      If InStr(1, NodeId.Text, "[") = 0 Then
         NodeId.Text = NodeId.Text & " [" & RS.RecordCount & " files]"
      End If
      
      Do While Not RS.EOF
         iLoop = iLoop + 1
         
         If RS!Warning_Flag = True Then
            iconId = 25
         Else
            iconId = 13
         End If
         
         If Left(RS!ImpExp_File, 1) = "+" Then
            lType = "ErrReport"
            TempStr = Mid(RS!ImpExp_File, 2)
            fDir = RS!Date_Added
         Else
            TempStr = Trim(fs.GetFileName(Trim(RS!ImpExp_File)))
            lType = "Reports"
            fDir = fs.GetParentFolderName(Trim(RS!ImpExp_File))
         End If
         
         Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                       tvwChild, _
                                       mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                                        fDir, _
                                                        lType, _
                                                        RS!ImpExp_File), _
                                       TempStr, _
                                       iconId, _
                                       iconId)
         
         If RS!Error_Found > 0 Or _
            RS!Status_Flag > 0 Then
            tvNode(0).ForeColor = BPRED
            tvNode(0).Image = 26
            tvNode(0).SelectedImage = 26
         ElseIf RS!Status_Flag < 0 Then
            tvNode(0).ForeColor = BPBLUE
         Else
            tvNode(0).ForeColor = BPGREEN
         End If
         
         tv1.Nodes.Add tvNode(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey("Tmp", _
                                        "Tmp"), _
                       "Please wait...", _
                       13, _
                       13
         
         frmWait.Label1.Caption = "Building tree view...(" & iLoop & " records)"
         frmWait.ProgressBar1.value = iLoop
         frmWait.Refresh
         DoEvents
         
         RS.MoveNext
      Loop
   Else
      tv1.Nodes.Add tvNode(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", _
                                     "Tmp"), _
                    "<No Files>", _
                    13, _
                    13
   End If
   
   tv1.Visible = True
   Unload frmWait
   
   inProgress = False
   RS.Close
   Set RS = Nothing
End Sub

Public Sub Files(NodeId As Node)
   On Local Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim tvNode(4) As Node
   Dim iconId As Integer
   Dim TempStr As String
   Dim fDir As String
   Dim lType As String
   Dim i As Integer
   Dim tRecs As Integer
   Dim iLoop As Integer
   Dim mbResp As Integer
   Dim blnShowReps As Boolean
   Dim nc As Integer
   
   lType = objTV.NodeLevel(objTV.TopLevelNode(objTV.ActiveNode, True))
      
   vData = objTV.ReadNodeData(NodeId.Parent.Parent)
   
   logLTSIndex = vData(0)
      
   tv1.Visible = False
   frmMain.Refresh
   DoEvents
   
   If lType = "Patient" Then
      lType = objTV.NodeLevel(objTV.ActiveNode.Parent)
   End If
   
   vData = objTV.ReadNodeData(NodeId)
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
    
   rsFile = NodeId.Tag
   
   If fs.FileExists(rsFile) Then
      RS.Open rsFile, , , , adCmdFile
      tRecs = RS.RecordCount
      RS.Close
      
      If tRecs > maxFilesPerNode Then
         Times NodeId
      Else
         FilesForDay NodeId
      End If
      
   Else
      rsFile = "files_" & logLTSIndex & "_" & lType & "_" & Format(CDate(vData(0)), "yyyymmdd")
      rsFile = fs.BuildPath(App.Path, rsFile & ".RS")
      NodeId.Tag = rsFile
   
      iLoop = 0
      
      With frmWait
         .Label1.Caption = "Please wait - reading Logs..."
         .ProgressBar1.value = 0
         .ProgressBar1.max = 1
         .ProgressBar1.Visible = False
         .tmrBusy.Enabled = True
         .fraBusy.Visible = True
         .Show
         .Refresh
      End With
      
      DoEvents
      
      If lType = "1" Then
         strSQL = "SELECT DISTINCT sh.Date_Added, sh.Service_ImpExp_Id, ImpExp_File, Status_Flag, Warning_Flag, Error_Found " & _
                  "FROM Service_ImpExp_HeadeRS sh " & _
                     "LEFT JOIN Service_ImpExp_Messages sm " & _
                     "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
                  "WHERE Service_Type = " & lType
      
      Else
         strSQL = "SELECT DISTINCT sh.Date_Added, sh.Service_ImpExp_Id, ImpExp_File, Status_Flag, Warning_Flag, Error_Found " & _
                  "FROM Service_ImpExp_HeadeRS sh " & _
                     "LEFT JOIN Service_ImpExp_Messages sm " & _
                     "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
                     "LEFT JOIN EDI_Recipients "
               
         If phoenix Then
            strSQL = strSQL & "  ON PatIndex('%' + EDI_NatCode + '%',ImpExp_File) > 0 "
         Else
            strSQL = strSQL & " er" & vbCrLf & _
                  "    INNER JOIN EDI_Recipient_Ref err " & vbCrLf & _
                  "    ON er.Ref_Index=err.Ref_Index " & vbCrLf & _
                  "  ON Trader_Code = EDI_Trader_Account + EDI_Free_Part " & vbCrLf
            'strSQL = strSQL & "ON Trader_Code = EDI_Trader_Account + EDI_Free_Part "
         End If
         
         If logLTSIndex = 0 Then
            strSQL = strSQL & "WHERE Service_Type = 2"
         Else
            strSQL = strSQL & "WHERE Service_Type = " & lType
         End If
      End If
      
      strSQL = strSQL & " AND (Case " & _
                           "When sm.EDI_LTS_Index is null Then sh.EDI_LTS_Index " & _
                           "When sh.EDI_LTS_Index=0 then 0 " & _
                           "When sm.EDI_LTS_Index = 0 then sh.EDI_LTS_Index " & _
                           "When sm.EDI_LTS_Index <> sh.EDI_LTS_Index then sm.EDI_LTS_Index " & _
                           "Else sm.EDI_LTS_Index " & _
                        "End) = " & logLTSIndex & _
                        "  AND (datepart(yyyy, sh.Date_Added)='" & Year(vData(0)) & "'" & _
                        "  AND datepart(m, sh.Date_Added) = '" & Month(vData(0)) & "'" & _
                        "  AND datepart(d, sh.Date_Added) = '" & Day(vData(0)) & "')"
                        '" AND convert(varchar,sh.Date_Added,103) = '" & Format(vData(0), "dd/mm/yyyy") & "' "
      
      
      
      If vData(1) <> "" Then
         If lType = "2" Then
            strSQL = strSQL & _
                     " AND EDI_NatCode = '" & vData(1) & "' "
         End If
      End If
      
      Select Case frmMain.LogDisplayType
         Case 1   '  ErroRS
            strSQL = strSQL & _
                     " AND Status_Flag > 0"
         
         Case 2   '  Warnings
            strSQL = strSQL & _
                     " AND Status_Flag < 0"
         
         Case 3   '  OK
            strSQL = strSQL & _
                     " AND Status_Flag = 0"
         
      End Select
      
      Select Case UCase(logSortOn)
         Case "ERROR"
            strSQL = strSQL & " ORDER BY Error_Found"
         
         Case "NAME"
            strSQL = strSQL & " ORDER BY ImpExp_File"
         
         Case Else
            strSQL = strSQL & " ORDER BY sh.Service_ImpExp_Id"
         
      End Select
      
      If logSortDesc Then
         strSQL = strSQL & " DESC"
      End If
      
      '***************************************************************************************
      '  Set execProc to indicate which procedure has set up the request. fileDate is set
      '  here to be picked up if the number of files exceeds ini specified amount.
      '
      '  These are utilised in the 'Execute_Complete' event handler for iceAsyncCon.
      '***************************************************************************************
      execProc = "Files"
            
      fileDate = Format(vData(0), "yyyymmdd")
      Set curNode = NodeId
      
      If asyncRS.State <> adStateClosed Then
         asyncRS.Close
      End If
      
      asyncRS.Open strSQL, iceAsyncCon, adOpenStatic, adLockReadOnly, adAsyncExecute
   
'      frmWait.Hide
'      Unload frmWait
      
   End If
   
   tv1.Visible = True
   frmMain.MousePointer = vbNormal
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
   Unload frmWait
   eClass.CurrentProcedure = "LoadLogs.Files"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub FilesByTime(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim fileTime As Date
   Dim i As Integer
   Dim iLoop As Integer
   Dim tvNode(4) As Node
   Dim iconId As Integer
   Dim fDir As String
   Dim lType As String
   Dim TempStr As String
   Dim blnDisplay As Boolean
   Dim sBatch() As String
   Dim eBatch() As String
   
   rsFile = NodeId.Parent.Tag
   execProc = ""

   vData = objTV.ReadNodeData(NodeId)
   sBatch = Split(vData(0), "#")
   eBatch = Split(vData(1), "#")
   
   'ADTGFile = fs.BuildPath(App.Path, vData(4) & "_log.RS")
   
   'RS.Open ADTGFile, , , , adCmdFile
   
   RS.Open rsFile, , , , adCmdFile
   
   tv1.Visible = False
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   If InStr(1, NodeId.Parent.Text, "[") = 0 Then
      NodeId.Parent.Text = NodeId.Parent.Text & " [" & RS.RecordCount & " files]"
   End If
   
   RS.Filter = "Service_ImpExp_Id >= " & vData(0) & " And Service_ImpExp_Id <= " & vData(1)
   RS.MoveFirst
   
   iLoop = 0
   frmWait.ProgressBar1.max = maxFilesPerNode
   Do
            
      If RS!Warning_Flag = True Then
         iconId = 25
      Else
         iconId = 13
      End If
      
      If Left(RS!ImpExp_File, 1) = "+" Then
         TempStr = Mid(RS!ImpExp_File, 2)
         lType = "ErrReport"
         fDir = RS!Date_Added
      Else
         TempStr = Trim(fs.GetFileName(Trim(RS!ImpExp_File)))
         lType = "Reports"
         fDir = fs.GetParentFolderName(Trim(RS!ImpExp_File))
      End If
      
      Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                    tvwChild, _
                                    mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                                     fDir, _
                                                     lType, _
                                                     RS!ImpExp_File), _
                                    TempStr, _
                                    iconId, _
                                    iconId)
      
      If RS!Error_Found > 0 Or _
         RS!Status_Flag > 0 Then
         tvNode(0).ForeColor = BPRED
         tvNode(0).Image = 26
         tvNode(0).SelectedImage = 26
      ElseIf RS!Status_Flag < 0 Then
         tvNode(0).ForeColor = BPBLUE
      Else
         tvNode(0).ForeColor = BPGREEN
      End If
      
      tv1.Nodes.Add tvNode(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", _
                                     "Tmp"), _
                    "Please wait...", _
                    13, _
                    13
      frmWait.Label1.Caption = "Building tree view...(" & iLoop & " records)"
      frmWait.ProgressBar1.value = iLoop
      frmWait.Refresh
      DoEvents
      iLoop = iLoop + 1
'      Else
'         If iLoop > 1 Then
'            Exit Do
'         End If
'      End If
      
      RS.MoveNext
   Loop Until RS.EOF
   
   tv1.Visible = True
   RS.Close
   Set RS = Nothing
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
   Unload frmWait
   eClass.CurrentProcedure = "LoadLogs.FilesByTime"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub FirstView()
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim svcDesc As String
   
   strSQL = "SELECT EDI_OrgCode, EDI_Msg_Type, EDI_LTS_Index " & _
            "FROM EDI_Local_Trader_Settings"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   Set tvNode(0) = tv1.Nodes.Add(, _
                                 , _
                                 mCtrl.NewNodeKey("0", _
                                                  "Error", _
                                                  "None"), _
                                 "Message Errors", _
                                 13, _
                                 13)
   Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
                                 tvwChild, _
                                 mCtrl.NewNodeKey("2", _
                                                  "IceMsg", _
                                                  "Dates"), _
                                 "IceMsg ErroRS" & "", _
                                 13, _
                                 13)
   
   Set tvNode(2) = tv1.Nodes.Add(tvNode(1), _
                                 tvwChild, _
                                 mCtrl.NewNodeKey("tmp", _
                                                  "Tmp"), _
                                 "Please wait...", _
                                 13, _
                                 13)

   Do Until RS.EOF
      Set tvNode(0) = tv1.Nodes.Add(, _
                                    , _
                                    mCtrl.NewNodeKey(RS!EDI_LTS_Index, _
                                                     RS!EDI_Msg_Type, _
                                                     "None"), _
                                    RS!EDI_OrgCode & " - " & RS!EDI_Msg_Type, _
                                    13, _
                                    13)
   
      eClass.FurtherInfo = "Reading Service_Types"
      strSQL = "SELECT Type_Index, Description " & _
               "FROM Service_Types " & _
               "WHERE Type_Index IN ( " & _
                  "SELECT DISTINCT Service_Type FROM Service_ImpExp_Headers) " & _
               "ORDER BY Type_Index"
      'stRSQL = "SELECT DISTINCT Service_Type, Description " & _
               "FROM Service_ImpExp_HeadeRS " & _
                  "LEFT JOIN Service_Types " & _
                  "ON Service_Type = Type_Index " & _
               "WHERE Organisation = '" & frmMain.cboTrust.Text & "' " & _
                  "AND EDI_LTS_Index = " & RS!EDI_LTS_Index & _
               " ORDER BY Service_Type"
      RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
      Do While Not RS2.EOF
         Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
                                       tvwChild, _
                                       mCtrl.NewNodeKey(RS2!Type_Index, _
                                                        svcDesc, _
                                                        "Dates"), _
                                       RS2!Description & "", _
                                       13, _
                                       13)
            
         Set tvNode(2) = tv1.Nodes.Add(tvNode(1), _
                                       tvwChild, _
                                       mCtrl.NewNodeKey("tmp", _
                                                        "Tmp"), _
                                       "Please wait...", _
                                       13, _
                                       13)
         RS2.MoveNext
      Loop
      RS2.Close
      RS.MoveNext
   Loop
   
   RS.Close
   Set RS2 = Nothing
   Set RS = Nothing
End Sub

Public Function Refresh() As String
   objTV.RefreshNode = objTV.ActiveNode
End Function

Public Sub Reports(NodeId As Node)
   On Error GoTo procEH
   Dim i As Integer
   Dim RS As New ADODB.Recordset
   Dim tvNode(4) As Node
   Dim OrgID As String
   Dim imgId As Integer
         
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   vData = objTV.ReadNodeData(NodeId)
   'fView.OutputFileName = vData(4)
   OrgID = frmMain.cboTrust.Text
'   strArray = Split(SvcImpExpId, ":")
   
   strSQL = "SELECT sm.*, Service_Type " & _
            "FROM Service_ImpExp_Headers sh " & _
               "LEFT JOIN Service_ImpExp_Messages sm " & _
               "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
            "WHERE sh.Service_ImpExp_Id = " & vData(0)
            
   If practDest <> "" Then
      strSQL = strSQL & _
               " AND Destination <> '' " & _
               "AND substring(Destination,7,len(Destination)-6) IN " & _
                  "(SELECT left(EDI_Local_Key1,10) " & _
                  "FROM EDI_Recipient_Individuals " & _
                  "WHERE EDI_Org_NatCode = '" & practDest & "')"
   
   End If
   
   strSQL = strSQL & " ORDER BY Service_Message_Id"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly

   If logLevel = 2 Then
      eClass.LogMessage "SQL = " & strSQL & vbCrLf & "Records returned: " & RS.RecordCount & " EOF = " & RS.EOF
   End If
      
   If IsNull(RS!Service_ImpExp_Message_ID) Then
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(CStr(vData(0)), _
                                     RS!Service_Type, _
                                     "Tracking"), _
                                     "Tracking", _
                                     13, _
                                     13
      
   Else
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                     RS!Service_Type, _
                                     "Tracking"), _
                                     "Tracking", _
                                     13, _
                                     13
      Do While Not RS.EOF
         If RS!EDI_LTS_Index = -1 Then
            imgId = 28
         Else
            imgId = 13
         End If
         Set tvNode(4) = tv1.Nodes.Add(NodeId, _
                                       tvwChild, _
                                       mCtrl.NewNodeKey(Trim(RS!Service_Report_Index), _
                                                        CStr(RS!Service_ImpExp_Message_ID), _
                                                        "RepDetails"), _
                                       Trim(Mid(RS!Service_Id, 1, 16)) & " (" & Trim(RS!Patient_Name) & " " & Trim(Mid(RS!Patient_Local_Id, 7, 30)) & ")", _
                                       imgId, _
                                       imgId)
         If RS!Rejected_Count > 0 Then
            tvNode(4).ForeColor = BPRED
         ElseIf RS!Rejected_Count < 0 Then
            tvNode(4).ForeColor = BPBLUE
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close
   Set RS = Nothing
   DisplayFile NodeId
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.Reports"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub DisplayFile(NodeId As Node)
   On Error GoTo procEH
   Dim sType As String
   Dim Tstr As String
   Dim fData As String
   Dim tBuf As New StringBuffer
   Dim RS As New ADODB.Recordset
   
   blnNoFile = False
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode(, True))
   
   If vData(2) = "PatientReport" Then
      sType = -1
   ElseIf vData(2) = "FileDets" Then
      sType = -1
   ElseIf vData(2) = "RepDetails" Then
      sType = objTV.NodeLevel(objTV.ActiveNode.Parent)
   ElseIf vData(2) = "NotPresent" Then
      sType = -1
   ElseIf vData(2) = "PatientFiles" Then
      sType = -1
   Else
      sType = vData(0)
   End If
   
   frmMain.MailBtn.Caption = "RESEND OR REQUEUE"
   frmMain.MailBtn.Tag = sType
'   If sType = "2" Then
'      frmMain.MailBtn.Caption = "RESEND OR REQUEUE"
'   Else
'      frmMain.MailBtn.Caption = "REQUEUE ALL REPORTS"
'   End If
   
   vData = objTV.ReadNodeData(NodeId)
   If sType = -1 Then
      sType = vData(4)
   End If
      
   If blnHL7db And IsNumeric(NodeId.Text) Then
      strSQL = "SELECT Message " & _
               "FROM HL7_History " & _
               "WHERE HL7_Index = " & NodeId.Text
      RS.Open strSQL, HL7Con, adOpenStatic, adLockReadOnly
      frmMain.LogText = ""
      frmMain.LogText = RS(0)
      RS.Close
   Else
      Tstr = fs.BuildPath(FileLocation(CStr(vData(1)), sType, (NodeId.ForeColor = BPRED)), NodeId.Text)
      
      eClass.FurtherInfo = Tstr
      Open Tstr For Input As #1
      LogFileName = Tstr
      fData = Tstr
      While Not EOF(1)
         Line Input #1, Tstr
         tBuf.Append Tstr
         If Len(Tstr) > 0 Then
            tBuf.Append vbCrLf
         End If
      Wend
      
      With frmMain.LogText
         .Text = ""
         .Text = tBuf.value
      End With
      Close (1)
   End If
   
   frmMain.OpenFileBtn.Visible = True
   Set RS = Nothing
   
AfterError:
   With frmMain
      If NodeId.Children <= 1 Then
         .MailBtn.Visible = False
      End If
      .MailBtn.Visible = True
   End With

   fView.Show Fra_LOGVIEW, fData
   Exit Sub
   
procEH:
'   If eClass.Behaviour = -1 Then
'      If (Err.Number <> 76 And Err.Number <> 71) Then
'         Stop
'         Resume
'      End If
'   End If
   frmMain.LogText.Text = ""
   frmMain.LogText.Text = "     >>>    Unable to display file    <<<" & vbCrLf & _
                          "            **********************" & _
                          vbCrLf & vbCrLf & " File path: " & vData(1) & vbCrLf & _
                          " Error Reported: " & Err.Number & vbCrLf & _
                          " Description: " & Err.Description
   blnNoFile = True
   Resume AfterError
'   eClass.CurrentProcedure = "LoadLogs.DisplayFile"
'   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub DisplayReport(NodeId As Node)
   On Error GoTo procEH
   Dim natCode As String
   Dim resRange As String
   Dim repDateTime As String
   Dim repType As String
   Dim RepId As Long
   Dim sampId() As Long
   Dim blnComment As Boolean
   Dim blnStandalone As Boolean
   Dim rCode As String
   Dim maxLen As Integer
   Dim TestLen As Integer
   Dim K As Integer
   Dim MaxRES As Integer
   Dim MaxUOM As Integer
   Dim maxRange As Integer
   Dim TabStr As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim j As Integer
   Dim sid As Integer
   Dim errRet As Long
   Dim objRTF As rtfSelection
   Dim collRTF As New Collection
   Dim iceCmd As New ADODB.Command
   Dim icecmd2 As New ADODB.Command
   Dim repClass As Integer
   Dim DownloadInd As Integer
   Dim hpData As New clsHealthParties
   
   vData = objTV.ReadNodeData(NodeId)
   RepId = vData(0)
   frmMain.MailBtn.Caption = "REQUEUE THIS REPORT"
   frmMain.MailBtn.Visible = True
   
   If RepId > 0 Then
   
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICECONFIG_ReportLocalTrader"
         .Parameters.Append .CreateParameter("ReportId", adInteger, adParamInput, , RepId)
         Set RS = .Execute
         
         logLTSIndex = RS!EDI_LTS_Index
         intUseRCIndex = RS!UseLabReadCodes
         RS.Close
      End With
         
      With iceCmd
         .CommandText = "ICELABCOMM_Report_Patient"
         Set RS = .Execute
         natCode = Trim(RS!EDI_NatCode & "") & " (" & RS!EDI_Local_Key1 & ")"
         repDateTime = RS!DateTime_Of_Report
         repType = RS!Service_Report_Type
         blnComment = RS!Comment_Marker
         repClass = Val(RS!Report_Class & "")
         DownloadInd = Val(RS!GP_Download & "")
      End With
      
      With frmMain.LogText
         .Visible = False
         
         .Text = ""
         .Text = vbCrLf & Trim(RS!Surname) & ", " & Trim(RS!Forename) & vbTab & RS!Date_Of_Birth & "  " & IIf(RS!Sex = 1, "Male", IIf(RS!Sex = 2, "Female", "Unknown"))
         If Trim(RS!Pat_Addr_Line1) <> "" Then
            .Text = .Text & vbCrLf & "<<" & Trim(RS!Pat_Addr_Line1) & ", "
         End If
         If Trim(RS!Pat_Addr_Line2) <> "" Then
            .Text = .Text & Trim(RS!Pat_Addr_Line2) & ", "
         End If
         If Trim(RS!Pat_Addr_Line3) <> "" Then
            .Text = .Text & Trim(RS!Pat_Addr_Line3) & ", "
         End If
         If Trim(RS!Pat_Addr_Line4) <> "" Then
            .Text = .Text & Trim(RS!Pat_Addr_Line4) & ", "
         End If
         If Trim(RS!Pat_Addr_Line5) <> "" Then
            .Text = .Text & Trim(RS!Pat_Addr_Line5) & ", "
         End If
         If Trim(RS!Pat_Postcode) <> "" Then
            .Text = .Text & Trim(RS!Pat_Postcode) & ">>"
         End If
         .Text = .Text & vbCrLf & "NHS No.: " & vbTab & RS!New_Nhs_No
         .Text = .Text & vbCrLf & "Hosp. No.: " & vbTab & RS!Hospital_Number
         .Text = .Text & vbTab & "Service ID:  " & RS!Service_Report_Id
         RS.Close
         
         hpData.IndividualIndex = -1
         hpData.DownloadRequest = DownloadInd
         hpData.Read RepId, False
         
'         With iceCmd
'            .CommandText = "ICELABCOMM_Report_HealthParties"
'            Set RS = .Execute
'         End With
         
'         RS.Find "EDI_HP_Type = '902'"
         .Text = .Text & vbCrLf & "Clinician/Specialty: " & hpData.HP902Name & " (" & hpData.HP902Code & ")"
         .Text = .Text & vbCrLf & "Destination: " & natCode
         
         If hpData.DownloadRequest > 0 Then
            .Text = .Text & vbTab & "##" & "  GP Download to " & hpData.Practice & "  ##"
         End If
         
         .Text = .Text & vbCrLf & "Report Date: " & repDateTime
         .Text = .Text & vbTab & "Type: " & repType & vbCrLf
'         RS.Close
      End With
      
      strSQL = "SELECT Colour_Code " & _
               "FROM Service_Tubes_ColouRS " & _
               "WHERE Colour_Name IN (" & _
                  "SELECT Report_Colour " & _
                  "FROM Service_Reports_ColouRS " & _
                  "WHERE Report_Type IN (" & _
                     "SELECT Specialty_Code " & _
                     "FROM Specialty " & _
                     "WHERE Specialty LIKE '" & repType & "'))"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
      If RS.RecordCount > 0 Then
   '             LogBackColour = RS2!colour_code
         frmMain.LogText.BackColor = RS!Colour_Code
      End If
         
      RS.Close
         
      iceCmd.CommandText = "ICELABCOMM_Report_Sample"
   '   iceCmd.ParameteRS.Delete (1)
      Set RS = iceCmd.Execute
      If RS.RecordCount > 0 Then
         i = 1
         With frmMain.LogText
'            If hpData.ReportClass = 3 Then
'               .Text = .Text & vbCrLf & vbTab & "## GP Downloaded Report ##" & vbCrLf & vbCrLf
'            End If
            
            Do Until RS.EOF
               ReDim Preserve sampId(i)
               sampId(i - 1) = RS!Sample_Index
               .Text = .Text & vbCrLf & "[" & i & "] Sample Details: " & RS!Sample_Code & " (" & RS!Sample_Text & ")" & vbCrLf
               .Text = .Text & vbTab & "Collected: " & RS!Collection_DateTime
               .Text = .Text & vbTab & "Received: " & RS!Collection_DateTimeReceived & vbCrLf
               i = i + 1
               RS.MoveNext
            Loop
            RS.Close
         End With
         
         If blnComment Then
            With iceCmd
               .CommandText = "ICELABCOMM_Report_Comments"
      '         .ParameteRS.Append .CreateParameter("Type", adVarChar, adParamInput, 1, "R")
               Set RS = .Execute
      '         .ParameteRS.Delete 1
            End With
            With frmMain.LogText
               Do Until RS.EOF
                  .Text = .Text & vbCrLf & vbTab & RS!Service_Report_Comment
                  RS.MoveNext
               Loop
            End With
            RS.Close
         End If
         
         With iceCmd
            If intUseRCIndex > 0 Then
               .CommandText = "ICELABCOMM_Report_Invest_By_Index"
            Else
               .CommandText = "ICELABCOMM_Report_Invest_By_Code"
            End If
            .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
         End With
         
         With frmMain.LogText
            For sid = 0 To UBound(sampId) - 1
               With iceCmd
'                  .CommandText = "ICELABCOMM_Report_Investigation"
                  .Parameters("LTSIndex").value = logLTSIndex
                  .Parameters("SampId").value = sampId(sid)
                  Set RS = .Execute
               End With
            
               If iceCmd.Parameters("Abnormal").value Then
                  Set objRTF = New rtfSelection
                  objRTF.TextPosition = Len(.Text) + 2
                  objRTF.TextLength = 29
                  objRTF.TextColour = BPRED
                  .Text = .Text & vbCrLf & "CONTAINS OUT OF RANGE RESULTS" & vbCrLf
                  collRTF.Add objRTF
                  Set objRTF = Nothing
               End If
               
               maxLen = 0
               MaxUOM = 0
               MaxRES = 0
               maxRange = 0
               
               Do Until RS.EOF
                  If IsNull(RS!Read_V2RC) Then
                     rCode = "[Not Coded]"
                  Else
                     rCode = vbTab & "[" & RS!Read_V2RC & "]"
                     If RS!EDI_Op_Active = False Then
                        rCode = rCode & " INACTIVE"
                     End If
                  
                     If IsNull(RS!Result_Index) Or RS!Result_Recs > 1 Then
                        If RS!Read_Battery = "F" Then
                           Set objRTF = New rtfSelection
                           objRTF.TextPosition = Len(.Text) + Len(Trim(RS!Investigation_Requested)) + 4
                           rCode = rCode & " (Warning - only applicable to a test)"
                           objRTF.TextLength = Len(rCode)
                           objRTF.TextColour = BPBLUE
                           collRTF.Add objRTF
                           Set objRTF = Nothing
                        End If
                     
                        If RS!Read_Status = "D" Then
                           Set objRTF = New rtfSelection
                           objRTF.TextPosition = Len(Trim(RS!Investigation_Requested)) + 4
                           rCode = rCode & " (Error - Read code deleted)"
                           objRTF.TextLength = Len(rCode)
                           objRTF.TextColour = BPBLUE
                           collRTF.Add objRTF
                           Set objRTF = Nothing
                        End If
                     Else
                        rCode = " [See result for read code if applicable]"
                     End If
                  End If
                  
                  With frmMain.LogText
                     .Text = .Text & vbCrLf & vbCrLf & RS!Investigation_Requested & vbTab & rCode
                  End With
                  
                  If RS!Comment_Marker Then
                     With icecmd2
                        .ActiveConnection = iceCon
                        .CommandType = adCmdStoredProc
                        .CommandText = "ICELABCOMM_Report_InvestComments"
                        .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RS!Investigation_Index)
                        Set RS2 = .Execute
                     End With
                     
                     With frmMain.LogText
                        Do Until RS2.EOF
                           .Text = .Text & vbCrLf & vbTab & RS2!Service_Investigation_Comment
                           RS2.MoveNext
                        Loop
                     End With
                     RS2.Close
                  End If
                  Set icecmd2 = New ADODB.Command
                  With icecmd2
                     .ActiveConnection = iceCon
                     .CommandType = adCmdStoredProc
                     If intUseRCIndex > 0 Then
                        .CommandText = "ICELABCOMM_Report_Results_By_Index"
                     Else
                        .CommandText = "ICELABCOMM_Report_Results_By_Code"
                     End If
                     .Parameters.Append .CreateParameter("InvId", adInteger, adParamInput, , RS!Investigation_Index)
                     .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , logLTSIndex)
                     Set RS2 = .Execute
                  End With
                  
                  Do Until RS2.EOF
                     If Len(Trim(RS2!Result_Rubric)) > maxLen Then
                         maxLen = Len(Trim(RS2!Result_Rubric)) + 1
                     End If
                     
                     If Len(Trim(RS2!UOM_Code)) > MaxUOM Then
                           MaxUOM = Len(Trim(RS2!UOM_Code)) + 1
                     End If
                     
                     If Len(Trim(RS2!Result)) > MaxRES Then
                           MaxRES = Len(Trim(RS2!Result)) + 1
                     End If
                     
                     If RS2!Lower_Range = "" And RS2!Upper_Range = "" Then
                        resRange = ""
                     ElseIf Trim(RS2!Lower_Range) = "" Then
                        resRange = "<" & Trim(RS2!Upper_Range)
                     ElseIf Trim(RS2!Upper_Range) = "" Then
                        resRange = ">" & Trim(RS2!Lower_Range)
                     Else
                        resRange = Trim(RS2!Lower_Range) & " to " & Trim(RS2!Upper_Range)
                     End If
                     
                     If Len(resRange) > maxRange Then
                        maxRange = Len(resRange) + 1
                     End If
                     RS2.MoveNext
                  Loop
                  
                  If RS2.RecordCount > 0 Then
                     RS2.MoveFirst
                  End If
                     
                  Do Until RS2.EOF
                     .Text = .Text & vbCrLf & "  " & RS2!Result_Rubric
                              
                     TabStr = Space(maxLen - Len(Trim(RS2!Result_Rubric)))
                     
                     If RS2!Abnormal_Flag Then
                        Set objRTF = New rtfSelection
                        objRTF.TextPosition = Len(.Text) - Len(Trim(RS2!Result_Rubric))
                        objRTF.TextColour = BPRED
                        .Text = .Text & TabStr & " *" & RS2!Result
                     Else
                        .Text = .Text & TabStr & "  " & RS2!Result
                     End If
                                 
                     TabStr = Space(MaxRES - Len(Trim(RS2!Result)))
                     
                     .Text = .Text & TabStr & RS2!UOM_Code
                     TabStr = Space(MaxUOM - Len(Trim(RS2!UOM_Code)))
                              
                     If Trim(RS2!Lower_Range & "") = "" And Trim(RS2!Upper_Range & "") = "" Then
                        resRange = ""
                     ElseIf Trim(RS2!Lower_Range) = "" Then
                        resRange = "<" & Trim(RS2!Upper_Range)
                     ElseIf Trim(RS2!Upper_Range) = "" Then
                        resRange = ">" & Trim(RS2!Lower_Range)
                     Else
                        resRange = Trim(RS2!Lower_Range) & " to " & Trim(RS2!Upper_Range)
                     End If
                     
                     .Text = .Text & TabStr & resRange
                     TabStr = Space(maxRange - Len(resRange))
                              
         '            TabStr = ""
         '            TabStr = Space(MaxUOM - Len(Trim(RS2!Deviation)))
                     
                     If IsNull(RS2!Read_V2RC) Then
                        rCode = ""
                     Else
                        rCode = TabStr & "[" & RS2!Read_V2RC & "]"
                        
                        If RS2!EDI_Op_Active = False Then
                           rCode = rCode & " INACTIVE"
                        End If
                        
                        If RS2!EDI_OP_Suppress Then
                           rCode = rCode & " SUPPRESSED"
                        End If
                        
                        If RS2!Read_Test = "F" Then
                           rCode = rCode & " (Warning - only applicable to a battery header)"
                        End If
                           
                        If RS2!Read_Status = "D" Then
                           rCode = rCode & " (Error - Read code deleted)"
                        End If
                        
                        If RS2!Read_Ratio <> "T" Then
                           If IsNumeric(RS2!Result) Then
                              If RS2!UOM_Code = "" And _
                                 RS2!EDI_OP_UOM = "" Then
                                 rCode = rCode & " (NOT OUTPUT - UOM required with result)"
                              End If
                           End If
                        End If
                     End If
                           
                     .Text = .Text & rCode
                     
                     If RS2!Abnormal_Flag Then
                        objRTF.TextLength = Len(.Text) - objRTF.TextPosition
                        collRTF.Add objRTF
                        Set objRTF = Nothing
                     End If
                     
                     If RS2!Comment_Marker Then
                        Set icecmd2 = Nothing
                        Set icecmd2 = New ADODB.Command
                        With icecmd2
                           .ActiveConnection = iceCon
                           .CommandType = adCmdStoredProc
                           .CommandText = "ICELABCOMM_Report_ResultComment"
                           .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RS2!Result_Index)
                           Set RS3 = .Execute
                        End With
                        
                        Do Until RS3.EOF
                           .Text = .Text & vbCrLf & "   " + RS3!Service_Result_Comment
                           RS3.MoveNext
                        Loop
                        RS3.Close
                     End If
                     RS2.MoveNext
                  Loop
                  RS2.Close
                  Set icecmd2 = Nothing
                  RS.MoveNext
               Loop
               RS.Close
            Next sid
              
            If collRTF.Count > 0 Then
               .Text = .Text & " "
               For i = 1 To collRTF.Count
                  Set objRTF = collRTF.item(1)
                  .SelStart = objRTF.TextPosition
                  .SelLength = objRTF.TextLength
                  .SelColor = objRTF.TextColour
                  Debug.Print .SelStart
                  Debug.Print .SelLength
                  collRTF.Remove (1)
               Next i
               .SelStart = Len(.Text) - 1
               .SelLength = 1
               .SelColor = vbBlack
            End If
            Set objRTF = Nothing
         End With
      Else
         If repClass = 2 Then
            frmMain.LogText.Text = frmMain.LogText.Text & vbCrLf & "     * * *   Discharge Notification   * * *"
         Else
            frmMain.LogText.Text = frmMain.LogText.Text & vbCrLf & "      << No Sample Data available >>"
         End If
      End If
      
     frmMain.LogText.Visible = True
   Else
      frmMain.LogText = vbCrLf & "<< No report data available >>"
   End If
   
   Set RS3 = Nothing
   Set RS2 = Nothing
   Set RS = Nothing
   Exit Sub
    
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.DisplayReport"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewDisplayReport(NodeId As Node)
   On Error GoTo procEH
   Dim natCode As String
   Dim resRange As String
   Dim repDateTime As String
   Dim repType As String
   Dim RepId As Long
   Dim sampId() As Long
   Dim blnComment As Boolean
   Dim blnStandalone As Boolean
   Dim rCode As String
   Dim maxLen As Integer
   Dim TestLen As Integer
   Dim K As Integer
   Dim MaxRES As Integer
   Dim MaxUOM As Integer
   Dim maxRange As Integer
   Dim TabStr As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim j As Integer
   Dim sid As Integer
   Dim errRet As Long
   Dim objRTF As rtfSelection
   Dim collRTF As New Collection
   Dim iceCmd As New ADODB.Command
   Dim icecmd2 As New ADODB.Command
   Dim repClass As Integer
   Dim hpData As New clsHealthParties
   Dim sBuf As New StringBuffer
   
   vData = objTV.ReadNodeData(NodeId)
   RepId = vData(0)
   frmMain.MailBtn.Caption = "REQUEUE THIS REPORT"
   frmMain.MailBtn.Visible = True
   
   If RepId > 0 Then
   
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICECONFIG_ReportLocalTrader"
         .Parameters.Append .CreateParameter("ReportId", adInteger, adParamInput, , RepId)
         Set RS = .Execute
         
         logLTSIndex = RS!EDI_LTS_Index
         intUseRCIndex = RS!UseLabReadCodes
         RS.Close
      End With
         
      With iceCmd
         .CommandText = "ICELABCOMM_Report_Patient"
         Set RS = .Execute
         natCode = Trim(RS!EDI_NatCode & "") & " (" & RS!EDI_Local_Key1 & ")"
         repDateTime = RS!DateTime_Of_Report
         repType = RS!Service_Report_Type
         blnComment = RS!Comment_Marker
         repClass = Val(RS!Report_Class & "")
      End With
      
      With sBuf
         .Append vbCrLf & Trim(RS!Surname) & ", " & Trim(RS!Forename) & vbTab & RS!Date_Of_Birth & "  " & IIf(RS!Sex = 1, "Male", IIf(RS!Sex = 2, "Female", "Unknown"))
         If Trim(RS!Pat_Addr_Line1) <> "" Then
            .Append vbCrLf & "<<" & Trim(RS!Pat_Addr_Line1) & ", "
         End If
         If Trim(RS!Pat_Addr_Line2) <> "" Then
            .Append Trim(RS!Pat_Addr_Line2) & ", "
         End If
         If Trim(RS!Pat_Addr_Line3) <> "" Then
            .Append Trim(RS!Pat_Addr_Line3) & ", "
         End If
         If Trim(RS!Pat_Addr_Line4) <> "" Then
            .Append Trim(RS!Pat_Addr_Line4) & ", "
         End If
         If Trim(RS!Pat_Addr_Line5) <> "" Then
            .Append Trim(RS!Pat_Addr_Line5) & ", "
         End If
         If Trim(RS!Pat_Postcode) <> "" Then
            .Append Trim(RS!Pat_Postcode) & ">>"
         End If
         
         .Append vbCrLf & "NHS No.: " & vbTab & RS!New_Nhs_No
         .Append vbCrLf & "Hosp. No.: " & vbTab & RS!Hospital_Number
         .Append vbTab & "Service ID:  " & RS!Service_Report_Id
         
         RS.Close
         
         hpData.IndividualIndex = -1
         hpData.Read RepId, False
         
         .Append vbCrLf & "Clinician/Specialty: " & hpData.HP902Name & " (" & hpData.HP902Code & ")"
         .Append vbCrLf & "Destination: " & natCode
         .Append vbCrLf & "Report Date: " & repDateTime
         .Append vbTab & "Type: " & repType & vbCrLf
      End With
      
      strSQL = "SELECT Colour_Code " & _
               "FROM Service_Tubes_ColouRS " & _
               "WHERE Colour_Name IN (" & _
                  "SELECT Report_Colour " & _
                  "FROM Service_Reports_ColouRS " & _
                  "WHERE Report_Type IN (" & _
                     "SELECT Specialty_Code " & _
                     "FROM Specialty " & _
                     "WHERE Specialty LIKE '" & repType & "'))"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
      If RS.RecordCount > 0 Then
         frmMain.LogText.BackColor = RS!Colour_Code
      End If
         
      RS.Close
         
      iceCmd.CommandText = "ICELABCOMM_Report_Sample"
      Set RS = iceCmd.Execute
      If RS.RecordCount > 0 Then
         i = 1
         With sBuf
            Do Until RS.EOF
               ReDim Preserve sampId(i)
               sampId(i - 1) = RS!Sample_Index
               .Append vbCrLf & "[" & i & "] Sample Details: " & RS!Sample_Code & " (" & RS!Sample_Text & ")" & vbCrLf
               .Append vbTab & "Collected: " & RS!Collection_DateTime
               .Append vbTab & "Received: " & RS!Collection_DateTimeReceived & vbCrLf
               i = i + 1
               RS.MoveNext
            Loop
            RS.Close
         End With
         
         If blnComment Then
            With iceCmd
               .CommandText = "ICELABCOMM_Report_Comments"
               Set RS = .Execute
            End With
            
            With sBuf
               Do Until RS.EOF
                  .Append vbCrLf & vbTab & RS!Service_Report_Comment
                  RS.MoveNext
               Loop
            End With
            
            RS.Close
         End If
         
         With iceCmd
            If intUseRCIndex > 0 Then
               .CommandText = "ICELABCOMM_Report_Invest_By_Index"
            Else
               .CommandText = "ICELABCOMM_Report_Invest_By_Code"
            End If
            .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
         End With
         
         With sBuf
            For sid = 0 To UBound(sampId) - 1
               With iceCmd
                  .Parameters("LTSIndex").value = logLTSIndex
                  .Parameters("SampId").value = sampId(sid)
                  Set RS = .Execute
               End With
            
               If iceCmd.Parameters("Abnormal").value Then
                  Set objRTF = New rtfSelection
                  objRTF.TextPosition = Len(.ActualValue) + 2
                  objRTF.TextLength = 29
                  objRTF.TextColour = BPRED
                  .Append vbCrLf & "CONTAINS OUT OF RANGE RESULTS" & vbCrLf
                  collRTF.Add objRTF
                  Set objRTF = Nothing
               End If
               
               maxLen = 0
               MaxUOM = 0
               MaxRES = 0
               maxRange = 0
               
               Do Until RS.EOF
                  If IsNull(RS!Read_V2RC) Then
                     rCode = "[Not Coded]"
                  Else
                     rCode = vbTab & "[" & RS!Read_V2RC & "]"
                     If RS!EDI_Op_Active = False Then
                        rCode = rCode & " INACTIVE"
                     End If
                  
                     If IsNull(RS!Result_Index) Or RS!Result_Recs > 1 Then
                        If RS!Read_Battery = "F" Then
                           Set objRTF = New rtfSelection
                           objRTF.TextPosition = Len(.ActualValue) + Len(Trim(RS!Investigation_Requested)) + 4
                           rCode = rCode & " (Warning - only applicable to a test)"
                           objRTF.TextLength = Len(rCode)
                           objRTF.TextColour = BPBLUE
                           collRTF.Add objRTF
                           Set objRTF = Nothing
                        End If
                     
                        If RS!Read_Status = "D" Then
                           Set objRTF = New rtfSelection
                           objRTF.TextPosition = Len(Trim(RS!Investigation_Requested)) + 4
                           rCode = rCode & " (Error - Read code deleted)"
                           objRTF.TextLength = Len(rCode)
                           objRTF.TextColour = BPBLUE
                           collRTF.Add objRTF
                           Set objRTF = Nothing
                        End If
                     Else
                        rCode = " [See result for read code if applicable]"
                     End If
                  End If
                  
                  sBuf.Append vbCrLf & vbCrLf & RS!Investigation_Requested & vbTab & rCode
                  
                  If RS!Comment_Marker Then
                     With icecmd2
                        .ActiveConnection = iceCon
                        .CommandType = adCmdStoredProc
                        .CommandText = "ICELABCOMM_Report_InvestComments"
                        .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RS!Investigation_Index)
                        Set RS2 = .Execute
                     End With
                     
                     With sBuf
                        Do Until RS2.EOF
                           .Append vbCrLf & vbTab & RS2!Service_Investigation_Comment
                           RS2.MoveNext
                        Loop
                     End With
                     
                     RS2.Close
                  End If
                  
                  Set icecmd2 = New ADODB.Command
                  With icecmd2
                     .ActiveConnection = iceCon
                     .CommandType = adCmdStoredProc
                     If intUseRCIndex > 0 Then
                        .CommandText = "ICELABCOMM_Report_Results_By_Index"
                     Else
                        .CommandText = "ICELABCOMM_Report_Results_By_Code"
                     End If
                     .Parameters.Append .CreateParameter("InvId", adInteger, adParamInput, , RS!Investigation_Index)
                     .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , logLTSIndex)
                     Set RS2 = .Execute
                  End With
                  
                  Do Until RS2.EOF
                     If Len(Trim(RS2!Result_Rubric)) > maxLen Then
                         maxLen = Len(Trim(RS2!Result_Rubric)) + 1
                     End If
                     
                     If Len(Trim(RS2!UOM_Code)) > MaxUOM Then
                           MaxUOM = Len(Trim(RS2!UOM_Code)) + 1
                     End If
                     
                     If Len(Trim(RS2!Result)) > MaxRES Then
                           MaxRES = Len(Trim(RS2!Result)) + 1
                     End If
                     
                     If RS2!Lower_Range = "" And RS2!Upper_Range = "" Then
                        resRange = ""
                     ElseIf Trim(RS2!Lower_Range) = "" Then
                        resRange = "<" & Trim(RS2!Upper_Range)
                     ElseIf Trim(RS2!Upper_Range) = "" Then
                        resRange = ">" & Trim(RS2!Lower_Range)
                     Else
                        resRange = Trim(RS2!Lower_Range) & " to " & Trim(RS2!Upper_Range)
                     End If
                     
                     If Len(resRange) > maxRange Then
                        maxRange = Len(resRange) + 1
                     End If
                     RS2.MoveNext
                  Loop
                  
                  If RS2.RecordCount > 0 Then
                     RS2.MoveFirst
                  End If
                     
                  Do Until RS2.EOF
                     .Append vbCrLf & "  " & RS2!Result_Rubric
                              
                     TabStr = Space(maxLen - Len(Trim(RS2!Result_Rubric)))
                     
                     If RS2!Abnormal_Flag Then
                        Set objRTF = New rtfSelection
                        objRTF.TextPosition = Len(.ActualValue) - Len(Trim(RS2!Result_Rubric))
                        objRTF.TextColour = BPRED
                        .Append TabStr & " *" & RS2!Result
                     Else
                        .Append TabStr & "  " & RS2!Result
                     End If
                                 
                     TabStr = Space(MaxRES - Len(Trim(RS2!Result)))
                     
                     .Append TabStr & RS2!UOM_Code
                     TabStr = Space(MaxUOM - Len(Trim(RS2!UOM_Code)))
                              
                     If Trim(RS2!Lower_Range & "") = "" And Trim(RS2!Upper_Range & "") = "" Then
                        resRange = ""
                     ElseIf Trim(RS2!Lower_Range) = "" Then
                        resRange = "<" & Trim(RS2!Upper_Range)
                     ElseIf Trim(RS2!Upper_Range) = "" Then
                        resRange = ">" & Trim(RS2!Lower_Range)
                     Else
                        resRange = Trim(RS2!Lower_Range) & " to " & Trim(RS2!Upper_Range)
                     End If
                     
                     .Append TabStr & resRange
                     TabStr = Space(maxRange - Len(resRange))
                              
                     If IsNull(RS2!Read_V2RC) Then
                        rCode = ""
                     Else
                        rCode = TabStr & "[" & RS2!Read_V2RC & "]"
                        
                        If RS2!EDI_Op_Active = False Then
                           rCode = rCode & " INACTIVE"
                        End If
                        
                        If RS2!EDI_OP_Suppress Then
                           rCode = rCode & " SUPPRESSED"
                        End If
                        
                        If RS2!Read_Test = "F" Then
                           rCode = rCode & " (Warning - only applicable to a battery header)"
                        End If
                           
                        If RS2!Read_Status = "D" Then
                           rCode = rCode & " (Error - Read code deleted)"
                        End If
                        
                        If RS2!Read_Ratio <> "T" Then
                           If IsNumeric(RS2!Result) Then
                              If RS2!UOM_Code = "" And _
                                 RS2!EDI_OP_UOM = "" Then
                                 rCode = rCode & " (NOT OUTPUT - UOM required with result)"
                              End If
                           End If
                        End If
                     End If
                           
                     .Append rCode
                     
                     If RS2!Abnormal_Flag Then
                        objRTF.TextLength = Len(.ActualValue) - objRTF.TextPosition
                        collRTF.Add objRTF
                        Set objRTF = Nothing
                     End If
                     
                     If RS2!Comment_Marker Then
                        Set icecmd2 = Nothing
                        Set icecmd2 = New ADODB.Command
                        With icecmd2
                           .ActiveConnection = iceCon
                           .CommandType = adCmdStoredProc
                           .CommandText = "ICELABCOMM_Report_ResultComment"
                           .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RS2!Result_Index)
                           Set RS3 = .Execute
                        End With
                        
                        Do Until RS3.EOF
                           .Append vbCrLf & "   " + RS3!Service_Result_Comment
                           RS3.MoveNext
                        Loop
                        RS3.Close
                     End If
                     RS2.MoveNext
                  Loop
                  
                  RS2.Close
                  Set icecmd2 = Nothing
                  RS.MoveNext
               Loop
               
               RS.Close
            Next sid
         End With
              
         With frmMain.LogText
            .Text = sBuf.ActualValue & " "
            
            If collRTF.Count > 0 Then
               For i = 1 To collRTF.Count
                  Set objRTF = collRTF.item(1)
                  .SelStart = objRTF.TextPosition
                  .SelLength = objRTF.TextLength
                  .SelColor = objRTF.TextColour
                  Debug.Print .SelStart
                  Debug.Print .SelLength
                  collRTF.Remove (1)
               Next i
               .SelStart = Len(.Text) - 1
               .SelLength = 1
               .SelColor = vbBlack
            End If
         End With
         Set objRTF = Nothing
      Else
         If repClass = 2 Then
            frmMain.LogText.Text = frmMain.LogText.Text & vbCrLf & "     * * *   Discharge Notification   * * *"
         Else
            frmMain.LogText.Text = frmMain.LogText.Text & vbCrLf & "      << No Sample Data available >>"
         End If
      End If
      
     frmMain.LogText.Visible = True
   Else
      frmMain.LogText = vbCrLf & "<< No report data available >>"
   End If
   
   Set RS3 = Nothing
   Set RS2 = Nothing
   Set RS = Nothing
   Exit Sub
    
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.DisplayReport"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub PatientSearch(PatientId As String, _
                         Optional Forename As String = "") ', _
                         Optional StartDate As Date = 0, _
                         Optional EndDate As Date = 0)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim PatId As String
   Dim nd(2) As Node
   Dim strName As String
   
   'LogFileName = ""
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("PatId", adVarChar, adParamInput, 35, PatientId)
      .Parameters.Append .CreateParameter("StartDate", adVarChar, adParamInput, 19, StartDate)
      .Parameters.Append .CreateParameter("EndDate", adVarChar, adParamInput, 19, EndDate)
      
      If frmMain.optLogSearch(2).value Then
         .Parameters.Append .CreateParameter("Forename", adVarChar, adParamInput, 35, Forename)
         .CommandText = "ICECONFIG_Search_PatientName"
      ElseIf frmMain.optLogSearch(3).value Then
         If New_NHSno_Check(PatientId) = 1 Then
            .CommandText = "ICECONFIG_Search_PatientNHS"
         Else
            .CommandText = "ICECONFIG_Search_PatientHosp"
         End If
      End If
      
      Set RS = .Execute
   End With
   
   If RS.RecordCount = 0 Then
      Set tvNode(0) = tv1.Nodes.Add(, _
                                    , _
                                    mCtrl.NewNodeKey("None", _
                                                     "None", _
                                                     "NotPresent"), _
                                    "No reports found", _
                                    13)
   Else
      Do Until RS.EOF
         strName = RS!Surname & " " & RS!Forename & " " & RS!Date_Of_Birth & " ("
         If (RS!New_Nhs_No & "") <> "" Then
            strName = strName & "NHS: " & RS!New_Nhs_No & "/"
         End If
         
         strName = strName & "Hosp: " & RS!Hospital_Number & ")"
         Set tvNode(0) = tv1.Nodes.Add(, _
                                       , _
                                       mCtrl.NewNodeKey(RS!Patient_Id_Key, _
                                                        "None", _
                                                        "PatientReport"), _
                                       strName, _
                                   13)
         tv1.Nodes.Add tvNode(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey("Temp", _
                                        "Temp", _
                                        "None"), _
                       "Please wait...", _
                       13
         RS.MoveNext
      Loop
   End If
   
   RS.Close
   Set RS = Nothing
   Set iceCmd = Nothing
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.PatientSearch"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub LabSearch(LabNo As String) ', _
                     Optional StartDate As Date = 0, _
                     Optional EndDate As Date = 0)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim svcId As String
   Dim displayId As String
   
   'LogFileName = ""
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("RepId", adVarChar, adParamInput, 35, LabNo)
      
      .CommandText = "ICECONFIG_Search_Lab"
      .Parameters.Append .CreateParameter("sDate", adVarChar, adParamInput, 19, StartDate)
      .Parameters.Append .CreateParameter("eDate", adVarChar, adParamInput, 19, EndDate)
      
      Set RS = .Execute
   End With
   
   If RS.RecordCount = 0 Then
      Set tvNode(0) = tv1.Nodes.Add(, _
                                    , _
                                    mCtrl.NewNodeKey("None", _
                                                     "None", _
                                                     "NotPresent"), _
                                    "No reports found", _
                                    13)
   Else
      Do Until RS.EOF
         If svcId <> RS!Service_Report_Id Then
            svcId = RS!Service_Report_Id
            displayId = Left(RS!Service_Report_Id, InStr(1, RS!Service_Report_Id, " "))
            Set tvNode(0) = tv1.Nodes.Add(, _
                                          , _
                                          mCtrl.NewNodeKey(RS!Service_Report_Index, _
                                                           RS!Service_Report_Id, _
                                                           "NotPresent"), _
                                          RS!Service_Report_Id & " (" & RS!Surname & " " & RS!Forename & ")", _
                                          13)
         End If
         
         strSQL = "SELECT Service_ImpExp_Message_Id " & _
                  "FROM Service_ImpExp_HeadeRS sh " & _
                     "INNER JOIN Service_ImpExp_Messages sm " & _
                     "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
                  "WHERE Service_Report_Index = " & RS!Service_Report_Index & _
                  "ORDER BY sh.Date_Added DESC"

         RS2.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         
         Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
                                        tvwChild, _
                                        mCtrl.NewNodeKey(RS!Service_Report_Index, _
                                                         RS2!Service_ImpExp_Message_ID, _
                                                         "PatientFiles", _
                                                         RS!Date_Added), _
                                        "Added " & RS!Date_Added, _
                                        13)
         tv1.Nodes.Add tvNode(1), _
                       tvwChild, _
                       mCtrl.NewNodeKey("Temp", _
                                        "Temp", _
                                        "None"), _
                       "Please wait...", _
                       13
         RS2.Close
         RS.MoveNext
      Loop
   End If
   
   RS.Close
   Set RS = Nothing
   Set RS2 = Nothing
   Set iceCmd = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   
   End If
   eClass.CurrentProcedure = "LoadLogs.LabSearch"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub
'
'Public Sub LabRequest(NodeId As Node)
'   Dim iceCmd As New ADODB.Command
'   Dim RS As New ADODB.Recordset
'   Dim RS2 As New ADODB.Recordset
'   Dim i As Integer
'
'   vData = objTV.ReadNodeData(NodeId)
'
'   With iceCmd
'      .ActiveConnection = icecon
'      .CommandType = adCmdStoredProc
'      iceCmd.CommandText = "ICECONFIG_Search_PatientFiles"
'      .ParameteRS.Append .CreateParameter("ReqId", adVarChar, adParamInput, 10, vData(0))
'      Set RS = .Execute
'   End With
'
'
'   For i = 1 To NodeId.Children
'      tv1.Nodes.Remove NodeId.Child.Index
'   Next i
'
'   Do Until RS.EOF
'      Set tvNode(0) = tv1.Nodes.Add(NodeId, _
'                                tvwChild, _
'                                mCtrl.NewNodeKey(RS!Service_Report_Index, _
'                                                 "none", _
'                                                 "SearchDisplayReport"), _
'                                "Added " & RS!Date_Added, _
'                                13)
'
'      With iceCmd
'         .ParameteRS("ReqId").value = RS!Service_Report_Index
'         Set RS2 = .Execute
'      End With
'
'      RS2.Filter = "Service_Type = 1"
'      Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
'                                    tvwChild, _
'                                    mCtrl.NewNodeKey(RS!EDI_LTS_Index, _
'                                                     "LAB", _
'                                                     "None"), _
'                                    "Laboratory Reports", _
'                                    13)
'
'      Do Until RS2.EOF
'         Set tvNode(2) = tv1.Nodes.Add(tvNode(1), _
'                                       tvwChild, _
'                                       mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
'                                                        fs.GetParentFolderName(RS2!ImpExp_File), _
'                                                        "SearchDisplayFile", _
'                                                        RS!Service_Type), _
'                                       fs.GetFileName(RS!ImpExp_File), _
'                                       13)
'
'         If RS!Error_Found > 0 Or _
'            RS!Status_Flag > 0 Then
'            tvNode(2).ForeColor = BPRED
'            tvNode(2).Image = 26
'            tvNode(2).SelectedImage = 26
'         ElseIf RS!Status_Flag < 0 Then
'            tvNode(2).ForeColor = BPBLUE
'         Else
'            tvNode(2).ForeColor = BPGREEN
'         End If
'
'         tv1.Nodes.Add tvNode(2), _
'                       tvwChild, _
'                       mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
'                                        RS!Service_Type, _
'                                        "Tracking"), _
'                       "Tracking", _
'                       13, _
'                       13
'
'         RS2.MoveNext
'      Loop
'      RS2.MoveFiRSt
'
'      RS2.Filter = "Service_Type = 2"
'      Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
'                                    tvwChild, _
'                                    mCtrl.NewNodeKey(RS2!EDI_LTS_Index, _
'                                                     "LAB", _
'                                                     "None"), _
'                                    "EDI Reports", _
'                                    13)
'
'      Do Until RS2.EOF
'         tv1.Nodes.Add tvNode(1), _
'                       tvwChild, _
'                       mCtrl.NewNodeKey(RS2!Service_ImpExp_Id, _
'                                        fs.GetParentFolderName(RS2!ImpExp_File), _
'                                        "SearchDisplayFile", _
'                                        RS2!Service_Type), _
'                       fs.GetFileName(RS2!ImpExp_File), _
'                       13
'         RS2.MoveNext
'      Loop
'
'      RS2.Close
'      RS.MoveNext
'   Loop
'
'   RS.Close
'   Set RS = Nothing
'   Set iceCmd = Nothing
'End Sub

Public Sub PatientFiles(NodeId As Node)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim sType As String
   Dim i As Integer
   
   vData = objTV.ReadNodeData(NodeId)
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_Search_PatientFiles"
      .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , vData(0))
      Set RS = .Execute
   End With

   logLTSIndex = RS!EDI_LTS_Index
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   Do Until RS.EOF
      If RS!Service_Type <> sType Then
         Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                       tvwChild, _
                                       mCtrl.NewNodeKey(RS!EDI_LTS_Index, _
                                                        RS!Service_Type, _
                                                        "None"), _
                                       RS!Description & " (" & RS!EDI_Msg_Type & ")", _
                                       13, _
                                       13)
         sType = RS!Service_Type
      End If
      
      Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                                 fs.GetParentFolderName(RS!ImpExp_File), _
                                                 "FileDets", _
                                                 RS!Service_Type), _
                                fs.GetFileName(RS!ImpExp_File), _
                                13)
      
      If RS!Error_Found > 0 Or _
         RS!Status_Flag > 0 Then
         tvNode(1).ForeColor = BPRED
         tvNode(1).Image = 26
         tvNode(1).SelectedImage = 26
      ElseIf RS!Status_Flag < 0 Then
         tvNode(1).ForeColor = BPBLUE
      Else
         tvNode(1).ForeColor = BPGREEN
      End If
      
      tv1.Nodes.Add tvNode(1), _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                     RS!Service_Type, _
                                     "Tracking"), _
                    "Tracking", _
                    13, _
                    13
      
      RS.MoveNext
   Loop
   
   DisplayReport NodeId
   RS.Close
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.FurtherInfo = "LodaLogs.PatientFiles"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub PatientReport(NodeId As Node)
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim FileId As String
   Dim i As Integer
   
   vData = objTV.ReadNodeData(NodeId)
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_Search_PatientReports"
      .Parameters.Append .CreateParameter("Index", adVarChar, adParamInput, 30, vData(0))
      .Parameters.Append .CreateParameter("sDate", adVarChar, adParamInput, 19, StartDate)
      .Parameters.Append .CreateParameter("eDate", adVarChar, adParamInput, 19, EndDate)
      Set RS = .Execute
   End With
   
'   iceCmd.CommandText = "ICECONFIG_SearchFiles"
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   Do Until RS.EOF
      strSQL = "SELECT Service_ImpExp_Message_Id " & _
               "FROM Service_ImpExp_HeadeRS sh " & _
                  "INNER JOIN Service_ImpExp_Messages sm " & _
                  "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
               "WHERE Service_Report_Index = " & RS!Service_Report_Index & _
               "ORDER BY sh.Date_Added DESC"
'                  " AND Service_Type = 2 " & _

      RS2.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      
      Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Service_Report_Index, _
                                                 RS2!Service_ImpExp_Message_ID, _
                                                 "PatientFiles"), _
                                Trim(RS!Service_Report_Id) & " - Added " & RS!Date_Added, _
                                13)
      RS2.Close
      
      tv1.Nodes.Add tvNode(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Temp", _
                                     "Temp", _
                                     "None"), _
                    "Please wait...", _
                    13
      RS.MoveNext
   Loop
   
   
   RS.Close
   Set RS2 = Nothing
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Patient(PatientId, _
                   Optional Forename As String = "", _
                   Optional StartDate As Date = 0, _
                   Optional EndDate As Date = 0, _
                   Optional dataType As Integer = 0)

   On Local Error GoTo procEH
'  A  dataType of 0 = Patient Name, 1 = NHS Number, 2 = Hospital Number
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim OrgID As String
   Dim iceCmd As New ADODB.Command
   Dim dbv As Variant
   Dim strSQL As String
   Dim thisPatient As String
   Dim thisDate As String
   Dim thisFile As String
   Dim strForeName As String
   Dim stRSurname As String
   Dim i As Integer
   Dim j As Integer
   Dim sNode() As Node
   Dim curRep As String
   Dim TempStr As String
   Dim fDir As String
   Dim tIndex As Long
   Dim lType As String

   ReDim sNode(0)
   OrgID = frmMain.cboTrust.Text
   
   Set iceCmd = Nothing
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_PatientData"
      .Parameters.Append .CreateParameter("NHSorLocal", adInteger, adParamInput, , New_NHSno_Check(PatientId))
      .Parameters.Append .CreateParameter("lookUp", adVarChar, adParamInput, 10, Left(PatientId, 10))
   End With
   
   Set dbv = iceCmd
   
   RS.Open dbv, , adOpenKeyset, adLockReadOnly
   If RS.BOF = False And RS.EOF = False Then
      stRSurname = Trim(RS!Surname)
      strForeName = Trim(RS!Forename)
   Else
      stRSurname = Replace(PatientId, "'", "`")
      strForeName = Replace(Forename, "'", "`")
   End If
   RS.Close
      
   Set RS = Nothing
   Set iceCmd = Nothing
   
   strSQL = "SELECT TOP 100 PERCENT Service_ImpExp_Messages.Service_ImpExp_ID, " & _
                                    "Service_ImpExp_Messages.Patient_Name, " & _
                                    "Service_ImpExp_Messages.Patient_DOB, " & _
                                    "Service_ImpExp_Messages.Service_Report_Index, " & _
                                    "Service_ImpExp_Messages.Service_ImpExp_Message_ID, " & _
                                    "Service_ImpExp_Messages.Patient_Local_ID, " & _
                                    "Service_ImpExp_Messages.Report_Identifier, " & _
                                    "Service_ImpExp_HeadeRS.Date_Added, " & _
                                    "Service_ImpExp_HeadeRS.ImpExp_File, " & _
                                    "Service_ImpExp_HeadeRS.Service_Type, " & _
                                    "Service_ImpExp_HeadeRS.Status_Flag, " & _
                                    "Service_Id, " & _
                                    "Service_ImpExp_HeadeRS.Error_Found, " & _
                                    "Service_ImpExp_HeadeRS.Status_Flag " & _
            "FROM Service_ImpExp_Messages INNER JOIN " & _
               "Service_ImpExp_HeadeRS ON Service_ImpExp_Messages.Service_ImpExp_ID = Service_ImpExp_HeadeRS.Service_ImpExp_ID " & _
            "WHERE (Service_ImpExp_Messages.Patient_Name IN " & _
               "(SELECT DISTINCT Patient_Name " & _
               "FROM Service_ImpExp_Messages " & _
               "WHERE Patient_Name LIKE '% " & strForeName & "%' AND Patient_Name IN " & _
                  "(SELECT Patient_Name " & _
                  "FROM Service_ImpExp_Messages " & _
                  "WHERE Patient_Name LIKE '" & stRSurname & "%'))) "
                     
   If StartDate > 0 Or EndDate > 0 Then
      strSQL = myDateForm & strSQL & "AND (Service_ImpExp_Messages.Service_ImpExp_ID IN " & _
                           "(SELECT    Service_ImpExp_Id " & _
                           "FROM Service_ImpExp_HeadeRS " & _
                           "WHERE "
      If StartDate > 0 Then
         strSQL = strSQL & "datediff(Dd,'" & Format(StartDate, "dd/mm/yyyy") & "',Date_Added) >= 0 "
         If EndDate > 0 Then
            strSQL = strSQL & "AND "
         Else
            strSQL = strSQL & "))"
         End If
      End If
      
      If EndDate > 0 Then
         strSQL = strSQL & "datediff(Dd,Date_Added,'" & Format(EndDate, "dd/mm/yyyy") & "') >= 0)) "
      End If
   End If
   
   strSQL = strSQL & "ORDER BY Patient_Local_ID, Service_ImpExp_HeadeRS.Date_Added" & defDateForm
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.RecordCount = 0 Then
      tv1.Nodes.Add , _
                    , _
                    mCtrl.NewNodeKey(0, _
                                     "Not Found", _
                                     "None"), _
                                     "No Data Found"
   Else
      Do Until RS.EOF
         If Trim(RS!Patient_Local_Id) <> thisPatient Then
            thisPatient = Trim(RS!Patient_Local_Id)
            Set tvNode(0) = tv1.Nodes.Add(, _
                                          , _
                                          mCtrl.NewNodeKey(RS!Service_Report_Index, _
                                                           RS!Service_ImpExp_Message_ID, _
                                                           "RepDetails"), _
                                          RS!Patient_Name & " " & RS!Patient_DOB & " " & thisPatient, _
                                          13, _
                                          13)
            thisDate = ""
            thisFile = ""
         End If
            
         If thisDate <> Format(RS!Date_Added, "dd/mm/yyyy") Then
            For i = 0 To UBound(sNode)
               If TypeName(sNode(i)) <> "Nothing" Then
                  If sNode(i).Children = 0 Then
                     tv1.Nodes.Remove (sNode(i).Index)
                  End If
               End If
            Next i
            thisDate = Format(RS!Date_Added, "dd/mm/yyyy")
            Set tvNode(1) = tv1.Nodes.Add(tvNode(0), _
                                          tvwChild, _
                                          mCtrl.NewNodeKey(thisDate, _
                                                           ""), _
                                          thisDate, _
                                          13, _
                                          13)
            
            strSQL = "SELECT * " & _
                     "FROM Service_Types " & _
                     "ORDER BY Type_Index"
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            For i = 0 To RS2.RecordCount - 1
               ReDim Preserve sNode(i)
               Set sNode(i) = tv1.Nodes.Add(tvNode(1), _
                                            tvwChild, _
                                            mCtrl.NewNodeKey(RS2!Type_Index, _
                                                             ""), _
                                            RS2!Description, _
                                            13, _
                                            13)
               RS2.MoveNext
            Next i
            RS2.Close
         End If
         
         If thisFile <> Trim(RS!ImpExp_File) Then
            thisFile = Trim(RS!ImpExp_File)
            If Left(RS!ImpExp_File, 1) = "+" Then
               lType = "5"
               TempStr = Mid(RS!ImpExp_File, 2)
               fDir = RS!Date_Added
            Else
               TempStr = Trim(fs.GetFileName(Trim(RS!ImpExp_File)))
               lType = "2"
               fDir = fs.GetParentFolderName(Trim(RS!ImpExp_File))
            End If
            
            
            Set tvNode(2) = tv1.Nodes.Add(sNode(RS!Service_Type - 1), _
                                          tvwChild, _
                                          mCtrl.NewNodeKey(RS!Service_ImpExp_Id, _
                                                           fDir, _
                                                           "Reports"), _
                                          TempStr, _
                                          13, _
                                          13)
            If RS!Error_Found > 0 Then
               tvNode(2).ForeColor = BPRED
               tvNode(2).Image = 26
               tvNode(2).SelectedImage = 26
            ElseIf RS!Status_Flag < 0 Then
               tvNode(2).ForeColor = BPBLUE
            Else
               tvNode(2).ForeColor = BPGREEN
            End If
            tv1.Nodes.Add tvNode(2), _
                          tvwChild, _
                          mCtrl.NewNodeKey("Tmp", _
                                           "Tmp"), _
                          "Please wait...", _
                          13, _
                          13
            
'            tv1.Nodes.Add tvNode(2), _
'                          tvwChild, _
'                          mCtrl.NewNodeKey(3, RS!Service_ImpExp_ID), _
'                          "Tracking", _
'                          13, _
'                          13
         End If
'         tv1.Nodes.Add tvNode(2), _
'                       tvwChild, _
'                       mCtrl.NewNodeKey(4, OrgID & "+" & _
'                                           Trim(RS!Service_Report_Index) & "+" & _
'                                           CStr(RS!Service_ImpExp_Message_ID)), _
'                       Left(RS!Service_Id, 14), _
'                       13, 13
         RS.MoveNext
      Loop
      For i = 0 To UBound(sNode)
         If TypeName(sNode(i)) <> "Nothing" Then
            If sNode(i).Children = 0 Then
'            Debug.Print sNode(i).Text & " Removed from " & sNode(i).Parent.Text & "(from " & sNode(i).Parent.Parent.Text & ")"
               tv1.Nodes.Remove (sNode(i).Index)
            End If
         End If
      Next i
   End If
   tv1.Refresh
   tv1.Visible = True
'   BuildTreeView
   Set RS = Nothing
   Set RS2 = Nothing
'   Set RS3 = Nothing
'   Set RS4 = Nothing
'   Set RS5 = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "NewTreeClass.LoadPatient"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub RepDets(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   
   vData = objTV.ReadNodeData(NodeId)
   
   With frmMain
      .LogText.Text = ""
      DisplayReport NodeId
      .MailBtn.Caption = "REQUEUE THIS REPORT"
      .MailBtn.Visible = True
      strSQL = "SELECT * " & _
               "FROM Service_ImpExp_Messages " & _
               "WHERE Service_ImpExp_Message_Id = " & vData(1) & _
                  " AND Service_Report_Index = " & vData(0)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS!Rejected_Count > 0 Then
         .LogText.Text = "Error: "
      ElseIf RS!Rejected_Count < 0 Then
         .LogText.Text = "Comment: "
         .MailBtn.Visible = False
      End If
      .LogText.Text = .LogText.Text & vbCrLf & RS!Service_Message_Comment & vbCrLf
      RS.Close
   End With
   Set RS = Nothing
   'DisplayReport NodeId
'   Call Load_Report(OrgStr, IDStr)
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.RepDets"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub RequeueOptions(OptRequired As String)
   On Error GoTo procEH
   Dim tNode As Node
   Dim intNoGP As Integer
   Dim totReps As Integer
   Dim blnRequeue As Boolean
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim strAction As String
   Dim wMsg As String
   Dim fStr As String
   Dim fType As Integer
   
   If blnUseDMO Then
      Set tNode = objTV.ActiveNode.child
      rCtrl.CallerForm = frmMain
'      frmRequeue.FileDirectory = ""
      
      Select Case OptRequired
         Case "REQUEUE ALL REPORTS"
            rCtrl.RequeueReport = True
            Set tNode = tNode.Next
            Do Until tNode Is Nothing
               vData = objTV.ReadNodeData(tNode)
               rCtrl.RequeueItem = vData(1)
               Set tNode = tNode.Next
            Loop
            
            rCtrl.RequeueData
            
         Case "REQUEUE THIS REPORT"
            rCtrl.RequeueReport = True
            rCtrl.RequeueItem = objTV.nodeKey(objTV.ActiveNode)
            rCtrl.RequeueData
         
         Case "RESEND OR REQUEUE"
            
            With frmEDIRequeue
               If blnNoFile Then
                  .optRequeue(1).value = True
                  .optRequeue(0).Enabled = False
               Else
                  .optRequeue(0).Enabled = True
                  .optRequeue(0).value = True
               End If
               .Show 1
            End With
            
            strAction = frmEDIRequeue.RequeueValue
            If strAction = "Resend" Then
               rCtrl.RequeueItem = vData(0)
               rCtrl.RequeueReport = False
               
               strSQL = "SELECT * " & _
                        "FROM Service_ImpExp_HeadeRS " & _
                        "WHERE Service_ImpExp_Id = " & vData(0)
               RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
               fStr = RS!ImpExp_File
               fType = RS!Service_Type
               RS.Close
               rCtrl.RequeueFile fStr, fType
            
            ElseIf strAction = "Reprocess" Then
               rCtrl.RequeueReport = True
               Set tNode = tNode.Next
               Do Until tNode Is Nothing
                  vData = objTV.ReadNodeData(tNode)
                  rCtrl.RequeueItem = vData(1)
                  Set tNode = tNode.Next
               Loop
               rCtrl.RequeueData
            End If
      
      End Select
   Else
      MsgBox "Requeue option not available on this machine." & vbCrLf & _
             "The 'SQLDMO' option has not been installed." & vbCrLf & _
             "Please ask your IT Support dept to have the option installed" & _
             " or use the management machine to requeue reports or files", _
             vbExclamation, "Required SQL option not present"
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   iceCon.RollbackTrans
   eClass.CurrentProcedure = "LoadLogs.RequeueOptions"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub RunWhat(NodeId As Node, _
                   Optional RefreshId As String = "")
   On Local Error GoTo procEH
   Dim vData As Variant
   Dim rType As String
   Dim tNode As MSComctlLib.Node
   
   If inProgress = False Then
      vData = objTV.ReadNodeData(NodeId)
      NodeId.Selected = True
      frmMain.LogText.BackColor = &HC0E0FF
      LogFileName = ""
      frmMain.OpenFileBtn.Visible = False
      fView.FrameToShow = Fra_LOGVIEW
      blnShowBrowser = True
      fView.FrameToShow = fra_None
      frmMain.MailBtn.Visible = False
      
      Select Case vData(2)
         Case "Dates"
            Dates NodeId
            blnShowBrowser = True
            fView.FrameToShow = fra_None
            fView.Show fra_None
         
         Case "Files"
            Files NodeId
            
         Case "FilesByTime"
            FilesByTime NodeId
            
         Case "FileDets"
            DisplayFile NodeId
            fView.FrameToShow = Fra_LOGVIEW
            blnShowBrowser = False
            
         Case "LabRequest"
   '         LabRequest NodeId
            
         Case "Reports"
            frmMain.OpenFileBtn.Visible = True
            Reports NodeId
            fView.FrameToShow = Fra_LOGVIEW
            blnShowBrowser = False
            
         Case "RepDetails"
            RepDets NodeId
            fView.FrameToShow = Fra_LOGVIEW
            blnShowBrowser = False
         
         Case "Tracking"
            Tracking NodeId
            fView.FrameToShow = Fra_LOGVIEW
            blnShowBrowser = False
            
         Case "PatientReport"
            PatientReport NodeId
            
         Case "PatientFiles"
            PatientFiles NodeId
            fView.FrameToShow = Fra_LOGVIEW
            blnShowBrowser = False
            
         Case "SearchDisplayReport"
            SearchDisplayReport NodeId
            
         Case "SearchDisplayFile"
            DisplayFile NodeId
            
         Case Else
            blnShowBrowser = True
            fView.FrameToShow = fra_None
            
      End Select
   End If
Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub Times(NodeId As Node)
   Dim batchTot As Integer
   Dim sBatch As String
   Dim eBatch As String
   Dim iLoop As Integer
   Dim errTot As Integer
   Dim totRecs As Integer
   Dim totNodes As Integer
   Dim lastNode As Integer
   Dim descStr As String
   Dim RS As New ADODB.Recordset
   
   RS.Open rsFile, , , , adCmdFile
   
   With frmWait
      .fraBusy.Visible = False
      .ProgressBar1.Visible = True
      .ProgressBar1.max = Int(RS.RecordCount / maxFilesPerNode)
      .Label1.Caption = "Building tree view..."
      .Show
      .Refresh
   End With
   
   DoEvents
   
   iLoop = 1
   
   tv1.Visible = False
   
   Do While Not RS.EOF
      If iLoop = 1 Then
         sBatch = RS!Service_ImpExp_Id
         descStr = "From " & TimeSerial(Hour(RS!Date_Added), Minute(RS!Date_Added), Second(RS!Date_Added)) & " to "
      End If
         
      errTot = errTot + CInt(RS!Error_Found)
      
      If iLoop = maxFilesPerNode Then
         eBatch = RS!Service_ImpExp_Id
         descStr = descStr & TimeSerial(Hour(RS!Date_Added), Minute(RS!Date_Added), Second(RS!Date_Added))
         
         Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                       tvwChild, _
                                       mCtrl.NewNodeKey(sBatch, _
                                                        eBatch, _
                                                        "FilesByTime", _
                                                        fileDate), _
                                       descStr, _
                                       13, _
                                       13)
         
         If errTot > 0 Then
            tvNode(0).Image = 26
         End If
         
         tv1.Nodes.Add tvNode(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey("Tmp", _
                                        "Tmp"), _
                       "Please wait...", _
                       13, _
                       13
         
         iLoop = 0
         errTot = 0
         totNodes = totNodes + 1
      
         With frmWait
            .Label1.Caption = "Building tree view...(" & totNodes & " records)"
            .ProgressBar1.value = totNodes
            .Refresh
         End With
         
         DoEvents
      End If
      
      RS.MoveNext
      iLoop = iLoop + 1
   Loop
   
'**************************************************
'  Add the last group to the treeview
'**************************************************

   If iLoop > 1 Then '  The recordset is not a multiple of 400 records
      RS.MovePrevious
      eBatch = RS!Service_ImpExp_Id
      descStr = descStr & TimeSerial(Hour(RS!Date_Added), Minute(RS!Date_Added), Second(RS!Date_Added))
      
      Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                    tvwChild, _
                                    mCtrl.NewNodeKey(sBatch, _
                                                     eBatch, _
                                                     "FilesByTime", _
                                                     fileDate), _
                                    descStr, _
                                    13, _
                                    13)
      
      If errTot > 0 Then
         tvNode(0).Image = 26
      End If
      
      tv1.Nodes.Add tvNode(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", _
                                     "Tmp"), _
                    "Please wait...", _
                    13, _
                    13
   End If
   
   RS.Close
   Set RS = Nothing
      
   Unload frmWait
   
   tv1.Visible = True
   frmMain.MousePointer = vbNormal
   inProgress = False

End Sub

Public Sub Tracking(NodeId As Node)
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim Tstr As String
   Dim strTab As String
   Dim objRTF As rtfSelection
   
   Tstr = objTV.NodeLevel(NodeId)
   vData = objTV.ReadNodeData(NodeId)
   strSQL = "SELECT Service_ImpExp_Comment, Status_Flag, Service_ImpExp_Comments.Date_Added " & _
            "FROM Service_ImpExp_Comments " & _
               "INNER JOIN Service_ImpExp_HeadeRS ON " & _
               "Service_ImpExp_Comments.Service_impExp_Id = Service_ImpExp_HeadeRS.Service_impExp_Id " & _
            "WHERE Service_ImpExp_Comments.Service_ImpExp_Id = " & Val(Tstr) & " " & _
               "AND Service_Type <= " & vData(1) & _
            " ORDER BY Service_ImpExp_Comment_Id"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   With frmMain
      With .LogText
         .SelStart = 0
         .SelLength = Len(.Text)
         .SelColor = vbBlack
         .Text = ""
      End With
      
      If RS.RecordCount > 0 Then
         If Trim(RS!Status_Flag & "") <> "" Then
            strSQL = "SELECT Status_Message " & _
                     "FROM Status_Codes " & _
                     "WHERE Status_Code = " & Trim(RS!Status_Flag & "")
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            Set objRTF = New rtfSelection
            objRTF.TextPosition = 0
            
            .LogText.Text = "Status: " & RS!Status_Flag & " (" & RS2!Status_Message & ")" & vbCrLf & vbCrLf
            
            objRTF.TextLength = Len(.LogText.Text) - objRTF.TextPosition

            If RS!Status_Flag < 0 Then
               objRTF.TextColour = BPBLUE
            ElseIf RS!Status_Flag = 0 Then
               objRTF.TextColour = BPGREEN
            Else
               objRTF.TextColour = BPRED
            End If
            
            RS2.Close
         End If
         
         strTab = vbTab & vbTab & vbTab & vbTab
         Do While Not RS.EOF
            .LogText.Text = .LogText.Text & _
                            "[" & RS!Date_Added & "] " & Trim(RS!Service_ImpExp_Comment)
            .LogText.Text = .LogText.Text & vbCrLf & strTab & "***" & vbCrLf
            RS.MoveNext
         Loop
         RS.Close
         If Not objRTF Is Nothing Then
            With .LogText
               .Text = .Text & " "
               .SelStart = objRTF.TextPosition
               .SelLength = objRTF.TextLength
               .SelColor = objRTF.TextColour
               .SelStart = Len(.Text) - 1
               .SelLength = 1
               .SelColor = vbBlack
            End With
            Set objRTF = Nothing
         End If
         .MailBtn.Caption = ""
         .MailBtn.Visible = False
      End If
   End With
   Set RS2 = Nothing
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadLogs.Tracking"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub Class_Terminate()
   'HL7Con.Close
   iceAsyncCon.Close
   Set HL7Con = Nothing
   Set iceAsyncCon = Nothing
End Sub

Private Sub iceAsyncCon_ExecuteComplete(ByVal RecordsAffected As Long, _
                                        ByVal pError As ADODB.Error, _
                                        adStatus As ADODB.EventStatusEnum, _
                                        ByVal pCommand As ADODB.Command, _
                                        ByVal pRecordset As ADODB.Recordset, _
                                        ByVal pConnection As ADODB.Connection)

   'Dim rsFile As String
   Dim tFile As String
   Dim noRecs As Long
   
   If adStatus = adStatusErrorsOccurred Then
      Debug.Print "Error: " & pError
      inProgress = False
      
      If pError = "Operation has been cancelled by the user." Then
         execProc = "Error"
      Else
         If iceAsyncCon.Errors.Count > 0 Then
            Debug.Print "Errors found"
            If iceAsyncCon.Errors(0).SQLState <> "HYT00" Then
               Err.Raise iceAsyncCon.Errors(0).Number, iceAsyncCon.Errors(0).Description
            Else
               iceAsyncCon.Cancel
               Set asyncRS = Nothing
               Set asyncRS = New ADODB.Recordset
            End If
         Else
            Debug.Print "No Errors"
            iceAsyncCon.Cancel
            Set asyncRS = Nothing
            Set asyncRS = New ADODB.Recordset
         End If
         
         rtc = rtc + 1
         frmWait.Label1.Caption = "Database Timeout - Retrying " & rtc
         DoEvents
         
         If asyncRS.State <> adStateClosed Then
            asyncRS.Close
         End If
         
         If rtc < 4 Then
            asyncRS.Open pCommand.CommandText, iceAsyncCon, adOpenStatic, adLockReadOnly, adAsyncExecute
         
            frmWait.Refresh
         Else
            MsgBox "Unable to read database - " & pError & vbCrLf & _
                   "Either increase the timeout value or try again later", vbInformation, "Database retries exceeded"
            
            Dim fStr As TextStream
            Set fStr = fs.CreateTextFile(App.Path & "\log.sql")
            fStr.Write pCommand.CommandText
            fStr.Close
            
            rtc = 0
            'execProc = ""
            Unload frmWait
            If curNode.Children = 0 Then
               tv1.Nodes.Add curNode, _
                             tvwChild, _
                             mCtrl.NewNodeKey("Tmp", _
                                              "Tmp"), _
                             "Failed - Database Timeout = " & conTimeOut & " secs", _
                             13, _
                             13
            Else
               curNode.child.Text = "Failed - Database Timeout = " & conTimeOut & " secs"
            End If
         End If
      End If
   Else
      rtc = 0
      If execProc <> "" Then
         frmWait.Label1.Caption = "Fetch complete - Saving data"
         frmWait.Refresh
         DoEvents
         
         noRecs = asyncRS.RecordCount
         If fs.FileExists(rsFile) Then
            fs.DeleteFile rsFile
         End If
         
         asyncRS.Save rsFile, adPersistADTG
         asyncRS.Close
         Set asyncRS.ActiveConnection = Nothing
         asyncRS.Open rsFile
         
         Select Case execProc
            Case "Dates"
               DatesForServiceType curNode
      
            Case "Files"
               'asyncRS.Open rsFile, , , , adCmdFile
            
               If noRecs > maxFilesPerNode Then
                  Times curNode
'                  tFile = fs.BuildPath(App.Path, fileDate & "_log.RS")
'
'                  If fs.FileExists(tFile) Then
'                     fs.DeleteFile tFile, True
'                  End If
'
'                  fs.MoveFile rsFile, tFile
               Else
                  FilesForDay curNode
               End If
      
         End Select
      End If
   End If
   
   inProgress = False
End Sub

Private Sub iceAsyncCon_WillConnect(ConnectionString As String, userID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
   'Debug.Print "Will Connect"
End Sub

Private Sub iceAsyncCon_WillExecute(Source As String, _
                                    CursorType As ADODB.CursorTypeEnum, _
                                    LockType As ADODB.LockTypeEnum, _
                                    Options As Long, _
                                    adStatus As ADODB.EventStatusEnum, _
                                    ByVal pCommand As ADODB.Command, _
                                    ByVal pRecordset As ADODB.Recordset, _
                                    ByVal pConnection As ADODB.Connection)
   
   iceAsyncCon.CommandTimeout = conTimeOut
   'Debug.Print "WillExecute (Timeout: " & iceAsyncCon.CommandTimeout & ")"
   If inProgress Then
      adStatus = adStatusCancel
      Debug.Print "Will execute - In Progress"
   Else
      inProgress = True
      adStatus = adStatusOK
   End If
End Sub

Private Sub iceAsyncCon_InfoMessage(ByVal pError As ADODB.Error, _
                                    adStatus As ADODB.EventStatusEnum, _
                                    ByVal pConnection As ADODB.Connection)
   MsgBox pError.Description, vbCritical, "Error detected"
End Sub


