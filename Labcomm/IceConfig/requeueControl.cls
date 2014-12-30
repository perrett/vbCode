VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "requeueControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private parentForm As Form
Private pBar As ProgressBar
Private stepValue As Long
Private inPath As String
Private pendPath As String
Private outPath As String
Private histPath As String
Private errPath As String
Private svcPath(2) As String
Private fileSrc As String
Private fileDest As String

Private nd(2) As Node
Private rq As RequeueData
Private outDir As String
Private reportNotFile As Boolean
Private UDLConnectData As String
Private UDLServer As String
Private UDLDatabase As String
Private strSQL As String

Private EDIReps As SQLDMO.Table
Private SvcComments As SQLDMO.Table
Private wrkTab As SQLDMO.Table
Private wrkCol As SQLDMO.Column
Private pKey As SQLDMO.Key
Private colTNames As SQLDMO.Names

Private blnCodeAfterDate As Boolean

Private colRepList As New Collection
Private colRequeue As New Collection
Private colKey As New Collection

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim colPos As Integer

   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_ReadTable"
      .Parameters.Append .CreateParameter("TableId", adVarChar, adParamInput, 50, "EDI_Rep_List")
      Set RS = .Execute
   End With

   Do Until RS.EOF
      colRepList.Add CStr(RS!Name), "Col_" & RS!colId
      RS.MoveNext
   Loop

   RS.Close
   Set iceCmd = Nothing

'  Set up application directories
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Module_Directories"
      .Parameters.Append .CreateParameter("ModId", adVarChar, adParamInput, 20, "IceMsg.exe")
      Set RS = .Execute
      
'     Ignore if none set up
      If Not RS.EOF Then
         pendPath = RS!Connection_PendingDirs
         outPath = RS!Connection_TargetDirectory
         svcPath(1) = outPath
         histPath = RS!Connection_HistoryDirs
         errPath = RS!Connection_ErrorDirs
      End If
      
      RS.Close
      
      .Parameters(0).value = "IceImp.exe"
      Set RS = .Execute
      
      If Not RS.EOF Then
         svcPath(0) = RS!Connection_TargetDirectory
      End If
   End With
   
   RS.Close
   
   If blnUseDMO Then
      Set EDIReps = sqlDb.Tables("EDI_Rep_List")
      Set SvcComments = sqlDb.Tables("Service_ImpExp_Comments")
      
'      strSQL = "SELECT sc.name, colorder " & _
'                "FROM sysobjects so " & _
'                "   INNER JOIN syscolumns sc " & _
'                "   ON so.id = sc.id " & _
'                "WHERE so.name='service_impexp_comments' " & _
'                "ORDER BY colorder"
'      RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
'
'      RS.Filter = "name='Date_Added'"
'      colPos = RS!colorder
'      RS.Filter = "name='Code'"
'      blnCodeAfterDate = (Int(RS!colorder) > colPos)
'
'      RS.Close
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "requeueControl.Class_Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let CallerForm(frmId As Form)
   Set parentForm = frmId
End Property

Public Function CreateWorkTable(BasedOn As SQLDMO.Table) As SQLDMO.Table
   On Error GoTo procEH
   Dim tTable As SQLDMO.Table
   Dim tmpTab As SQLDMO.Table
   Dim tCol() As SQLDMO.Column
   Dim oCol As SQLDMO.Column
   Dim NoOfCols As Integer
   Dim i As Integer
   
   NoOfCols = BasedOn.Columns.Count
   
   ReDim tCol(NoOfCols)
   
   For i = 1 To NoOfCols
      If BasedOn.Columns(i).Identity = False Then
         Set oCol = BasedOn.Columns(i)
         Set tCol(i) = New SQLDMO.Column
         tCol(i).Name = oCol.Name
         tCol(i).dataType = oCol.dataType
         tCol(i).AllowNulls = oCol.AllowNulls
         tCol(i).Length = oCol.Length
      End If
   Next
   
   Set tTable = CreateObject("SQLDMO.Table") 'New SQLDMO.Table
   tTable.Name = "cfgTMP_" & BasedOn.Name
   tTable.FileGroup = "PRIMARY"
   tTable.Owner = "dbo"
   
   For i = 1 To NoOfCols
      If BasedOn.Columns(i).Identity = False Then
         tTable.Columns.Add tCol(i)
      End If
   Next i
   
   sqlDb.Tables.Add tTable
   tTable.TruncateData
   Set CreateWorkTable = tTable
   Exit Function
   
procEH:
   If Err.Number = -2147218790 And Err.Source = "Microsoft SQL-DMO (ODBC SQLState: 42S01)" Then
      Set tmpTab = sqlDb.Tables(tTable.Name)
      tmpTab.Remove
      Resume
   End If
   eClass.CurrentProcedure = "RequeueControl.CreateWorkTable"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function
'
'Private Function ReadUDL() As String
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
'   buf = Space(1500)
'   Open DB_UDL_FILE For Binary As #1
'   Get #1, , buf
'   Close #1
'
'   pos = InStr(1, buf, pTag, vbBinaryCompare)
'   UDLConnectData = Replace(Mid(buf, pos), Chr(0), "")
'   constr = UDLConnectData
'
'   pos = InStr(1, constr, "Source=") + 7
'   clen = InStr(pos, constr, vbCrLf) - pos
'   UDLServer = Mid(constr, pos, clen)
'
'   pos = InStr(1, constr, "User ID=") + 8
'   clen = InStr(pos, constr, ";") - pos
'   dbUser = Mid(constr, pos, clen)
'
'   pos = InStr(1, constr, "Password=") + 10
'   clen = InStr(pos, constr, ";") - pos - 1
'   dbPass = Mid(constr, pos, clen)
'
'   pos = InStr(1, constr, "Initial Catalog") + 16
'   clen = InStr(pos, constr, ";") - pos
'   UDLDatabase = Mid(constr, pos, clen)
'
'   ReadUDL = constr
'End Function

Public Property Let RequeueItem(strNewValue As String)
   On Error GoTo procEH
   Dim rq As New RequeueData
   
   rq.Identifier = strNewValue
   colRequeue.Add rq, "Key_" & strNewValue
   Set rq = Nothing
   Exit Property
   
procEH:
   If Err.Number = 457 Then
      Resume Next
   Else
      Stop
      Resume
      eClass.CurrentProcedure = "RequeueControl.RequeueItem"
      eClass.Add Err.Number, Err.F, Err.Source
   End If
End Property

Public Sub RequeueData()
   On Error GoTo procEH
   Dim rq As RequeueData
   Dim i As Integer
   Dim j As Integer
   Dim rVal As Long
   Dim max(3) As Long
   Dim tLen As Long
   Dim RequeueTot As Long
   Dim DupeTot As Long
   Dim noteTot As Long
   Dim totSelected As Long
   Dim strDest As String
   
   Dim rqEDIFile As String
   Dim rqSVCFile As String
   Dim rqERRFile As String
   Dim rlKey As String
   Dim fSep As String
   
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim strRepList As String
   Dim strRepList1 As String
   Dim strComment As String
   Dim strDateTime As String
   Dim lKey As String
   Dim hdrSQL As String
   Dim mStat As Integer
   
   Dim dbBcp As SQLDMO.BulkCopy
   Dim EDIRec As Long
   Dim SVCRec As Long
   Dim rqCount As Long
   Dim msgStat As Long
   
   Dim sBufEDI As New StringBuffer
   Dim sBufSVC As New StringBuffer
   Dim blnInTrans As Boolean
   
   blnInTrans = False
   
   rqSVCFile = fs.BuildPath(App.Path, "Comment_Requeue.txt")
   rqERRFile = fs.BuildPath(App.Path, "SQLErrors.txt")
   fSep = "¬"
   strDateTime = Format(Now(), "yyyymmdd hh:nn:ss")
   
   With frmShowRequeue.fgReq
      .rows = colRequeue.Count + 1
      .TextMatrix(0, 0) = "Report/File"
      .TextMatrix(0, 1) = "Practice"
      .TextMatrix(0, 2) = "Comments"
      .Col = 0
   End With
   
   parentForm.Refresh
   
   If reportNotFile Then
      rqEDIFile = fs.BuildPath(App.Path, "EDI_Requeue.txt")
      
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICECONFIG_RequeueReport"
         .Parameters.Append .CreateParameter("Id", adInteger, adParamInput, , 0)
         
         totSelected = colRequeue.Count
         
         For i = 1 To colRequeue.Count
            If Not pBar Is Nothing Then
               pBar.value = i
            End If
            
            Set rq = colRequeue(i)
            eClass.FurtherInfo = "Message ImpExp: " & rq.Identifier
            .Parameters(0).value = rq.Identifier
            Set RS = .Execute
            
            Load frmShowRequeue
            lKey = "Key_" & RS!Service_Report_Index
            
            With frmShowRequeue.fgReq
               colKey.Add RS!Service_Report_Index, lKey
               .Row = i
               .Col = 0
               .ColSel = 2
               
               If frmShowRequeue.TextWidth(Trim(RS!Service_Id)) > max(0) Then
                  max(0) = frmShowRequeue.TextWidth(Trim(RS!Service_Id))
               End If
               
               If frmShowRequeue.TextWidth(Trim(RS!natCode & "")) > max(1) Then
                  max(1) = frmShowRequeue.TextWidth(Trim(RS!natCode))
               End If
               
               .TextMatrix(i, 0) = RS!Service_Id
               .TextMatrix(i, 1) = Trim(RS!natCode & "")
               
               If IsNull(RS!EDI_Report_Index) Then
                  If IsNull(RS!EDI_Local_Key1) Then
                     .TextMatrix(i, 2) = "No Local Key for " & Trim(RS!Dest)
                     .CellForeColor = BPRED
                  Else
                     If Left(lKey, 3) = "Key" Then
                        .TextMatrix(i, 2) = "To be Requeued"
                        .CellForeColor = BPGREEN
                        RequeueTot = RequeueTot + 1
                        msgStat = CLng(Val(RS!status & "")) Or MS_REQUEUE
                        
                        strRepList = ""
                        For j = 1 To colRepList.Count
                           rlKey = "Col_" & CStr(j)

                           Select Case colRepList(rlKey)
                              Case "EDI_Report_Index"
                                 strRepList = RS!Service_Report_Index & fSep

                              Case "EDI_Provider_Org"
                                 strRepList = strRepList & RS!OrgID & fSep

                              Case "EDI_Loc_Nat_Code_To"
                                 strRepList = strRepList & RS!natCode & fSep

                              Case "EDI_Individual_Index_To"
                                 strRepList = strRepList & RS!To_GP & fSep

                              Case "EDI_Service_ID"
                                 strRepList = strRepList & RS!Service_Id & fSep

                              Case "Date_Added"
                                 strRepList = strRepList & strDateTime & fSep

                              Case "EDI_LTS_Index"
                                 strRepList = strRepList & RS!LTS_Index & fSep

                              Case "EDI_Rep_Specialty"
                                 strRepList = strRepList & RS!Discipline & fSep

                              Case "Extra_Copy_To"
                                 strRepList = strRepList & fSep

                              Case "Report_Class"
                                 strRepList = strRepList & RS!Report_Class & fSep

                           End Select
                        Next j
                        
                        strRepList = Left(strRepList, Len(strRepList) - 1) & vbCrLf
                        
                        'strRepList1 = RS!Service_Report_Index & fSep & _
                                      RS!OrgID & fSep & _
                                      RS!natCode & fSep & _
                                      RS!To_GP & fSep & _
                                      RS!Service_Id & fSep & _
                                      strDateTime & fSep & _
                                      RS!LTS_Index & fSep & _
                                      RS!Discipline & fSep & _
                                      RS!Report_Class & fSep & vbCrLf
                        
                        strComment = RS!ImpExpId & fSep & _
                                     RS!Service_Id & " (" & Trim(RS!Patient_Name) & ") Requeued to: " & RS!natCode & fSep & _
                                     "REQUEUE" & fSep & _
                                     MS_REQUEUE & fSep & _
                                     strDateTime & fSep & _
                                     RS!Service_ImpExp_Message_ID & fSep & _
                                     RS!Service_Report_Index & vbCrLf
                                     
'                        If blnCodeAfterDate Then
'                           strComment = strComment & _
'                                        strDateTime & fSep & _
'                                        MS_REQUEUE & fSep
'
'                        Else
'                           strComment = strComment & _
'                                     MS_REQUEUE & fSep & _
'                                     strDateTime & fSep
'                        End If
                        
'                        strComment = strComment & _
'                                     RS!Service_ImpExp_Message_ID & fSep & _
'                                     RS!Service_Report_Index & vbCrLf
                                     
                        
                        strSQL = strSQL & "UPDATE Service_ImpExp_Headers SET " & _
                                             "Comment_Status = Comment_Status"
                        
                        sBufEDI.Append strRepList
                        sBufSVC.Append strComment
                  
                     Else
                        .TextMatrix(i, 2) = "Duplicated Report"
                        .CellForeColor = BPBLUE
                        DupeTot = DupeTot + 1
                     End If
                  End If
                  
               Else
                  .TextMatrix(i, 2) = "Already on Replist"
                  .CellForeColor = BPBLUE
                  noteTot = noteTot + 1
               End If
               
               If frmShowRequeue.TextWidth(Trim(RS!natCode & "")) > max(2) Then
                  max(2) = frmShowRequeue.TextWidth(Trim(RS!natCode))
               End If
               
            End With
            
            RS.Close
         Next i
         
         frmShowRequeue.Resize max(0), max(1), max(2)
      End With
      
'     A shorter file will not amend the EOF marker when using binary file handling so delete the original
      If fs.FileExists(rqEDIFile) Then
         fs.DeleteFile rqEDIFile
      End If
      
      If fs.FileExists(rqSVCFile) Then
         fs.DeleteFile rqSVCFile
      End If
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If
      
      Open rqEDIFile For Binary As #1
      Put #1, , sBufEDI.ActualValue
      Close #1
      
      Open rqEDIFile & ".bak" For Append As #2
      Print #2, sBufEDI.ActualValue
      Close #2
       
      Open rqSVCFile For Binary As #1
      Put #1, , sBufSVC.ActualValue
      Close #1
      
      
      Set dbBcp = Nothing
      Set dbBcp = CreateObject("SQLDMO.BulkCopy")
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If
      
      With dbBcp
         .MaximumErrorsBeforeAbort = totSelected   '  Allow all record to be in error
         .DataFilePath = rqEDIFile
         .DataFileType = SQLDMODataFile_SpecialDelimitedChar
         .ColumnDelimiter = fSep
         .RowDelimiter = vbCrLf
         .ErrorFilePath = rqERRFile
         .UseExistingConnection = True
      End With

'     Use a temporary table because an Import cannot run in a transaction
      Set wrkTab = CreateWorkTable(EDIReps)
      
'     Set up a primary key to trap duplicate requests
      Set pKey = New SQLDMO.Key
      pKey.Type = SQLDMOKey_Primary
      Set colTNames = pKey.KeyColumns
      colTNames.Add "EDI_Report_Index"
      
      wrkTab.BeginAlter
      wrkTab.Keys.Add pKey
      wrkTab.DoAlter
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If

'     Perform the bulk copy
      EDIRec = wrkTab.ImportData(dbBcp)
'      strSQL = "INSERT EDI_Rep_List SELECT * FROM " & wrkTab.Name & "; DROP TABLE " & wrkTab.Name
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If
      
'     Now do the same with service comments
      Set wrkTab = CreateWorkTable(SvcComments)
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If
      
      Set wrkCol = New SQLDMO.Column
      wrkCol.AllowNulls = True
      wrkCol.dataType = "Int"
      wrkCol.Name = "Service_Report_Index"
      wrkTab.Columns.Add wrkCol
      
'     Add Service_Report_Index coloum to trap duplicates here as well. This column is ignored by the later copy statement
      Set pKey = New SQLDMO.Key
      pKey.Type = SQLDMOKey_Primary
      Set colTNames = pKey.KeyColumns
      wrkCol.AllowNulls = False
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If
      
      colTNames.Add "Service_Report_Index"
      
      wrkTab.BeginAlter
      wrkTab.Keys.Add pKey
      wrkTab.DoAlter
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If
      
      dbBcp.DataFilePath = rqSVCFile
      SVCRec = wrkTab.ImportData(dbBcp)
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.value + stepValue
      End If
      
'      strSQL = strSQL & _
               "; INSERT Service_ImpExp_Comments " & _
                  "SELECT Service_ImpExp_ID, " & _
                  "Service_ImpExp_Comment, " & _
                  "Service_ImpExp_Process, " & _
                  "Code, " & _
                  "Date_Added, " & _
                  "Service_ImpExp_Msg_Id " & _
               "FROM " & wrkTab.Name & "; DROP TABLE " & wrkTab.Name
      
'      Set wrkTab = Nothing
      
'     User confirmirmation
      If SVCRec = EDIRec Then
         With frmShowRequeue
            .txtRequeued = EDIRec
            .txtDuplicate = DupeTot
            .lblNote = "Already on Rep List"
            .txtNote = noteTot
            .Show 1
         End With
         
         parentForm.MousePointer = vbHourglass
         
         If frmShowRequeue.FormAction = "Commit" Then
            iceCon.BeginTrans
            blnInTrans = True
            With iceCmd
               .ActiveConnection = iceCon
               .CommandType = adCmdStoredProc
               .CommandText = "ICECONFIG_RequeueCommit"
               .Parameters.Delete 0
               .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
               .Execute
               rVal = .Parameters("Return").value
            End With
            
            If rVal = 0 Then
               iceCon.CommitTrans
            Else
               iceCon.RollbackTrans
            End If
            
            blnInTrans = False
'            sqlServer.BeginTransaction
'            blnInTrans = True
'            sqlDb.ExecuteImmediate strSQL
'            sqlServer.CommitTransaction
            blnInTrans = False
         End If
      Else
         MsgBox "Replist Inserts (" & EDIRec & ") do not match comment inserts ( " & SVCRec & ")", _
                vbExclamation, "Please call Sunquest on 0845 519 4020"
      End If
   
   Else
      
'*************************************************************
'     Requeue the file(s)
'*************************************************************
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICECONFIG_RequeueFiles"
         .Parameters.Append .CreateParameter("Id", adInteger, adParamInput, , 0)
      End With
      
      If parentForm.Name = "frmLogs" Then
         frmShowRequeue.fgReq.Tag = frmLogs.comboSvcType.ListIndex ' + 1
      End If
      
      hdrSQL = ""
      For i = 1 To colRequeue.Count
         If Not pBar Is Nothing Then
            pBar.value = pBar.value + 1
         End If
            
         Set rq = colRequeue(i)
         With iceCmd
            .Parameters(0).value = rq.Identifier
            Set RS = .Execute
         End With
         
         lKey = "Key" & rq.Identifier
         With frmShowRequeue.fgReq
            .Row = i
            .Col = 0
            .ColSel = 2
            If frmShowRequeue.TextWidth(Trim(RS!ImpExp_File)) > max(0) Then
               max(0) = frmShowRequeue.TextWidth(Trim(RS!ImpExp_File))
            End If
               
            .TextMatrix(i, 0) = RS!ImpExp_File
         
            If frmShowRequeue.TextWidth(Trim(RS!EDI_NatCode & "")) > max(1) Then
               max(1) = frmShowRequeue.TextWidth(Trim(RS!EDI_NatCode & ""))
            End If
               
            .TextMatrix(i, 1) = Trim(RS!EDI_NatCode & "")
            
            If fs.FileExists(fs.BuildPath(outPath, fs.GetBaseName(RS!ImpExp_File))) Then
               .TextMatrix(i, 2) = "Already in directory"
               .CellForeColor = BPBLUE
               DupeTot = DupeTot + 1
               
            ElseIf fs.FileExists(RS!ImpExp_File) = False Then
               .TextMatrix(i, 2) = "No Source file"
               .CellForeColor = BPRED
               noteTot = noteTot + 1
            
            Else
               .TextMatrix(i, 2) = "Ready to Copy"
               .CellForeColor = BPGREEN
               EDIRec = EDIRec + 1
               
               sBufSVC.Append fSep & _
                              rq.Identifier & fSep & _
                              RS!ImpExp_File & " re-transmitted to " & RS!EDI_NatCode & fSep & _
                              "REQUEUE" & fSep & _
                              MS_REQUEUE & fSep & _
                              strDateTime & fSep & _
                              Null & vbCrLf

'               If blnCodeAfterDate Then
'                  sBufSVC.Append strDateTime & fSep & _
'                                 MS_REQUEUE & fSep & _
'                                 Null & vbCrLf
'               Else
'                  sBufSVC.Append MS_REQUEUE & fSep & _
'                                 strDateTime & fSep & _
'                                 Null & vbCrLf
'               End If
               
               If (RS!Header_Status And MS_AWAIT_ACK) = MS_AWAIT_ACK Then
                  mStat = -1
               Else
                  mStat = 0
               End If
               
'              Reset the Msg OK flag and set the Requeue flag. The stored procedure returns the Header_Status
'              with the Ack_Received, Ack failure and DTS error flags reset to 0
               hdrSQL = hdrSQL & _
                        "UPDATE Service_ImpExp_Headers SET " & _
                           "Header_Status = " & (RS!Header_Status Or MS_MSGOK Or MS_REQUEUE) & ", " & _
                           "Status_Flag = " & mStat & _
                        " WHERE Service_ImpExp_Id = " & rq.Identifier & "; " & vbCrLf
            End If
            
            If frmShowRequeue.TextWidth(Trim(.TextMatrix(i, 2))) > max(2) Then
               max(2) = frmShowRequeue.TextWidth(Trim(.TextMatrix(i - 1, 2)))
            End If
            
         End With
         
         RS.Close
      Next i
      
      If Not pBar Is Nothing Then
         pBar.value = pBar.max
      End If
      
      With frmShowRequeue
         .txtRequeued = EDIRec
         .txtDuplicate = DupeTot
         .txtNote = noteTot
         .lblNote = "No Source file"
         .Resize max(0), max(1), max(2)
         .Show 1
      End With
      
      If frmShowRequeue.FormAction = "Commit" Then
         parentForm.MousePointer = vbHourglass
         
         With frmShowRequeue.fgReq
            If fileDest = "" Then
               totSelected = 0
               For i = 1 To .rows - 1
                  If .TextMatrix(i, 2) = "Ready to Copy" Then
                     fs.CopyFile .TextMatrix(i, 0), fs.BuildPath(svcPath(.Tag), _
                                 fs.GetFileName(.TextMatrix(i, 0)))
                     totSelected = totSelected + 1
                  End If
               Next i
            Else
               If fs.FileExists(fileDest) Then
                  fs.DeleteFile (fileDest)
               End If
               
               totSelected = 1
               
               fs.CopyFile .TextMatrix(1, 0), fileDest
               fileDest = ""
            End If
         End With
         
         If totSelected > 0 Then
            If fs.FileExists(rqSVCFile) Then
               fs.DeleteFile rqSVCFile
            End If
            
            Open rqSVCFile For Binary As #1
            Put #1, , sBufSVC.ActualValue
            Close #1
            
            Set dbBcp = Nothing
            Set dbBcp = CreateObject("SQLDMO.BulkCopy")
            
            Set SvcComments = sqlDb.Tables("Service_ImpExp_Comments")
            
            With dbBcp
               .MaximumErrorsBeforeAbort = totSelected
               .DataFilePath = rqSVCFile
               .DataFileType = SQLDMODataFile_SpecialDelimitedChar
               .ColumnDelimiter = fSep
               .RowDelimiter = vbCrLf
               .ErrorFilePath = rqERRFile
               .UseExistingConnection = True
            End With
            
            SVCRec = SvcComments.ImportData(dbBcp)
            iceCon.Execute hdrSQL
         End If
      End If
   End If
   
   Unload frmShowRequeue
   parentForm.MousePointer = vbNormal
   
   Set colRequeue = Nothing
   Set colKey = Nothing
   Set pBar = Nothing
   Exit Sub
   
procEH:
   If Err.Number = 457 Then
      lKey = "Dupe_" & RS!Service_Report_Index & "_" & RS!Service_ImpExp_Message_ID
      Resume
   End If
   
   If blnInTrans Then
      iceCon.RollbackTrans
   End If
   
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
'   If Not iceCmd Is Nothing Then
'      MsgBox Err.Number & "; " & Err.Description
'      rVal = iceCmd.Parameters("Return").value
'      Resume Next
'   End If
   eClass.CurrentProcedure = "RequeueControl.RequeueData"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Public Sub RequeueFile(FileId As String, _
                       FileType As Integer)
   Dim Tstr As String
   Dim strDest As String
   Dim RS As New ADODB.Recordset
   Dim fileBuf As String
   Dim buf As String
   Dim strHdr As String
   Dim strEncrypt As String
   Dim strDecrypt As String
   Dim mPos As Long
   Dim pos As Long
   Dim blnEncrypt As Boolean
   Dim blnAck As Boolean
   Dim blnHTML As Boolean
   Dim sBuf As New StringBuffer
   
   frmShowRequeue.fgReq.Tag = FileType
   If FileType = 1 Then
      Tstr = fs.GetBaseName(FileId)
      Tstr = Tstr & "." & fs.GetExtensionName(FileId)
      
      strSQL = "SELECT Connection_CollectHow " & _
               "FROM Connections " & _
                  "INNER JOIN Connect_Modules ON " & _
                  "Connections.Connection_Name = Connect_Modules.Connection_Name " & _
               "WHERE Module_Name = 'ICEIMP.EXE" & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      fileDest = fs.BuildPath(fs.GetParentFolderName(RS!Connection_CollectHow), Tstr)
      RS.Close
      
   ElseIf FileType = 2 Then
'     copy file (get target directory from ICEMAPI connection)
      strSQL = "SELECT Connection_TargetDirectory " & _
               "FROM Connections C " & _
                  "INNER JOIN Connect_Modules CM " & _
                  "ON C.Connection_Name = CM.Connection_Name " & _
               "WHERE Module_Name = 'ICEMSG.EXE' " & _
                  "AND Connection_InFlightMapping = 'EDIRECIPLIST'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      fileDest = Trim(RS!Connection_TargetDirectory) & "\" & fs.GetFileName(FileId)
      RS.Close
    
'      SourceFile(strDest) = buf
      
'      ElseIf UCase(fs.GetExtensionName(FileId)) = "HTML" Then
'         blnHTML = True
'         strSQL = "HTML file resent "

'      Else
'         fileDest = strDest & "." & fs.GetExtensionName(FileId)
'         strSQL = "File requeued to be RESENT as " & strDest
'      End If
   End If
   
   RequeueData
   Unload frmShowRequeue
End Sub

Public Sub xRequeueFile(FileId As String, _
                       FileType As Integer)
   Dim Tstr As String
   Dim strDest As String
   Dim RS As New ADODB.Recordset
   Dim fileBuf As String
   Dim buf As String
   Dim strHdr As String
   Dim strEncrypt As String
   Dim strDecrypt As String
   Dim mPos As Long
   Dim pos As Long
   Dim blnEncrypt As Boolean
   Dim blnAck As Boolean
   Dim blnHTML As Boolean
   Dim sBuf As New StringBuffer
   
   frmShowRequeue.fgReq.Tag = FileType
   If FileType = 1 Then
      Tstr = fs.GetBaseName(FileId)
      Tstr = Tstr & "." & fs.GetExtensionName(FileId)
      
      strSQL = "SELECT Connection_CollectHow " & _
               "FROM Connections " & _
                  "INNER JOIN Connect_Modules ON " & _
                  "Connections.Connection_Name = Connect_Modules.Connection_Name " & _
               "WHERE Module_Name = 'ICEIMP.EXE" & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      strDest = fs.BuildPath(fs.GetParentFolderName(RS!Connection_CollectHow), Tstr)
      RS.Close
      
   ElseIf FileType = 2 Then
'     copy file (get target directory from ICEMAPI connection)
      strSQL = "SELECT Connection_TargetDirectory " & _
               "FROM Connections C " & _
                  "INNER JOIN Connect_Modules CM " & _
                  "ON C.Connection_Name = CM.Connection_Name " & _
               "WHERE Module_Name = 'ICEMSG.EXE' " & _
                  "AND Connection_InFlightMapping = 'EDIRECIPLIST'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      strDest = Trim(RS!Connection_TargetDirectory) & "\" & fs.GetBaseName(FileId)
      RS.Close
    
'     Use binary access to avoid crlf on end of file
      
      fileBuf = String(8192, Chr(0))
      Open FileId For Binary As #1
      Do Until EOF(1)
         Get #1, , fileBuf
         sBuf.Append fileBuf
      Loop
      buf = Mid(sBuf.ActualValue, 1, LOF(1))
      Close #1
      strHdr = ""
      
      If InStr(1, buf, "NHS003") > 0 Then
'        Locate the end of UNB
         mPos = InStr(buf, "'UNH")
         pos = InStr(1, buf, "MEDRPT") + 6
         strHdr = Left(buf, pos - 1)
         blnEncrypt = False
         
         If mPos = pos Then
'           File is Edifact 3 - not encrypted & no acks required
            strEncrypt = "+++P1"
            strDecrypt = ""
            
         Else
            If Mid(buf, pos, 4) = "++1+" Then
               blnAck = True
'              Acknowledgement flag set
               Select Case Mid(buf, pos + 3, 4)
                  Case "+P1+"
'                    With Encrypt + Test flags
                     blnEncrypt = True
                     strEncrypt = "++1+P1+1"
                     strDecrypt = "++1++1"
                     
                  Case "+P1'"
'                    No Test Flag
                     blnEncrypt = True
                     strEncrypt = "++1+P1"
                     strDecrypt = "++1"
                  
                  Case "++1'"
'                    Test flag only
                     strEncrypt = "++1+P1+1"
                     strDecrypt = "++1++1"
                     
                  Case Else
'                    Acknowledgement only
                     strEncrypt = "++1+P1"
                     strDecrypt = "++1"
                     
               End Select
               
            ElseIf Mid(buf, pos, 3) = "+++" Then
'              No Acknowledgement
               Select Case Mid(buf, pos + 3, 3)
                  Case "P1'"
'                    Encrypt flag
                     blnEncrypt = True
                     strEncrypt = "+++P1"
                     strDecrypt = ""
                     
                  Case "P1+"
'                    Encrypt + test flags
                     blnEncrypt = True
                     strEncrypt = "+++P1+1"
                     strDecrypt = "++++1"
                  
                  Case Else
'                    Test flag only
                     strEncrypt = "+++P1+1"
                     strDecrypt = "++++1"
                     
               End Select
            End If
         End If
         
         buf = Mid(buf, mPos)
         
         If blnEncrypt Then
'            strSQL = "Encrypted File "
            blnEncrypt = MsgBox("This file was originally encrypted - Send encrypted again?", _
                                vbYesNo, "Encrypted File") = vbYes
         Else
'            strSQL = "Unencrypted File "
            blnEncrypt = MsgBox("This file was originally sent Un-encrypted - Do you wish to encrypt it?", _
                               vbYesNo, "Plain Text File") = vbYes
         End If
         
'         strSQL = strSQL & "requeued to be RESENT "
         
         If blnEncrypt Then
'           Add the encryption char.
            buf = strHdr & strEncrypt & buf
            strDest = strDest & ".XEN"
'            strSQL = strSQL & "(Encrypted) as " & strDest
         
         Else
            buf = strHdr & strDecrypt & buf
            If UCase(fs.GetExtensionName(FileId)) = "XEN" Then
               strDest = strDest & ".XMS"
            Else
               strDest = strDest & "." & fs.GetExtensionName(FileId)
            End If
            
'            strSQL = strSQL & "(Unencrypted) as " & strDest
         End If
         SourceFile(strDest) = buf
      
      ElseIf UCase(fs.GetExtensionName(FileId)) = "HTML" Then
         blnHTML = True
'         strSQL = "HTML file resent "

      Else
         strDest = strDest & "." & fs.GetExtensionName(FileId)
'         strSQL = "File requeued to be RESENT as " & strDest
      End If
   End If
   
   RequeueData
   Unload frmShowRequeue
End Sub

Public Property Let RequeueReport(blnNewValue As Boolean)
   reportNotFile = blnNewValue
End Property

Public Property Let SourceFile(Filename As String, ByRef FileContent As String)
   fileSrc = FileContent
   fileDest = Filename
End Property

Public Sub UseProgressBar(ProgBar As ProgressBar, _
                          IncrementBy As Long)
   Set pBar = ProgBar
   stepValue = IncrementBy
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   sqlDb.Tables.Remove "cfgTMP_EDI_Rep_List"
   sqlDb.Tables.Remove "cfgTMP_Service_ImpExp_Comments"
End Sub


