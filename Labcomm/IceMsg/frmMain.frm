VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ICE ... Msg"
   ClientHeight    =   2505
   ClientLeft      =   6180
   ClientTop       =   4170
   ClientWidth     =   6330
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6330
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      Picture         =   "frmMain.frx":0882
      ScaleHeight     =   2535
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Line Line2 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         X1              =   3600
         X2              =   6200
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         X1              =   3600
         X2              =   6200
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblNatCode 
         BackStyle       =   0  'Transparent
         Caption         =   "NatCode"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblMsgFormat 
         BackStyle       =   0  'Transparent
         Caption         =   "MsgFormat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © Sunquest 1998-2009"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Versionlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: "
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "IceMsg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Filename"
         ForeColor       =   &H00FF8080&
         Height          =   360
         Left            =   3600
         TabIndex        =   2
         Top             =   840
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private curprocess As String
Private recpAccount As String
Private orgId As String
Private natCode As String
Private MsgFormat As String
Private lastInterchange As Long
Private EDIRepId As Long
Private runStart As Date

Private colRT As New Collection

'*********************************************************************
' Form Header
'
'
'  Identification:   ICEMsg - Edifact & ASTM Message generation
'  Copyright (c) 2000-02 Anglia Healthcare Systems Ltd
'
'  Author:  Bernie Perrett, Anglia Healthcare Systems Ltd
'
'
'*********************************************************************

Private Sub Form_Load()
 
'  Check_Single_Application Me
   'Timer1.Enabled = False
   UDLPath = App.Path & "\"
'   Ice_MakeConnections
   Versionlbl.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   'Command2.Caption = "Data source: " & iceCon.DefaultDatabase
   'Command2.Caption = iceCon.DefaultDatabase & "    (" & DBVERSION & ")"
   lblDB.Caption = iceCon.DefaultDatabase & " version " & DBVERSION
   CreateMessenger Me, "127.0.0.1", 9000
   frmMain.Visible = True
   frmMain.Show
   frmMain.Caption = "ICE...Log On"
   'Versionlbl.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
'   Timer1.Interval = 60000
   frmMain.Visible = True
   frmMain.Show
   
End Sub

Public Sub NewFileRequired()
   curprocess = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   DestroyMessenger Me
End Sub

Public Sub Read_Reports()
   On Local Error GoTo procEH
   Dim RSFile As String
   Dim iceCmd As New ADODB.Command
   Dim blnEMIS As Boolean
   Dim HtmlFile As String
   Dim varTemp As Variant
   Dim RepId As String
   Dim natCode As String
   Dim MsgFormat As String
   Dim RepSpec As String
   Dim lTrade As String
   Dim blnSMTP As Boolean
   Dim blnSingleReportPerMsg As Boolean
   Dim strSMTP As String
   Dim blnProcess As String
   Dim blnIgnoreHidden As Boolean
   Dim practiceSMTP As String
   Dim RLData As RepData
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   Dim RS4 As New ADODB.Recordset   '  Used for recovery after error

   frmMain.Caption = "Preparing Reports"
   DoEvents
   eClass.FurtherInfo = "Selecting Nat Codes from EDI_Rep_List"
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_ReadReplist"
      Set repListRS = .Execute
   End With
   
   RSFile = fs.BuildPath(App.Path, "RepListRS.ADTG")
   If fs.FileExists(RSFile) Then
      fs.DeleteFile RSFile
   End If
   
   Set repListRS.ActiveConnection = Nothing
   repListRS.Save RSFile, adPersistADTG
   repListRS.Close
   Set iceCmd = Nothing
   repListRS.Open RSFile
   
   blnIgnoreHidden = False
   If Val(Read_Ini_Var("OUTPUT", "IgnoreHidden", INIFile)) = 1 Then
      blnIgnoreHidden = True
   End If
   
   eClass.FurtherInfo = "Obtaining path and polling details for this connection"
   varTemp = Get_Connection_Details(orgId, "EDIRECIPLIST")
   OutputPath = Trim(varTemp(0))
   ErrorPath = Trim(varTemp(1))
   HistoryPath = Trim(varTemp(3))
   PendingPath = Trim(varTemp(4))
   curprocess = ""

   If blnIgnoreHidden Then
      repListRS.Filter = "Category <> 'HIDE'"
   End If
   
   SendMessage "IceMsg", "Messaging Commencing - " & repListRS.RecordCount & " reports to process"
   eClass.LogMessage "Messaging Commencing - " & repListRS.RecordCount & " reports to process"
   
   Do Until repListRS.EOF
      DoEvents
      eClass.FurtherInfo = "Processing report: " & repListRS!Service_Id & " For Nat Code: " & repListRS!EDI_Loc_Nat_Code_To & _
                           ", Specialty/Message Format: " & repListRS!service_Report_Type & "/" & Trim(repListRS!EDI_Msg_Format & "")
      
      RepId = repListRS!EDI_Report_Index
      natCode = repListRS!EDI_Loc_Nat_Code_To
      MsgFormat = Trim(repListRS!EDI_Msg_Format & "")
      RepSpec = repListRS!service_Report_Type
      lTrade = repListRS!EDI_LTS_Index
      
      If repListRS.Fields(repListRS.Fields.Count - 1).Name = "EDI_Msg_IO" Then
         blnSingleReportPerMsg = (repListRS!EDI_Msg_IO = "S")
      Else
         blnSingleReportPerMsg = False
      End If
'      blnProduceXML = (blnProduceXML And (UCase(EDI_GX_Code) = "XML"))
      
      If Trim(repListRS!EDI_SMTP_Mail & "") = "" Then
         practiceSMTP = X400Address(natCode)
      Else
         practiceSMTP = repListRS!EDI_SMTP_Mail
      End If
      
      blnProcess = True
      Select Case repListRS!EDI_Output_GroupBy
         Case "R"
'            repListRS!EDI_Run_Frequency = ""
            If Not (msgControl Is Nothing) Then
               msgControl.WriteFile
            End If
            Set msgControl = Nothing
            Set msgControl = New IceMsgControl
            msgControl.NewFile natCode, repListRS!EDI_Msg_Format, RepId
            
         Case "S"
            If blnSingleReportPerMsg Then
               curprocess = ""
               
               If Not (msgControl Is Nothing) Then
                  msgControl.WriteFile
               End If
               
               Set msgControl = Nothing
               Set msgControl = New IceMsgControl
               msgControl.NewFile natCode, repListRS!EDI_Msg_Format, RepId
            Else
               If natCode & MsgFormat & RepSpec & lTrade <> curprocess Then
                  If Not (msgControl Is Nothing) Then
                     msgControl.WriteFile
                  End If
                  Set msgControl = Nothing
                  Set msgControl = New IceMsgControl
                  msgControl.NewFile natCode, repListRS!EDI_Msg_Format, RepSpec
                  curprocess = natCode & MsgFormat & RepSpec & lTrade
               End If
            End If

         Case Else
            If blnSingleReportPerMsg Then
               curprocess = ""
               
               If Not (msgControl Is Nothing) Then
                  msgControl.WriteFile
               End If
               
               Set msgControl = Nothing
               Set msgControl = New IceMsgControl
               msgControl.NewFile natCode, repListRS!EDI_Msg_Format, RepId
            Else
               If natCode & MsgFormat & lTrade <> curprocess Then
                  If Not (msgControl Is Nothing) Then
                     msgControl.WriteFile
                  End If
                  Set msgControl = Nothing
                  Set msgControl = New IceMsgControl
                  msgControl.NewFile natCode, Trim(repListRS!EDI_Msg_Format & "")
                  curprocess = natCode & MsgFormat & lTrade
               End If
            End If
      End Select
      
      msgControl.FormatIndex = repListRS!FormatIndex
      
      If repListRS!EDI_Output_GroupBy <> "R" Then
         blnProcess = NextRunTime(natCode, RepSpec, MsgFormat, CInt(lTrade))
      Else
         blnProcess = True
      End If
           
'      blnProcess = True
      lblNatCode.Caption = "Processing: " & Trim(repListRS!EDI_Loc_Nat_Code_To)
      lblMsgFormat.Caption = MsgFormat
      Me.Caption = Trim(Left(repListRS!Service_Id, 16))
      
      If blnProcess Then
         msgControl.SetReport repListRS!EDI_Report_Index, repListRS!service_Report_Type, repListRS!Service_Id, , , repListRS!Extra_Copy_To, repListRS!EDI_LTS_Index
                     
'        Update the last run times
'        RS3!EDI_Last_Run = lastRD
'        RS3.Update
                     
         frmMain.Refresh
         
         SendMessage "IceMsg", "Processing: " & repListRS!Service_Id
         eClass.LogMessage "Processing: " & repListRS!Service_Id
         msgControl.CreateMessage repListRS!EDI_Report_Index
      Else
         'Me.Caption = "Run Time Restriction"
         SendMessage "IcgMsg", repListRS!Service_Id & ": " & natCode & " (Specialty " & RepSpec & ") - run-time restriction imposed"
         eClass.LogMessage repListRS!Service_Id & ": " & natCode & " (Specialty " & RepSpec & ") - run-time restriction imposed"
      End If
      
AfterError:
      repListRS.MoveNext
      
      If Not msgControl Is Nothing Then
         If msgControl.MsgInBatch = maxForPractice Then
            curprocess = "Writefile"
         End If
      End If
      DoEvents
   Loop
   
   If msgControl Is Nothing Then
      SendMessage "IceMsg", "No reports to process"
   Else
      msgControl.WriteFile
      '  Now update the Specialty Run Times. This needs to be done outside main loop because the if the
      '  max files per report is reached, the last run time field is updated and subsequent reports are
      '  incorrectly held.
      UpdateRunTimes
      'msgControl.UpdateRunTimes
   End If
   
   iceErr.WriteFile
   
   If blnIgnoreHidden Then
      repListRS.Filter = "Category = 'HIDE'"
      SendMessage "IceMsg", "Deleting " & repListRS.RecordCount & " hidden reports..."
      strSQL = ""
      
      Do Until repListRS.EOF
         strSQL = strSQL & _
                  "DELETE FROM EDI_Rep_List " & _
                  "WHERE EDI_Report_Index = " & repListRS!EDI_Report_Index & "; "
         repListRS.MoveNext
      Loop
      
      If strSQL <> "" Then
         iceCon.BeginTrans
         iceCon.Execute strSQL
         If blnRollBack Then
            iceCon.RollbackTrans
         Else
            iceCon.CommitTrans
         End If
      End If
   End If
   
   repListRS.Close
   Set repListRS = Nothing
   
   Set msgControl = Nothing
   Set iceErr = Nothing
   Set iceCmd = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
   If Err.Number = 3263 Then
      SendMessage "IceMsg", "Database inconsistencies found"
      Set RLData = New RepData
      RLData.EDIIndex = RepId
      RLData.Discipline = Val(repListRS!service_Report_Type)
      RLData.OrStatus MS_DATA_INTEGRITY
      RLData.MessageFormat = MsgFormat
      RLData.LogData(MS_DATA_INTEGRITY) = eClass.LogError
      
      If orgCode = "" Then
         orgCode = repListRS!EDI_Provider_Org
      End If
      
      iceErr.Add ">>> Database inconsistencies reading rep list - see tracking for further information" & vbCrLf & _
                 "Report Index: " & RepId, RLData
      Resume AfterError
   End If
   
   If Val(RepId) = 0 Then
      eClass.LogMessage "Report id not available." & vbCrLf & "Last SQL Statement: " & vbCrLf & strSQL
   End If
   
   eClass.CurrentProcedure = "frmMain.Read_Reports"
   eClass.FurtherInfo = eClass.FurtherInfo & " (Report Index = " & RepId & ")"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Process_Acknowledgements()
   On Error GoTo procEH
   Dim varTemp As Variant
   Dim pendDir As String
   Dim logDir As String
   Dim errDir As String
   Dim penDir As String
   Dim outDir As String
   Dim inDir As String
   Dim inMask As String
   Dim fileDir As String
   Dim fs As New FileSystemObject
   Dim fl As File
   Dim curFile As String
   Dim errFile As String
   Dim outFile As String
   Dim ackDir As String
   Dim ackErr As String
   Dim mainAck As String
   Dim I As Integer
   Dim fType As Integer
   Dim retryCount As Integer
   
   frmMain.Caption = "Processing Acknowledgements"
   SendMessage "IceMsg", "Processing Acknowledgements"
   
   varTemp = Get_Connection_Details(orgId, "ACKS")
   outDir = varTemp(0)
   errDir = varTemp(1)
   pendDir = varTemp(4)
   inMask = varTemp(5)
   ackDir = varTemp(3)
   fileDir = fs.GetParentFolderName(varTemp(5))
   inDir = ValidateFilepath(fs.BuildPath(fs.GetParentFolderName(fileDir), "Rejected"))
   
   curFile = Dir(inMask)
   
   runStart = Now()
   
   Do Until curFile = ""
      frmMain.Caption = curFile
      eClass.FurtherInfo = "Processing " & curFile
      eClass.LogMessage "Processing: " & curFile
      curFile = fs.BuildPath(fileDir, curFile)
      
      SendMessage App.EXEName, "Processing: " & curFile
      
        Set msgControl = New IceMsgControl
'     Read and decode the acknowledgement file
      fType = msgControl.ReadAck(curFile)
      
      Select Case fType
         Case 1
            '  NHSACK file - Determine the status of the ack
            If (msgControl.MessageStatus And MS_ACK_FAIL) = MS_ACK_FAIL Then
   '           Failed to parse the acknowledgement file
               errFile = fs.BuildPath(errDir, fs.GetFileName(curFile))
               outFile = fs.BuildPath(outDir, "ParseFail_" & fs.GetBaseName(curFile) & ".wng")
               fs.CopyFile curFile, errFile
               eClass.LogMessage "Unable to parse file: " & curFile & " Copied to: " & errFile
               
               msgControl.SetReport 0, "HDR", "Interchange", 0, True
               msgControl.ErrorReport = errFile
               msgControl.LogReportMessage MS_ACK_FAIL, "Acknowledgement file has an invalid format - original file " & _
                                           "saved as " & errFile
               SendMessage App.EXEName, "Acknowledgment parsing error (" & curFile & ")"
            End If
            
            msgControl.WriteAcks
         
         Case 0
            '  Not an Ack file - copy to input rejects
            outFile = fs.BuildPath(inDir, fs.GetFileName(curFile))
            eClass.FurtherInfo = "Copy invalid file (" & curFile & ") to " & outFile
            fs.CopyFile curFile, outFile
            eClass.LogMessage curFile & ": not a valid Ack file - copied to: " & outFile
         
         Case -1
            '  This is an NHS002/3 file - copy to 'pending' directory to await further processing
            outFile = fs.BuildPath(pendDir, fs.GetFileName(curFile))
            eClass.FurtherInfo = "Copy Edifact MEDRPT file (" & curFile & ") to " & outFile
            fs.CopyFile curFile, outFile
            eClass.LogMessage curFile & ": Edifact MEDRPT file - copied to: " & outFile
         
      End Select
      
'      If fType = 1 Then
''        Determine the status of the ack
'         If (msgControl.MessageStatus And MS_ACK_FAIL) = MS_ACK_FAIL Then
''           Failed to parse the acknowledgement file
'            errFile = fs.BuildPath(errDir, fs.GetFileName(curFile))
'            outFile = fs.BuildPath(outDir, "ParseFail_" & fs.GetBaseName(curFile) & ".wng")
'            fs.CopyFile curFile, errFile
'            eClass.LogMessage "Unable to parse file: " & curFile & " Copied to: " & errFile
'
'            msgControl.SetReport 0, "HDR", "Interchange", 0, True
'            msgControl.ErrorReport = errFile
'            msgControl.LogReportMessage MS_ACK_FAIL, "Acknowledgement file has an invalid format - original file " & _
'                                        "saved as " & errFile
'            SendMessage App.EXEName, "Acknowledgment parsing error (" & curFile & ")"
'         End If
'
'         msgControl.WriteAcks
'
'
'      Else
''        Not an Ack file - copy to input rejects
'         outFile = fs.BuildPath(inDir, fs.GetFileName(curFile))
'         eClass.FurtherInfo = "Copy invalid file (" & curFile & ") to " & outFile
'         fs.CopyFile curFile, outFile
'         eClass.LogMessage curFile & ": not a valid Ack file - copied to: " & outFile
'      End If

'     Delete the ack file
      eClass.FurtherInfo = "Delete Ack file " & curFile
      If (msgControl.MessageStatus And &H10600) > 0 Then
         mainAck = fs.BuildPath(errDir, msgControl.InterchangeErrorName)
         ackErr = mainAck
         Do Until fs.FileExists(ackErr & ".XMR") = False
            I = I + 1
            ackErr = mainAck & "_" & I
         Loop
         
         ackErr = ackErr + ".XMR"
         fs.MoveFile curFile, ackErr
      Else
         If blnSaveAckFiles Or blnRollBack Then
            fs.MoveFile curFile, fs.BuildPath(ackDir, fs.GetFileName(curFile))
         Else
            fs.DeleteFile curFile
         End If
      End If
      
AbortPoint:
      curFile = Dir
   Loop
   
   Dim docManq As String
   docManq = Read_Ini_Var("LETTER", "DocManQueue", INIFile)
    
   If docManq <> "" Then
      eClass.LogMessage "Processing Docman acknowledgement queue"
      SendMessage "IceMsg", "Processing Docman Queue"
      
      Dim dmQueue As New clsMSMQ
      Dim docManData As String
      Dim msg() As String
      Dim rVal As Integer
      
      dmQueue.QueueId = docManq
      
      On Error GoTo QueueReadFail
      docManData = dmQueue.ReadQueue
      
      Do Until docManData = ""
         msg = Split(docManData, "|")
         
         If msg(0) <> "Error" Then
            rVal = msgControl.UpdateDocmanStatus(msg(1), msg(2), (msg(0) <> "ICEMSG"))
            Select Case rVal
               Case -1
                  eClass.LogMessage "Docman Queue: Read aborted after database update fail"
                  SendMessage "IceMsg", "Docman Queue: Read aborted after database update fail"
                  
                  dmQueue.AbortRead
                  dmQueue.IncrementFail
               
               Case 0
                  dmQueue.CommitRead
                  dmQueue.IncrementFail True
                  
                  eClass.LogMessage "Docman report " & msg(1) & " accepted by Hub with reference " & msg(2)
                  SendMessage "IceMsg", "Docman report " & msg(1) & " accepted with reference " & msg(2)
               
               Case 1
                  eClass.LogMessage "Docman queue contains an invalid index (" & msg(1) & ")"
                  SendMessage "IceMsg", "Docman queue contains an invalid index (" & msg(1) & ")"
                  
                  '  Clear from queue after logging
                  dmQueue.CommitRead
                  dmQueue.IncrementFail True
               
            End Select
            
            docManData = dmQueue.ReadQueue
         Else
            SendMessage "IceMsg", "Error reading Docman Queue (" & docManq & ": " & msg(1)
            eClass.LogMessage "Error reading Docman Queue (" & docManq & ": " & msg(1)
            docManData = ""
         End If
      Loop
      
   GoTo RecordComplete
      
QueueReadFail: '  We do not want to keep retrying after a failure
      SendMessage "IceMsg", "Failed to read from " & docManq & " - " & Err.Description
   End If
   
RecordComplete:
   SendMessage "IceMsg", "Acknowledgements Completed"
   Set msgControl = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
'  Try to access the file 3 times. If this fails, leave it to be processed next time
   If retryCount < 3 Then
      SendMessage "IceMsg", "Unable to access " & curFile & " - retrying"
      eClass.LogMessage "Accessing " & curFile & " Permission denied - retrying"
      retryCount = retryCount + 1
      Resume
   Else
      eClass.LogMessage "Unable to access " & curFile & " (" & Err.Number & ": " & Err.Description & ")"
      SendMessage "IceMsg", "Unable to access " & curFile & " - aborting"
      retryCount = 0
      Resume AbortPoint
   End If
   
   eClass.CurrentProcedure = "IceMsg.frmMain.Process_Acknowledgements"
   eClass.Add Err.Number, Err.Description, Err.Source, False
   
   If tLevel > 0 Then
      iceCon.RollbackTrans
      tLevel = 0
   End If
End Sub

'Private Sub CreateAckFile()
'   Dim pMep As PMEPAck
'
'   pMep.batchStart.recId = "00"
'   pMep.batchStart.StartSignal = "PMEPACK"
'   pMep.fileSuccess.recId = "80"
'   pMep.fileFailure.recId = "85"
'   pMep.batchEnd.recId = "90"
'   pMep.batchEnd.eobSignal = "ENDPMEPACK"
'
'   pMep.batchStart.mWare_Ver = ""
'   pMep.batchStart.mWare_Code = ""
'   pMep.batchStart.TelNo = ""
'   pMep.batchStart.transNo = ""
'   pMep.batchStart.dateTime = Now()
'
' '  pmep.fileSuccess.
'
'End Sub

Private Function X400Address(natCode As String) As String
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim iceCmd As New ADODB.Command
   Dim strTemp As String
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_ReadX400"
      .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 6, natCode)
      Set RS = .Execute
   End With
      
   If Trim(RS!EDI_X400_GivenName & "") <> "" Then
      strTemp = "g=" & Trim(RS!EDI_X400_GivenName) & ";"
   End If
   If Trim(RS!EDI_X400_Surname & "") <> "" Then
      strTemp = "s=" & Trim(RS!EDI_X400_Surname) & ";"
   End If
   If Trim(RS!EDI_X400_Initials & "") <> "" Then
      strTemp = "i=" & Trim(RS!EDI_X400_Initials) & ";"
   End If
   If Trim(RS!EDI_X400_Generation & "") <> "" Then
      strTemp = "q=" & Trim(RS!EDI_X400_Generation) & ";"
   End If
   If Trim(RS!EDI_X400_Common & "") <> "" Then
      strTemp = "cn=" & Trim(RS!EDI_X400_Common) & ";"
   End If
   If Trim(RS!EDI_X400_org & "") <> "" Then
      strTemp = "o=" & Trim(RS!EDI_X400_org) & ";"
   End If
   If Trim(RS!EDI_X400_OU1 & "") <> "" Then
      strTemp = "ou1=" & Trim(RS!EDI_X400_OU1) & ";"
   End If
   If Trim(RS!EDI_X400_OU2 & "") <> "" Then
      strTemp = "ou2=" & Trim(RS!EDI_X400_OU2) & ";"
   End If
   If Trim(RS!EDI_X400_OU3 & "") <> "" Then
      strTemp = "ou3=" & Trim(RS!EDI_X400_OU3) & ";"
   End If
   If Trim(RS!EDI_X400_OU4 & "") <> "" Then
      strTemp = "ou4=" & Trim(RS!EDI_X400_OU4) & ";"
   End If
   If Trim(RS!EDI_X400_prd & "") <> "" Then
      strTemp = "p=" & Trim(RS!EDI_X400_prd) & ";"
   End If
   If Trim(RS!EDI_X400_Adm & "") <> "" Then
      strTemp = "a=" & Trim(RS!EDI_X400_Adm) & ";"
   End If
   If Trim(RS!EDI_X400_c & "") <> "" Then
      strTemp = "c=" & Trim(RS!EDI_X400_c) & ";"
   End If
   
   X400Address = "[" & strTemp & "]"
   RS.Close
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmMain.X400Address"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function NextRunTime(NationalCode As String, _
                             Specialty As String, _
                             msg As String, _
                             LTSIndex As Integer) As Boolean
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim elapsedTime As Date
   Dim sTime() As Date
   Dim runFreq As String
   Dim lastRD As Date
   Dim runNext As Date
   Dim blnProcess As Boolean
   Dim I As Integer
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_RunTimes"
      .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 6, NationalCode)
      .Parameters.Append .CreateParameter("Spec", adVarChar, adParamInput, 6, Specialty)
      .Parameters.Append .CreateParameter("Msg", adVarChar, adParamInput, 20, msg)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , LTSIndex)
      .Parameters.Append .CreateParameter("Freq", adInteger, adParamOutput)
      Set RS = .Execute
      runFreq = .Parameters("Freq").Value
   End With
                     
   If RS.EOF Then
      blnProcess = True
   Else
      elapsedTime = Val(Mid(Trim(RS!EDI_Run_Frequency & ""), 2))
      blnProcess = False
      lastRD = 0
      
      Select Case UCase(Left(RS!EDI_Run_Frequency, 1))
         Case "H" ' Hold
            blnProcess = False
            
         Case "N" '  Run every <nn> minutes
            If DateDiff("N", RS!EDI_Last_Run, Now()) > elapsedTime Then
               blnProcess = True
               lastRD = Now()
   '            eClass.LogMessage natCode & ":  " & Msg & " released for specialty " & Specialty, , "Frequency"
            End If
            
         Case "D" '  Daily - Start times need to be checked
            ReDim sTime(1)
            lastRD = Format(Now(), "dd/mm/yyyy")
            
            sTime(0) = CDate(lastRD & Format(RS!EDI_S_Time1, " hh:nn:ss"))
            'sTime(0) = CDate(lastRD & " 08:45:00")
            If Trim(RS!EDI_S_Time2 & "") <> "" Then
               ReDim Preserve sTime(2)
               sTime(1) = CDate(lastRD & Format(RS!EDI_S_Time2, " hh:nn:ss"))
            End If
            
            If Trim(RS!EDI_S_Time3 & "") <> "" Then
               ReDim Preserve sTime(3)
               sTime(2) = CDate(lastRD & Format(RS!EDI_S_Time3, " hh:nn:ss"))
            End If
            
            If Trim(RS!EDI_S_Time4 & "") <> "" Then
               ReDim Preserve sTime(4)
               sTime(3) = CDate(lastRD & Format(RS!EDI_S_Time4, " hh:nn:ss"))
            End If
                           
            runNext = 0
            For I = 0 To UBound(sTime)
               If (RS!EDI_Last_Run & "" < sTime(I)) Then
                  runNext = sTime(I)
                  Exit For
               End If
            Next I
                           
            If runNext <> 0 Then
               If Now() > runNext Then
                  blnProcess = True
                  lastRD = CDate(Format(Now(), "dd/mm/yyyy hh:mm:ss"))
               End If
            End If
   
         Case "W" '  Weekly - Note check of Start_time1
            If Trim(RS!EDI_S_Time1 & "") <> "" Then
               elapsedTime = CDate(lastRD & RS!EDI_S_Time1)
            Else
               elapsedTime = CDate(lastRD & "09:00")
            End If
            
            If Now() > DateAdd("D", 7, elapsedTime) Then
               blnProcess = True
               lastRD = Now()
   '            eClass.LogMessage natCode & "  " & Msg & " released for specialty " & Specialty, , "Weekly"
            End If
            
         Case Else
            blnProcess = True
            
      End Select
   End If
   RS.Close
   
   If blnProcess And lastRD <> 0 Then
      RecordRunTimes NationalCode, Specialty, msg, CStr(lastRD)
      'msgControl.RunTimeUpdate NationalCode, Specialty, Msg, CStr(lastRD)
   End If
   
   NextRunTime = blnProcess
   
   Set RS = Nothing
   Set iceCmd = Nothing
End Function

Public Sub RecordRunTimes(natCode As String, _
                          Specialty As String, _
                          msgType As String, _
                          lastRun As String)
   On Error GoTo procEH
   Dim rtIndex As String
   Dim rtc As clsRunTimes
   
   rtIndex = natCode & ":" & Specialty & ":" & msgType
   Set rtc = colRT(rtIndex)
   Set rtc = Nothing
   Exit Sub
   
procEH:
   If Err.Number = 5 Then
      Set rtc = New clsRunTimes
      With rtc
         .NationalCode = natCode
         .Specialty = Specialty
         .MessageType = msgType
         .LastRunAt = lastRun
      End With
      colRT.Add rtc, rtIndex
      Resume Next
   Else
      eClass.FurtherInfo = "frmMain.RecordRunTimes"
      eClass.Add Err.Number, Err.Description, Err.Source
   End If
End Sub

Public Sub UpdateRunTimes()
   Dim rtc As clsRunTimes

   iceCon.BeginTrans
   
   For Each rtc In colRT
      strSQL = "UPDATE EDI_Loc_Specialties SET " & _
                  "EDI_Last_Run = '" & Format(rtc.LastRunAt, "YYYYMMDD hh:nn:ss") & "' " & _
               "WHERE EDI_Nat_Code = '" & rtc.NationalCode & "' " & _
                  "AND EDI_Korner_Code = '" & rtc.Specialty & "' " & _
                  "AND EDI_Msg_Format = '" & rtc.MessageType & "'"
      iceCon.Execute strSQL
      
      With rtc
         SendMessage "IceMsg", .Specialty & ": Runtime updated to " & .LastRunAt & " for " & .NationalCode & " (" & .MessageType & ")"
         eClass.LogMessage .NationalCode & "  " & .MessageType & " released and run-time updated for specialty " & .Specialty, , "Release"
      End With
   Next
   
   If blnRollBack Then
      iceCon.RollbackTrans
   Else
      iceCon.CommitTrans
   End If
   
   Set rtc = Nothing
   Set colRT = Nothing
End Sub
