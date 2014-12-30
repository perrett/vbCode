Attribute VB_Name = "CommonCode"
Option Explicit

Public UDLPath As String
Public INIFile As String
Public NTS As Boolean
Public iceCon As New ADODB.Connection
Public currentLoc As String
Public distinctRst As ADODB.Recordset
Public timeInterval As Integer

Public blnUseDMO As Boolean
Public sqlServer As SQLDMO.sqlServer
Public sqlDb As SQLDMO.Database
Public UDLServer As Variant
Public UDLDatabase As String
Public dbUser As String
Public dbPass As String
Public Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type
   
Public Type Conformance
   cfTraderCode As String
   cfRunningTotal As Long
   cfTrigger As Long
   cfNatCode As String * 6
   cfReplaceCode As String * 6
End Type

Public Enum enumTestFlags
   TF_ACCEPT_ACK = &H80000100
   TF_IS_COMMENT = &H3
   TF_IS_WARN = &HC
   TF_IS_SUPPRESS = &H1F0
'   TF_VALID_RC = &HAD
   TF_ACK_REJECT = &H300000
   TF_SET_ERROR = &H7FFFFFFF
End Enum

Public repListRS As ADODB.Recordset
Public blnRollBack As Boolean
Public blnProduceXML As Boolean
Public blnUpper As Boolean
Public blnIncludeRepId As Boolean
Public blnUseSampleForDisc As Boolean
Public OutputPath As String
Public ErrorPath As String
Public HistoryPath As String
Public PendingPath As String
Public CopyToDir As String
Public copyRecipient As String
Public blnErrHalt As Boolean

Public blnOverride As Boolean
Public blnGlobalAnonymize As Boolean
Public blnGlobalTest As Boolean
Public GlobalMsgFormat As String
Public XMLLABPrefix As String
Public XMLMRNPrefix As String
Public XMLDefPrefix As String
Public XMLLocalCodeID As String
Public XMLPrelimId As String

Public blnClinNameUseAll As Boolean
Public blnUse906 As Boolean
Public blnReadCodeInfo As Boolean
Public blnRetainGrave As Boolean
Public blnASTMLocal As Boolean
Public blnASTMSameDate As Boolean
Public blnAmendUOM As Boolean

Public blnForceClinLoc As Boolean
Public blnClinInPV1 As Boolean
Public blnNoMSH As Boolean
Public blnSaveAckFiles As Boolean
Public blnUseRCIndex As Boolean
Public blnSeqReqd As Boolean
Public blnUseHTTPS As Boolean
Public orgCode As String
Public ltIndex As Long
Public iceErr As New msgErr

'Public xmlSender As String
'Public xmlSendFacility As String
'Public procID As String
'
Public maxRetries As Integer
Public fileData As Variant
Public tLevel As Long
Public fs As New FileSystemObject
Public msgControl As IceMsgControl
Public LogStatus As Long
Public KeyName As String
Public strSQL As String
Public outDir As String
Public errDir As String
Public MsgFormat As String
Public Specialty As String
Public SvcRepIndex As String
Public useSymphonia As Boolean
Public acksExpected As Integer
Public eClass As AHSLErrorLog_XML.errorControl
Public htmlBuf As String
Public htmlHeader As String
Public htmlBody As String
Public htmlTrailer As String
Public orgId As String
Public Const ERRFILEPATH = "c:\ice\server\errors\"
Public conName As String
Public eTrap As New errorExceptions
'Public ackFile As PMEPAck
'Public txtHandler As New AHSLMessaging.TextSplit
Public msgData As New MsgRoutines
Public maxForPractice As Integer
Public LogLevel As Integer
Public blnForceTestMsg As Boolean
Public ovrMsgType As String
Public ovrMsgQual As String
Public transCount As Integer

Public DBVERSION As Long
Public sDOMAIN As String
Public blnRemoveImports As Boolean
Public defUserIndex As String
Public sDocManSrc As String

Public Sub Main()
   On Error GoTo procEH
   Dim RSFile As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim logStat As Long
   Dim LogLevel As Integer
   Dim logSrc As String
   Dim logPath As String
   Dim eTemp As Integer
   Dim strTemp As String
   Dim disconRS As New ADODB.Recordset
   Dim pCheck As New clsWait
   Dim runExcl As String
   Dim Retries As String
   
   INIFile = App.Path & "\ICEMsg.ini"
   logStat = Val(Read_Ini_Var("Logging", "Mode", INIFile))
   
   Set eClass = CreateObject("AHSLErrorLog_XML.errorControl")
   With eClass
      .ApplicationLogging "IceMsg", logStat
      .FurtherInfo = "Setting exception module"
      .ExceptionModule = eTrap
      .Behaviour = 1  '   Ensure any errors hit before the behaviour has been read are handled.
   End With
   
   DMOAvailable
   
   Set iceCon = frmDatabase.ConnectionDetails()
   With eClass
      .FurtherInfo = "Setting database connection for XML Error log"
      .dbConnection = iceCon
   End With
   
   If iceCon Is Nothing Then
      SendMessage "IceMsg", frmDatabase.ConnectionError
   Else
      RSFile = fs.BuildPath(App.Path, "ConnectRS.ADTG")
      
      If fs.FileExists(RSFile) Then
         fs.DeleteFile RSFile
      End If
      
      'If Ice_CheckVersion Then
         
         If fs.FileExists(RSFile) Then
            disconRS.Open RSFile
         Else
         
            strSQL = "SELECT * " & _
                     "FROM Connections " & _
                        "INNER JOIN Connect_Modules ON " & _
                        "Connections.Connection_Name = Connect_Modules.Connection_Name " & _
                     "WHERE Module_Name = '" & App.EXEName & ".exe'"
            
            disconRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
            Set disconRS.ActiveConnection = Nothing
            disconRS.Save RSFile, adPersistADTG
            disconRS.Close
            disconRS.Open RSFile
         End If
         
         LogLevel = Val(Read_Ini_Var("Logging", "Level", INIFile))
         
         With eClass
            .FurtherInfo = "Setting up Logging"
            .LogDirectory = ValidateFilepath(Trim(disconRS!Connection_LogDirs))
            .FormCaption = "Ice Messaging Error"
            
            .LogLevel = LogLevel
            .LogMessage "IceMsg Version: " & App.Major & "." & App.Minor & " - Process Starting", , "Messaging"
   '         .LogMessage "UDL Path = " & UDLPath
         
         End With
         
         If App.PrevInstance Then
            eClass.LogMessage "IceMsg already running - this run aborted", , "Warning"
         Else
            
            eClass.FurtherInfo = "Setting INI file options"
            blnErrHalt = (Read_Ini_Var("General", "HaltOnError", INIFile) = 1)
            
            blnSaveAckFiles = (Read_Ini_Var("GENERAL", "SaveAcks", INIFile) = 1)
            
            blnUse906 = (Read_Ini_Var("OUTPUT", "EDI2Use902", INIFile) = 1)
            
            blnClinNameUseAll = (Read_Ini_Var("OUTPUT", "ClinicianNameParsed", INIFile) = 1)
            
            blnUpper = (Read_Ini_Var("OUTPUT", "UpperCaseRubric", INIFile) = 1)
            maxForPractice = Val(Read_Ini_Var("OUTPUT", "MaxReportsPerPractice", INIFile))
         
   '        Do we output the local code for ASTM?
            blnASTMLocal = (Read_Ini_Var("OUTPUT", "AstmLocal", INIFile) = 1)
         
   '        Replace ` with ' in the output
            blnRetainGrave = (Read_Ini_Var("OUTPUT", "IgnoreApostrophe", INIFile) = 1)
      
   '        Log "Not read-coded" messages (where lab or result mapping indicates no read code)
            blnReadCodeInfo = (Read_Ini_Var("LOGGING", "SuppressMissingRCinfo", INIFile) <> 1)
   
   '        Override message settings and use ini file values
            GlobalMsgFormat = ""
            blnOverride = (Read_Ini_Var("Testing", "Override", INIFile) = 1)
            If blnOverride Then
               GlobalMsgFormat = Read_Ini_Var("Testing", "MessageType", INIFile)
               blnGlobalTest = (Read_Ini_Var("Testing", "TestMessage", INIFile) = 1)
               blnGlobalAnonymize = (Read_Ini_Var("Testing", "Anonymize", INIFile) = 1)
            End If
            
   '        Output MSH segment in XML
            blnNoMSH = (Read_Ini_Var("OUTPUT", "XML_NoMSH", INIFile) = 1)
      
   '        Remove '^' from UOM?
            blnAmendUOM = (Read_Ini_Var("OUTPUT", "XML_AmendUOM", INIFile) = 1)
            
   '        Read PID.3 Prefix data
            XMLMRNPrefix = Read_Ini_Var("XML", "MRNPrefix", INIFile)
            XMLLABPrefix = Read_Ini_Var("XML", "LABPrefix", INIFile)
            XMLDefPrefix = Read_Ini_Var("XML", "DEFPrefix", INIFile)
            
            XMLPrelimId = Read_Ini_Var("XML", "PrelimId", INIFile)
            
            XMLLocalCodeID = Read_Ini_Var("XML", "LocalCodeId", INIFile)
            If XMLLocalCodeID = "" Then
               XMLLocalCodeID = "LC"
            End If
                     
   '        Report Index in PV1,19 for XML?
            blnIncludeRepId = (Read_Ini_Var("OUTPUT", "XML_PV19ReportId", INIFile) = 1)
      
   '        Clinician in PV1 for XML message?
            blnClinInPV1 = (Read_Ini_Var("OUTPUT", "XML_ClinicianInPV1", INIFile) = 1)
         
   '        Use sample-id prefix to identify discipline
            blnUseSampleForDisc = (Read_Ini_Var("OUTPUT", "XML_SampleForDiscipline", INIFile) = 1)
         
   '        Use Sample Collection Date rather than date/time received for ASTM
            blnASTMSameDate = (Read_Ini_Var("OUTPUT", "ASTM_UseCollectDate", INIFile) = 1)
      
   '        Force the clinician local id to be used in RFF segment
            blnForceClinLoc = (Read_Ini_Var("OUTPUT", "UseClinicianLocalId", INIFile) = 1)
      
   '        Do we rollback the transactions?
            blnRollBack = (Read_Ini_Var("GENERAL", "DBNoUpdate", INIFile) = 1)
         
            If blnRollBack Then
               MsgBox "Rollback option set - This run for testing only", vbInformation, "Please confirm"
               eClass.LogMessage "Rollback set - output WILL NOT update the audit trail", , "Diagnostic"
               SendMessage "IceMsg", "Rollback Set - output WILL NOT update the audit trail"
            End If
            
            CopyToDir = Read_Ini_Var("OUTPUT", "XMLDestination", INIFile)
            If CopyToDir <> "" Then
               ValidateFilepath CopyToDir
            End If
               
            blnForceTestMsg = (Read_Ini_Var("Testing", "TestMessage", INIFile) = 1)
            blnOverride = (Read_Ini_Var("Testing", "Override", INIFile) = 1)
               
            If blnOverride Then
               ovrMsgType = Read_Ini_Var("Testing", "MessageType", INIFile)
               blnOverride = (ovrMsgType <> "")
            End If
               
            Load frmMain
            
            eClass.FurtherInfo = "Setting Current Configuration"
            SendMessage App.ProductName, "Process commencing"
            DoEvents
            
            eTemp = Val(Read_Ini_Var("Logging", "Behaviour", INIFile))
            
            If RunningInIDE = False Then
               eTemp = Abs(eTemp)
            End If
            eClass.Behaviour = eTemp
            
            strTemp = Read_Ini_Var("GENERAL", "TimeOutRetries", INIFile)
            If IsNumeric(strTemp) Then
               maxRetries = Val(strTemp)
            Else
               maxRetries = 3
            End If
            
            runExcl = Read_Ini_Var("General", "RunExclusiveTo", INIFile)
            If runExcl <> "" Then
               Retries = Read_Ini_Var("General", "RunExclRetries", INIFile)
               pCheck.FeedbackObject = frmMain
               If IsNumeric(Retries) Then
                  If Retries > 0 Then
                     pCheck.Attempts = Val(Retries)
                  End If
               End If
            End If
      
            blnUseHTTPS = (Read_Ini_Var("LETTER", "SecureServer", INIFile) = "1")
            
            eClass.FurtherInfo = "Processing connections"
            Do Until disconRS.EOF
               Select Case Trim(disconRS!Connection_InFlightMapping & "")
                  Case "EDIRECIPLIST"
                     If disconRS!Connection_Active Then
                        eClass.LogMessage "Processing EDI_Rep_List"
                        
                        If runExcl <> "" Then
                           If pCheck.waitFor(runExcl) Then
                              frmMain.Read_Reports
                           Else
                              eClass.LogMessage runExcl & " - still running. No reports processed"
                              SendMessage "IceMsg", runExcl & " - still running. No reports processed"
                              End
                           End If
                        Else
                           frmMain.Read_Reports
                        End If
                     Else
                        SendMessage App.EXEName, Trim(disconRS!Connection_InFlightMapping) & " not active"
                     End If
                     
                  Case ""
   '                 Null or nothing so do not process
   
                  Case Else
                     If disconRS!Connection_Active Then
                        eClass.LogMessage "Processing acknowledgements", , "Acks"
                        frmMain.Process_Acknowledgements
                        eClass.LogMessage "Acknowledgements completed", , "Acks"
                     Else
                        eClass.LogMessage Trim(disconRS!Connection_InFlightMapping) & " not active", , "Warning"
                        SendMessage App.EXEName, Trim(disconRS!Connection_InFlightMapping) & " not active"
                     End If
               
               End Select
               disconRS.MoveNext
            Loop
            SendMessage App.ProductName, "Process complete"
            eClass.LogMessage "Process Complete", , "Messaging"
         End If
         
         disconRS.Close
         
         On Error GoTo 0
         On Error Resume Next
         fs.DeleteFile RSFile
      
      'End If
   End If
   
   Set iceCon = Nothing
   Set disconRS = Nothing
   Set msgControl = Nothing
   Set eClass = Nothing
   End

procEH:
   If eClass Is Nothing Then
      SendMessage "IceMsg", "Error: " & Err.Number & " " & Err.Description & " before error log created. " & _
                  "See IceLabcomm log for details"
   Else
      If eClass.Behaviour = -1 Then
         Stop
         Resume
      End If
      eClass.CurrentProcedure = "CommonCode.Main"
      eClass.Add Err.Number, Err.Description, Err.Source, False
   End If
   
   End
End Sub

Public Function Ice_CheckVersion() As Boolean
   On Error GoTo procEH '  LogOnFail
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim timeOut As Long
   Dim dbTime As String
   Dim connectString As String
   Dim UDLFile As String
   Dim timeOutCount As Integer
   
'   Set iceCon = frmDatabase.ConnectionDetails()
   
'   eClass.FurtherInfo = "db version " & App.Minor & " required."
   strSQL = "SELECT Version FROM dbVersion"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.RecordCount > 0 Then
      RS.MoveLast
      If RS!Version < App.Minor Then
         Err.Raise 3041, "Version Check", "Database version " & App.Minor & " required - This version is: " & RS!Version
      Else
         Ice_CheckVersion = True
      End If
   Else
      Ice_CheckVersion = False
   End If
   
   DBVERSION = RS!Version
   
   RS.Close
   Set RS = Nothing
   Exit Function

procEH:
   If Err.Number = -2147467259 Then
      SendMessage "IceMsg", "Timeout on initial connection to database"
      End
   End If
   
'   SendMessage "IceMsg", "CommonCode.Ice_CheckVersion " & Err.Number & " " & Err.Description
   eClass.CurrentProcedure = "CommonCode.Ice_CheckVersion"
   eClass.Add Err.Number, Err.Description, Err.Source
   Ice_CheckVersion = False
End Function

Public Function DMOAvailable() As Boolean
   On Error GoTo procEH
   blnUseDMO = True
   Set sqlServer = CreateObject("SQLDMO.sqlServer")
   blnUseDMO = False
   Set sqlServer = Nothing
   DMOAvailable = blnUseDMO
   Exit Function
   
procEH:
   blnUseDMO = False
   Resume Next
End Function

Public Function Get_Connection_Details(orgId As String, processId As String) As Variant
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim temp(6) As String
   
'  Read Connections to find target directories
   strSQL = "SELECT * " & _
            "FROM Connections " & _
               "INNER JOIN Connect_Modules ON " & _
               "Connections.Connection_Name = Connect_Modules.Connection_Name " & _
            "WHERE "

   If orgId <> "" Then
      strSQL = strSQL & "Organisation = '" & orgId & "' AND "
   End If
   strSQL = strSQL & "Module_Name = 'IceMsg.exe' " & _
            "AND Connection_InFlightMapping = '" & processId & "'"
   
   eClass.FurtherInfo = strSQL
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If (RS.EOF) Then
      MsgBox "Connections table not set up - No In-Flight_Mapping for " & processId, vbExclamation, "Database configuration"
    'need an error check here
   End If
   
   temp(0) = ValidateFilepath(Trim(RS!Connection_TargetDirectory))
   temp(1) = ValidateFilepath(Trim(RS!Connection_ErrorDirs))
   temp(2) = Val(RS!Connection_Frequency)
   temp(3) = ValidateFilepath(Trim(RS!Connection_HistoryDirs))
   temp(4) = ValidateFilepath(Trim(RS!Connection_PendingDirs))
   temp(5) = ValidateFilepath(fs.GetParentFolderName(Trim(RS!Connection_CollectHow))) & "\" & fs.GetFileName(Trim(RS!Connection_CollectHow))
   Get_Connection_Details = temp()
   RS.Close
   Set RS = Nothing
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "CommonCode.Get_Connection_Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function ReadUDL(fileName As String) As String
   Dim pTag As String
   Dim buf As String
   Dim constr As String
   Dim tmpCon As New ADODB.Connection
   Dim dblink As New MSDASC.DataLinks
   Dim pos As Integer
   
   constr = ""
   pTag = "P" & Chr(0) & _
          "r" & Chr(0) & _
          "o" & Chr(0) & _
          "v" & Chr(0) & _
          "i" & Chr(0) & _
          "d" & Chr(0) & _
          "e" & Chr(0) & _
          "r"
   
'   pTag = StrConv("PROVIDER", vbUnicode)
   If fs.FileExists(fileName) Then
      buf = Space(1500)
      Open fileName For Binary As #1
      Get #1, , buf
      Close #1
      pos = InStr(1, buf, pTag, vbBinaryCompare)
      constr = Replace(Mid(buf, pos), Chr(0), "")
      tmpCon.ConnectionString = constr
      dblink.PromptEdit tmpCon
      constr = tmpCon.ConnectionString
   Else
      MsgBox "Specified UDL (" & fileName & ") not found", _
             vbExclamation, "Invalid UDL"
   End If
   ReadUDL = constr
   Set tmpCon = Nothing
   Set dblink = Nothing
End Function

'Public Sub WriteConfirmationFile(fileType As Integer, _
'                                 RecordType As String, _
'                                 LocalTrader As Integer, _
'                                 Optional dateTime As String = "", _
'                                 Optional ReportId As String = "", _
'                                 Optional ReportStatus As String = "", _
'                                 Optional FileLocation As String = "")
'   Dim iSoft As iSoftFile
'   Dim fDate As String
'   Dim nextControlRef As String
'   Static bodyCnt As Integer
'   Dim FILENAME As String
'   Dim strSQL As String
'   Dim RS As New ADODB.Recordset
'   Dim blnWriteFile As Boolean
'   Dim i As Integer
'
''   fDate = Format(Now(), "yyyymmddhhnn")
'   blnWriteFile = False
'   Select Case fileType
'      Case 1
'         Select Case RecordType
'            Case "Header"
'               If iSoft.headerRec.Trigger = "" Then
'                  bodyCnt = 0
'                  iSoft.headerRec.Trigger = "STARTRFMESS"
'                  iSoft.headerRec.dateTime = dateTime
'                  iSoft.headerRec.Control = nextControlRef
'               End If
'
'            Case "Body"
'               ReDim iSoft.bodyRec(bodyCnt)
'               With iSoft.bodyRec(bodyCnt)
'                  .Trigger = "REPORT"
'                  .ReportId = ReportId
'                  .State = ReportStatus
'                  .dateTime = dateTime
'               End With
'               bodyCnt = bodyCnt + 1
'
'            Case "Trailer"
'               With iSoft
'                  With .headerRec
'                     fileData = .Trigger & .dateTime & .Control
'                  End With
'
'                  For i = 0 To UBound(.bodyRec)
'                     With .bodyRec(i)
'                        fileData = fileData & .Trigger & .ReportId & .State & .dateTime
'                     End With
'                  Next i
'
'                  fileData = fileData & "ENDRFMESS" & nextControlRef
'               End With
'               blnWriteFile = True
'
'         End Select
'
'   End Select
'   If blnWriteFile Then
'      strSQL = "SELECT * " & _
'               "FROM EDI_Local_Trader_Settings " & _
'               "WHERE EDI_LTS_Index = " & LocalTrader
'      RS.Open strSQL, ICEcon, adOpenKeyset, adLockReadOnly
'      nextControlRef = RS!Trader_Ack_Control
'      RS.Close
'
'      FILENAME = fs.BuildPath(FileLocation, "rfmess." & nextControlRef)
'
'      Open FILENAME For Output As #1
'      Print #1, fileData
'      Close #1
'   End If
'End Sub




