VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "errorControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum AppLogModes
   ELM_NORMAL = 0
   ELM_NAMEDFILE = 1
   ELM_OVERWRITE = 2
   ELM_OFF = 3
End Enum

Public Enum EnumStatusFlags
   SF_LOG_ONLY = 0
   SF_RECOVERABLE = 1
   SF_FATAL = 2
End Enum

Public Enum BehaviourTypes
   DEBUG_AT_SOURCE = -1
   EVALUATE_AND_RETURN = 0
   EVALUATE_AND_CONTINUE = 1
End Enum

Private domRoot As MSXML2.IXMLDOMElement
Private sessionRoot As MSXML2.IXMLDOMElement
Private logStatus As String
Private groupEl As MSXML2.IXMLDOMElement
Private curLogFile As String
Private LogMode As AppLogModes
Private FrmHdr As String
Private logPath As String
Private xmlDoc As MSXML2.DOMDocument

Private Sub Class_Initialize()
   With logBuf
      .Terminator = vbCrLf
   End With

   Set groupEl = Nothing
   blnLogOverWrite = False
   intLogMode = 2
   blnLogErrors = True
   intBehaviour = 0
   statusFlag = SF_FATAL
   errorEvaluated = False
   FrmHdr = "IceLabcomm Suite"
   Set xmlDoc = CreateObject("MSXML2.DOMDocument")
End Sub

Public Function Add(errNum As Long, _
                    errDesc As String, _
                    errSrc As String, _
                    Optional PassThru As Boolean = True) As Long
'  Add this error to the error collection module
   Dim errResult As Long
   Dim i As Integer
   Dim adoDesc As String
   Dim adoRes As Long
   
   errResult = errNum
   If errorEvaluated = False Then
      strDesc = errDesc
      strSrc = errSrc
   End If

'  Add the vb error to the collection
   Set eData = New errorData
   eData.Add procName, errResult, strDesc, strSrc, extraInfo
   colErr.Add eData

'  Is this an SQL error?
   If errResult < -2000000000 Then
      If TypeName(dbCon) <> "Nothing" Then
         For i = 0 To dbCon.Errors.Count - 1
            Set eData = New errorData
            With dbCon.Errors(i)
               adoRes = .NativeError
               Select Case adoRes
                  Case 0
                     If .SQLState = "HYT00" Then
                        adoDesc = "Database Access Timeout"
                        errResult = 3234
                     ElseIf .SQLState = "08S01" Then
                        adoDesc = "Database connection lost"
                        errResult = 3234
                     End If
                     
                  Case 105
                     adoDesc = "Missing quotation mark from SQL string"
                  
                  Case 207
                     adoDesc = "Column does not exist in table"
                     
                  Case 170
                     adoDesc = "Syntax error in query"
                     
                  Case 201
                     adoDesc = "Parameter error - check length values for nulls"
                     
               End Select
               eData.Add "SQL Server", adoRes, adoDesc, "ADO", "Server State = " & .SQLState
            End With
            colErr.Add eData
            Set eData = Nothing
         Next i
      
      Else
         Set eData = New errorData
         errResult = 3263
         eData.Add "Database Initialisation", errResult, "Connection problems - Check the UDL file", errSrc
         colErr.Add eData
         Set eData = Nothing
      End If
   End If
     
   Set eData = Nothing

'  Any specific action to take?
   If errorEvaluated = False Then
      If TypeName(errAction) <> "Nothing" Then
         errResult = errAction.Evaluate(errResult, extraInfo, statusFlag)
      End If
      
      errorEvaluated = True
   End If
   
'  Reset to avoid repetition
   extraInfo = "<No further information>"
   If TypeName(dbCon) <> "Nothing" Then
      dbCon.Errors.Clear
   End If
   
'  Deal with the error or pass it on?
   If errResult = 0 Then
'     The evaluation routine has handled the error - no further action required.
      errorEvaluated = False
      Set colErr = Nothing
   Else
      If PassThru Then
'        Pass this error to the next error handler
         Err.Raise errResult
      Else
'        Display the error details
         Show errResult
      End If
   End If
End Function

Public Property Let Behaviour(intNewValue As BehaviourTypes)
'  How do we want the handler to perform?
'  -1 = Debug
'  0 = Evaluate and generate run-time errors
'  1 = Evaluate and warn on fail
   If (TypeName(errAction) = "Nothing") And (intNewValue = EVALUATE_AND_RETURN) Then
      LogMessage "No error exception procedure specified - behaviour not amended"
   Else
      intBehaviour = intNewValue
   End If
End Property
   
Public Property Get Behaviour() As Long
   Behaviour = intBehaviour
End Property

Public Sub ClearErrors()
   Set colErr = Nothing
   errorEvaluated = False
End Sub

Public Property Let CurrentProcedure(strNewValue As String)
'  The name of the procedure producing the error
   procName = strNewValue
End Property

Public Property Let dbConnection(ByRef objNewValue As ADODB.Connection)
   Set dbCon = objNewValue
End Property

Friend Function ErrorDetails(ErrorIndex As Integer) As Variant
   Dim eData As errorData
   
   Set eData = colErr(ErrorIndex)
   ErrorDetails = eData.Retrieve
   Set eData = Nothing
End Function

Public Sub ApplicationLogging(Optional LogFile As String = "", _
                              Optional ModeId As AppLogModes = ELM_NORMAL)
'  Mode 0 = Off
'  Mode 1 = Daily log file
'  Mode 2 = Continuous log file
'  Mode 3 = Overwrite existing file
   
   Dim cNode As MSXML2.IXMLDOMElement
   Dim nAtt As MSXML2.IXMLDOMAttribute
   
   Set xmlDoc = New MSXML2.DOMDocument
   
   If logPath = "" Then
      logPath = App.Path
   End If
   
   If LogFile = "" Then
      LogFile = "ErrorLog.xml"
   End If
   
   modeFlag = ModeId
   
   If ModeId <> ELM_OFF Then
      If ModeId = ELM_NORMAL Then
         LogFile = fs.BuildPath(logPath, fs.GetBaseName(LogFile) & Format(Now(), "_yyyymmdd") & ".xml")
      Else
         If fs.GetExtensionName(LogFile) <> "xml" Then
            LogFile = fs.BuildPath(logPath, fs.GetBaseName(LogFile) & ".xml")
         End If
         
         If ModeId = ELM_NAMEDFILE Then
            If fs.FileExists(LogFile) Then
               fs.DeleteFile LogFile
            End If
         End If
      End If
   Else
      LogFile = ""
      logPath = ""
   End If
   
   curLogFile = LogFile
   LogMode = ModeId
   
   logStatus = "New"
   NewLogFile
   
   If LogMode <> ELM_OFF Then
      With xmlDoc
         .Load curLogFile
         Set domRoot = .selectSingleNode("/LogData")
                  
         Set cNode = .createElement("Session")
         
         Set nAtt = .createAttribute("Status")
         nAtt.Value = logStatus
         cNode.setAttributeNode nAtt
         
         Set nAtt = .createAttribute("DateTime")
         nAtt.Value = Format(Now(), "dd/mm/yyyy hh:nn:ss")
         cNode.setAttributeNode nAtt
         
         domRoot.appendChild cNode
         
         Set sessionRoot = .selectSingleNode("/LogData").lastChild
      End With
      
      logStatus = "Continued"
   End If
End Sub

Public Property Let ExceptionModule(ErrorTrap As Object)
'  The class in the calling procedure that evaluates the errors
   Set errAction = ErrorTrap
End Property

Public Property Let LogLevel(eNewValue As EnumStatusFlags)
'  Set from the program code and passed into the evaluation routine.

'  Permissable values
'  0 = Standard error handling
'  1 = Log all furtherInfo amendments
   statusFlag = eNewValue
End Property

Public Property Get LogLevel() As Long
   LogLevel = statusFlag
End Property

Public Property Let FormCaption(strNewValue As String)
   FrmHdr = strNewValue
End Property

Public Property Get FormCaption() As String
   FormCaption = FrmHdr
End Property

Public Property Let FurtherInfo(strNewValue As String)
'  Descriptive text to add the trace/debugging process
   extraInfo = strNewValue
   If statusFlag = 1 Then
      LogMessage extraInfo
   End If
End Property

Public Property Get FurtherInfo() As String
   FurtherInfo = extraInfo
End Property

Public Property Let GroupElement(elementName As String)
   Set groupEl = xmlDoc.createElement(elementName)
End Property

Public Property Let LogDirectory(strNewValue As String)
   Dim LogFile As String
   Dim logRoot As MSXML2.IXMLDOMElement
   
   If logPath <> "" Then
      LogFile = fs.GetFileName(curLogFile)
      fs.DeleteFile curLogFile
      curLogFile = fs.BuildPath(strNewValue, LogFile)
      
      If fs.FileExists(curLogFile) Then
         Set logRoot = xmlDoc.selectSingleNode("/LogData/Session")
         xmlDoc.Load curLogFile
         xmlDoc.selectSingleNode("/LogData").appendChild logRoot
      End If
      
      xmlDoc.Save curLogFile
   End If
   
   logPath = strNewValue
End Property

Public Function LogError() As String
   Dim i As Integer
   Dim vData As Variant
   Dim eData As errorData
   Dim strPath As String
   Dim strADO As String
   
   For i = 1 To colErr.Count
      Set eData = colErr(i)
      vData = eData.Retrieve
      
      If vData(0) = "SQL Server" Then
         strADO = strADO & vData(0) & ": Native error = " & vData(2) & " (" & vData(3) & ")" & vbCrLf
      Else
         strPath = strPath & vData(0) & " (" & vData(1) & ")" & vbCrLf
      End If
         
      LogMessage CStr(vData(3)), _
                 CStr(vData(0)), _
                 CStr(vData(1)), _
                 CStr(vData(2)), _
                 CStr(vData(4))
'      With logBuf
'         .DateTime = Format(Now(), "dd/mm/yyyy hh:nN:ss")
'         .Procedure = vData(0)
'         .Operation = vData(1)
'         .Description = vData(3)
'         .Number = vData(2)
'         .Source = vData(4)
'         .MsgData = "Error Report"
'      End With
'      LogMessage vbLogEventTypeError
'
      Set eData = Nothing
   Next i
   strPath = strADO & vData(3) & vbCrLf & "Error Path:" & vbCrLf & strPath
   LogError = strPath
End Function

Public Sub LogMessage(Optional Description As String = "", _
                      Optional ProcId As String = "IceMsg", _
                      Optional Operation As String = "Logging", _
                      Optional Number As String = "", _
                      Optional Source As String)
   Dim pNode As MSXML2.IXMLDOMElement
   Dim cNode As MSXML2.IXMLDOMElement
   Dim nAtt As MSXML2.IXMLDOMAttribute
   Dim nTxt As MSXML2.IXMLDOMText
   
   If LogMode <> ELM_OFF Then
      NewLogFile
      
      With xmlDoc
         Set pNode = .createElement("Data")
         
         Set nAtt = .createAttribute("Procedure")
         nAtt.Value = ProcId
         pNode.setAttributeNode nAtt
         
         Set nAtt = .createAttribute("Operation")
         nAtt.Value = Operation
         pNode.setAttributeNode nAtt
         
         Set nTxt = .createTextNode(Format(Now(), "dd/mm/yyyy hh:nn:ss"))
         pNode.appendChild nTxt
         
         Set cNode = .createElement("Description")
         
         If Number <> "" Then
            Set nAtt = .createAttribute("Number")
            nAtt.Value = Number
            cNode.setAttributeNode nAtt
         End If
         
         If Source <> "" Then
            Set nAtt = .createAttribute("Source")
            nAtt.Value = Source
            cNode.setAttributeNode nAtt
         End If
         
         Set nTxt = .createTextNode(Description)
         cNode.appendChild nTxt
               
         pNode.appendChild cNode
                  
         sessionRoot.appendChild pNode
         
         .Save curLogFile
      End With
   End If
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = colErr.[_NewEnum]
End Function

Private Function NewLogFile() As Boolean
   
   Dim fl As File
   Dim cNode As MSXML2.IXMLDOMElement
   Dim pNode As MSXML2.IXMLDOMNode
   Dim NPI As MSXML2.IXMLDOMProcessingInstruction
   
   Dim nAtt As MSXML2.IXMLDOMAttribute
   Dim nTxt As MSXML2.IXMLDOMText
   
   Dim blnNewLog As Boolean
   
   If LogMode <> ELM_OFF Then
      If fs.FileExists(curLogFile) Then
         Set fl = fs.GetFile(curLogFile)
         blnNewLog = (DateDiff("D", fl.DateCreated, Now()) > 0)
      Else
         blnNewLog = True
      End If
      
      If blnNewLog Then
         
         xmlDoc.Save curLogFile
         Set xmlDoc = Nothing
         Set xmlDoc = New MSXML2.DOMDocument
         
         With xmlDoc
            Set NPI = .createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
            .appendChild NPI
         
            Set domRoot = .createNode(NODE_ELEMENT, "LogData", "")
            .appendChild domRoot
         
            Set nAtt = .createAttribute("Version")
            nAtt.Value = "1.0"
            domRoot.setAttributeNode nAtt
         
            Set nAtt = .createAttribute("xmlns:xsi")
            nAtt.Value = "http://www.w3.org/2001/XMLSchema-instance"
            domRoot.setAttributeNode nAtt
            
            Set nAtt = .createAttribute("LogDate")
            nAtt.Value = Format(Now(), "ddd dd mmm yyyy")
            domRoot.setAttributeNode nAtt
            
            .Save curLogFile
         End With
      End If
   End If
End Function

Public Sub Show(errNo As Long)
'  Retrieve all the errors from the errordata object and display
   Dim i As Integer
   Dim strPath As String
   Dim strADO As String
   Dim eData As errorData
   Dim vData As Variant
   Dim tPos As Integer
   Dim ctrl As Integer
   
   tPos = 1680
   For i = 1 To colErr.Count
      ctrl = i - 1
      Set eData = colErr(i)
      vData = eData.Retrieve

      If vData(0) = "SQL Server" Then
         strADO = strADO & vData(0) & ": Native error = " & vData(2) & " (" & vData(3) & ")" & vbCrLf
      Else
         strPath = strPath & vData(0) & " (" & vData(1) & ")" & vbCrLf
      End If
            
      LogMessage CStr(vData(3)), _
                 CStr(vData(0)), _
                 CStr(vData(1)), _
                 CStr(vData(2)), _
                 CStr(vData(4))
      Set eData = Nothing
   Next i
   
   strPath = strADO & vData(3) & vbCrLf & "Error Path:" & vbCrLf & strPath
   
   Select Case intBehaviour
      Case 0   '  Display the error form
         frmErrorReport.ErrorClass = Me
         frmErrorReport.Show 1
         LogMessage strPath, , "ErrorPath"
      
      Case 1   '  Try to deal with the failure
         errAction.HandleFailure errNo, strPath
      
      Case 2   '  Generate a run-time error
         Err.Raise vData(2), vData(4), strPath
      
   End Select
   
   Set colErr = Nothing
   errorEvaluated = False
End Sub
