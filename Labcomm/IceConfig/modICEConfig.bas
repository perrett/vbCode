Attribute VB_Name = "modICEConfig"
Option Explicit

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Public Const SPI_GETWORKAREA& = 48

Public Declare Function SystemParametersInfo Lib "user32" _
   Alias "SystemParametersInfoA" _
  (ByVal uAction As Long, _
   ByVal uParam As Long, _
   lpvParam As Any, _
   ByVal fuWinIni As Long) As Long

'##NO_ERR_RAISE_TRAP
Public Enum ENUM_NodeLevel
   nl_ROOT = 0
   nl_FIRST = 1
   nl_SECOND = 2
   nl_THIRD = 3
   nl_FOURTH = 4
End Enum

Public Enum ENUM_MenuStatus
   ms_BOTH = -1
   ms_DISABLED = 0
   ms_ADD = 1
   ms_DELETE = 2
End Enum

Public Enum ENUM_GenericClassTypes
   GC_Practice = 0
   GC_Specialty = 1
   GC_MsgType = 2
   GC_Individual = 3
End Enum

Const DB_CONNECT_STRING = "FILE NAME=[NAME_TOKEN]"

Public Const BPGREEN = &H8000&
Public Const BPBLUE = &HC00000
Public Const BPRED = &HFF&
Public Const SW_SHOWNORMAL = 0

Public blnUseDMO As Boolean
Public rCtrl As requeueControl
Public sqlserver As SQLDMO.sqlserver
Public sqlDb As SQLDMO.Database
Public UDLServer As Variant
Public UDLDatabase As String
Public dbUser As String
Public dbPass As String
Public DBVERSION As String

Public scriptDir As String
Public styleDir As String

Public phoenix As Boolean
Public logHist As Integer
Public logSortOn As String
Public logSortDesc As Boolean
Public transId As New TransactionControl
Public formHeader As String
Public blnShowBrowser As Boolean
Public userID As String
Public filepathToUNC As Boolean
Public Requeuestr As String
Public connectString As String
Public iniFile As String
Public blnMultiUDL As Boolean
Public DB_UDL_FILE As String
Public DB_Name As String
Public ConfigPath As String
Public iceCon As ADODB.Connection
Public LogBackColour
Public lastNode As MSComctlLib.Node
Public fView As New frameDataClass
Public eClass As AHSLErrorLog_XML.errorControl
Public etrap As New errorExceptions
Public fs As New FileSystemObject

Public blnUseDTS As Boolean
Public blnClinNameUseAll As Boolean

Public intUseRCIndex As Integer
'Public objTView As New NewTreeClass
Public objTV As New TreeNodeControl
Public loadCtrl As Object
'Public objctrl As Class1
'Public tPList As PropertiesList
'Public blnAddStatus As Boolean
'Public strOldNatId As String
'Public strOldCode As String
'Public strOldMsg As String
'Public strOldKey As String
Public DefOrgID As String
Public PickedCol As Long
Public PickedColIndex As Integer
Public PickedColName As String
Public PickedProv As String
Public PickedProvIndex As Integer
Public PickedTube As String
Public PickedTubeIndex As Integer
Public PickedTubeCol As Long
'Public PickedOverride As String
'Public strPropVal As String
'Public GVSP As Boolean
Public transCount As Integer
Public logLevel As Long
Public errRet As Long
Public conTimeOut As Long


Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Public Const DIV = "|"

Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL


' Reg Key ROOT Types...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Global Const ERROR_SUCCESS = 0
Global Const REG_SZ = 1                         ' Unicode nul terminated string
Global Const REG_DWORD = 4                      ' 32-bit number

Global Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Global Const gREGVALSYSINFOLOC = "MSINFO"
Global Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Global Const gREGVALSYSINFO = "PATH"

Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Public Sub Main()
   On Error GoTo procEH
   Dim debugState As String
   Dim eTemp As String
   Dim src As String
   Dim logMode As String
   Dim logStat As Long
   Dim logPath As String
   Dim logSrc As String
   Dim fldr As String
   Dim pFldr As String
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim tmpINI As String
   Dim strArray() As String
   
   If App.PrevInstance Then
      MsgBox "An instance of IceConfig is already running on this machine", vbInformation, "Duplicate instance"
   Else
      Set eClass = CreateObject("AHSLErrorLog_XML.errorControl")
      iniFile = App.Path & "\ICEConfig.ini"
      eClass.ExceptionModule = etrap
      eClass.FurtherInfo = "Read INI file settings"
   
      eTemp = Val(Read_Ini_Var("Logging", "Behaviour", iniFile))
      
      If IsNumeric(eTemp) Then
         If RunningInIDE = False Then
            eTemp = Abs(Val(eTemp))
         End If
      Else
         eTemp = 0
      End If
      
'      intUseRCIndex = (Read_Ini_Var("GENERAL", "UseLabReadCodes", iniFile) = 1)
      blnUseDTS = (Read_Ini_Var("General", "UseDTS", iniFile) = 1)
      
      eClass.Behaviour = eTemp
      
      Set iceCon = frmDatabase.ConnectionDetails(sqlserver, _
                                                 sqlDb, _
                                                 dbUser, _
                                                 dbPass, _
                                                 UDLServer, _
                                                 UDLDatabase)
      
      If iceCon Is Nothing Then
         MsgBox frmDatabase.ConnectionError, vbCritical, "Database Connection Error"
         End
      End If
      
      blnUseDMO = frmDatabase.DMOStatus
      blnClinNameUseAll = False
      'conTimeOut = frmDatabase.DatabaseTimeOut
      
      If Not sqlDb Is Nothing Then
         Set sqlDb = sqlserver.Databases(iceCon.DefaultDatabase)
      End If
      
      Unload frmDatabase
      
      strSQL = "SELECT Version FROM dbVersion"
      
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount > 0 Then
         RS.MoveLast
      End If
      
      formHeader = "Ice Configuration - Using: " & iceCon.DefaultDatabase & " at version " & RS!Version
      
      RS.Close
      
      Set RS = Nothing
      eClass.DBConnection = iceCon
      
   '   CheckVersion
      
      logMode = Read_Ini_Var("Logging", "Mode", iniFile)
      
      If logMode = "" Then
         logLevel = 3
      Else
         logLevel = Val(logMode)
      End If
      
      eClass.LogDirectory = App.Path
      eClass.ApplicationLogging "IceConfig", logLevel
      
      transId.DBConnection = iceCon
      
      filepathToUNC = (Read_Ini_Var("General", "PathToUNC", iniFile) = 1)
      
      DefOrgID = Read_Ini_Var("General", "DefOrgID", iniFile)
      
      phoenix = (Read_Ini_Var("General", "Phoenix", iniFile) = 1)
      
      tmpINI = Read_Ini_Var("General", "LogHistory", iniFile)
      If IsNumeric(tmpINI) Then
         If Val(tmpINI) = 0 Then
            logHist = 10
         Else
            logHist = Abs(tmpINI)
         End If
      Else
         logHist = 10
      End If
      
      tmpINI = Read_Ini_Var("General", "LogSortOn", iniFile)
      If Left(tmpINI, 1) = "-" Then
         logSortOn = Mid(tmpINI, 2)
         logSortDesc = True
      Else
         logSortOn = tmpINI
         logSortDesc = False
      End If
      
      
      frmSplash.Show
   End If
   
   Exit Sub
   
procEH:
   If eClass Is Nothing Then
      MsgBox "Unable to create error handler", vbCritical, "AHSLErrorlog.dll not present or unregistered"
      End
   Else
      If eClass.Behaviour = -1 Then
         Stop
         Resume
      End If
      eClass.Add Err.Number, Err.Description, Err.Source
      HandleError "modICEConfig.Main"
   End If
End Sub
'
'Public Function ReadUDL(FILENAME As String, _
'                        Optional SetConnection As Boolean = True) As String
'   Dim pTag As String
'   Dim pos As Long
'   Dim clen As Integer
'   Dim buf As String
'   Dim conStr As String
'   Dim dbSource As String
'   Dim tmpCon As New ADODB.Connection
'   Dim dblink As New MSDASC.DataLinks
'
'   conStr = ""
'   pTag = "P" & Chr(0) & _
'          "r" & Chr(0) & _
'          "o" & Chr(0) & _
'          "v" & Chr(0) & _
'          "i" & Chr(0) & _
'          "d" & Chr(0) & _
'          "e" & Chr(0) & _
'          "r"
'   If fs.FileExists(FILENAME) Then
'      buf = Space(1500)
'      Open FILENAME For Binary As #1
'      Get #1, , buf
'      Close #1
'      pos = InStr(1, buf, pTag, vbBinaryCompare)
'      conStr = Replace(Mid(buf, pos), Chr(0), "")
'      pos = InStr(1, conStr, "Source=") + 7
'      clen = InStr(pos, conStr, vbCrLf) - pos
'      dbSource = Mid(conStr, pos, clen)
'
'      If SetConnection Then
'         tmpCon.ConnectionString = conStr
'         dblink.PromptEdit tmpCon
'         conStr = tmpCon.ConnectionString
'      Else
'         pos = InStr(1, conStr, "Initial Catalog") + 16
'         clen = InStr(pos, conStr, ";") - pos
'         conStr = Mid(conStr, pos, clen) & "|" & dbSource
'      End If
'
'   Else
'      MsgBox "Specified UDL (" & FILENAME & ") not found", _
'             vbExclamation, "Invalid UDL"
'   End If
'
'   ReadUDL = conStr
'   Set tmpCon = Nothing
'   Set dblink = Nothing
'End Function
'
'Public Function SetAsIce(NonIceUDl As String, _
'                         Connection_String As Boolean)
'   Dim iceFile As String
'   Dim buf As String
'   Dim fileHdr As String
'
'   iceFile = fs.BuildPath(ConfigPath, "ice.udl")
'
'   If Connection_String Then
'      buf = Space(1500)
'      Open iceFile For Binary As #1
'      Get #1, , buf
'      Close #1
'
'      fileHdr = Left(buf, 128) & StrConv(NonIceUDl, vbUnicode)
'      fs.DeleteFile iceFile
'      Open iceFile For Binary As #1
'      Put #1, , fileHdr
'      Close #1
'   Else
'      fs.CopyFile fs.BuildPath(ConfigPath, NonIceUDl), iceFile
'   End If
'End Function

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Function FormLoaded(frmId As String) As Boolean
   Dim frm As Form
   Dim blnLoaded As Boolean
   
   blnLoaded = False
   
   For Each frm In Forms
      If StrComp(frm.Name, frmId, vbTextCompare) = 0 Then
         blnLoaded = True
         Exit For
      End If
   Next
   FormLoaded = blnLoaded
End Function

Public Function GetConnection(Optional SetUpConnection As Boolean = True) As ADODB.Connection
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim iceCon As ADODB.Connection
   Dim UDLPath As String
   Dim UDLFile As String
   Dim IceUDL As String
   
   Dim ConfigPath As String
      
   Set iceCon = frmDatabase.ConnectionDetails(sqlserver, _
                                              sqlDb, _
                                              dbUser, _
                                              dbPass, _
                                              UDLServer, _
                                              UDLDatabase)
'      With iceCon
'         .CursorLocation = adUseClient
'         .Mode = adModeReadWrite
'         .CommandTimeout = 120
'         .Open "FILE NAME=" & UDLFile
'      End With
      
   Set GetConnection = iceCon
'   UDLPath = Read_Ini_Var("General", "ConfigPath", iniFile)
'
'   If UDLPath = "" Then
'      UDLPath = App.Path
'   ElseIf UCase(fs.GetExtensionName(UDLPath)) = "UDL" Then
'      UDLPath = fs.GetParentFolderName(UDLPath)
'   End If
'
'   UDLFile = fs.BuildPath(UDLPath, "ice.udl")
'
'   If fs.FileExists(UDLFile) Then
'      If SetUpConnection Then
'         If (Read_Ini_Var("UDLData", "MultiUDL", iniFile) = 1) Then
'            frmDatabase.ConfigPath = UDLPath
'            frmDatabase.Show 1
'            UDLFile = frmDatabase.UDLFile
'            Unload frmDatabase
'         Else
'            ReadUDL UDLFile
'         End If
'      End If
'
'      Set ICECon = New ADODB.Connection
'
'      With ICECon
'         .CursorLocation = adUseClient
'         .Mode = adModeReadWrite
'         .CommandTimeout = 120
'         .Open "FILE NAME=" & UDLFile
'      End With
'
'      Set GetConnection = ICECon
'   Else
'      MsgBox "UDL shown in title bar not found. Please check the 'ConfigPath' in IceConfig.ini", vbExclamation, UDLFile
'      End
'   End If
   Exit Function

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "modICEConfig.GetConnection"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Sub WriteLog(LogString As String)
    Open App.Path + "\Reporting.LOG" For Append As #1
    Print #1, Format(Now, "DD/MM/YYYY HH:MM:SS") & ": " & LogString
    Close #1
End Sub

Public Function GetOrganisationName(OrgID As String) As String
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "SELECT Organisation_Name " & _
            "FROM Organisation " & _
            "WHERE Organisation_National_Code = '" & frmMain.cboTrust.Text & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   GetOrganisationName = Trim(RS!organisation_Name & "")
   RS.Close
   Set RS = Nothing
End Function

Public Sub HandleError(ProcedureId As String, _
                       Optional Rollback As Boolean)
'   If Rollback Then
      transId.AbandonTransaction
'   End If
   If Not (iceCon Is Nothing) Then
      iceCon.Close
      Set iceCon = Nothing
      Set iceCon = GetConnection(False)
      eClass.DBConnection = iceCon
   End If
   frmMain.MousePointer = vbNormal
End Sub

Public Function dbObject(ProcOrTable As String) As String
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   strSQL = "SELECT Name, Type " & _
            "FROM dbo.sysobjects " & _
            "WHERE Name = '" & ProcOrTable & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   dbObject = "not in Database"
   Do While Not RS.EOF
      If Trim(RS!Type) = "P" Then
         dbObject = "StoredProcedure"
      ElseIf Trim(RS!Type) = "U" Then
         dbObject = "UserTable"
         Exit Do
      End If
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
End Function
'
'Public Function DMOAvailable() As Boolean
'   On Error GoTo procEH
'   blnUseDMO = True
'   Set sqlServer = CreateObject("SQLDMO.sqlServer")
'   DMOAvailable = blnUseDMO
'
'   Exit Function
'
'procEH:
'   blnUseDMO = False
'   Resume Next
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
         frmDatabase.cboDB.ListIndex = 0
      End If
   Else
      MsgBox "Specified UDL (" & UDLFile & ") not found", _
             vbExclamation, "Invalid UDL"
   End If
   
   ReadUDL = constr
End Function

Public Sub RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long)
    Call ShellExecute(GetDesktopWindow(), sTopic, sFile, sParams, sDirectory, nShowCmd)
End Sub

Public Sub HelpDesk_Click()
   Dim HelpDeskURL As String
   
   HelpDeskURL = Read_Ini_Var("General", "HelpDeskURL", iniFile) '  "<Read_From_INI_File, for testing, set to http://nhsnetserver/helpdesk>"
   If HelpDeskURL = "" Then
      MsgBox "No entry for 'HelpDeskURL' in the ini file: (" & iniFile & ")", vbInformation, "Missing INIfile option"
   Else
      Call RunShellExecute("open", HelpDeskURL, 0&, 0&, SW_SHOWNORMAL)
   End If
End Sub


