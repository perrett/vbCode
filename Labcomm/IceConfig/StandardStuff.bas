Attribute VB_Name = "StandardStuff"
Option Explicit
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const SW_RESTORE = 9
Private Declare Function apiGetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public RepQue() As Long
Public REPRS As ADODB.Recordset
Public Type RepRec
    EDI_Report_Index As Long
    EDI_Provider_Org As String
    EDI_Loc_Nat_Code_To As String
    EDI_Clin_Nat_Code_To As String
End Type

Public Const ERROR_BAD_DEVICE           As Long = 1200
Public Const ERROR_CONNECTION_UNAVAIL   As Long = 1201
Public Const ERROR_EXTENDED_ERROR       As Long = 1208
Public Const ERROR_MORE_DATA            As Long = 234
Public Const ERROR_NOT_SUPPORTED        As Long = 50
Public Const ERROR_NO_NET_OR_BAD_PATH   As Long = 1203
Public Const ERROR_NO_NETWORK           As Long = 1222
Public Const ERROR_NOT_CONNECTED        As Long = 2250
Public Const NO_ERROR                   As Long = 0

'This API returns a UNC from a drive letter
Declare Function WNetGetConnection Lib "mpr.dll" Alias _
    "WNetGetConnectionA" _
    (ByVal lpszLocalName As String, _
    ByVal lpszRemoteName As String, _
    cbRemoteName As Long) As Long


Dim WildCArd As String

Public Function CLDate(istr As String) As String
 Dim Mon As String
 Dim NewMon As String
 ' date should be in format dd mmm yyyy hh:mm:ss
 Mon = Mid$(istr, 4, 3)
 NewMon = Mon
 Select Case UCase(Mon)
    Case "MAJ"
        NewMon = "May"
    Case "OKT"
        NewMon = "Oct"
 End Select
CLDate = Mid$(istr, 1, 3) + NewMon + Mid$(istr, 7, Len(istr) - 6)
End Function
'checks to see if this is the only version of this application running and brings the other to front if not
'called from the main form with "Check_single_application me"
Public Sub Check_Single_Application(Name As Form)
    Dim Sname As String
    Dim hwnd As Long
    Dim lRetVal As Long
    Sname = Name.Caption
    If App.PrevInstance Then
        App.Title = "NewCopy"
        Name.Caption = "NewCopy"
        MsgBox "Another Version of '" + Sname + "' is already loaded", vbOKOnly
        hwnd = FindWindow(vbNullString, Sname)
        If hwnd <> 0 Then
            lRetVal = ShowWindow(hwnd, SW_RESTORE)
            lRetVal = SetForegroundWindow(hwnd)
            End If
        End
        End If
End Sub

'Write to a ini File

Function Write_Ini_Var(section, varname As String, value As String, Filename) As Boolean
   Dim Success, Message, Response
   Filename = UCase$(Filename)
        
Writeagain:
  
   Success = WritePrivateProfileString(section, varname, value, Filename)

   If Success Then
      Write_Ini_Var = True
   Else
      Message = "Unable to Write Profile"
      Response = MsgBox(Message, 277, "Error")
      If Response = 4 Then
         GoTo Writeagain
      End If
      Write_Ini_Var = False
   End If

End Function

'Read the INI file
Function Read_Ini_Var(section, varname, Filename)
 
 'this function gets a profile varable from either the
 '"WIN.INI" or a private "*.INI" file if the profile
 'does not exist an error message is displayed and the
 'function returns a null string

Dim Message, Success
Dim Sname$, Vname$, Fname$, ret$, kname$
Dim Response

    ret$ = String$(255, 0)
    Filename = UCase$(Filename)
    Sname$ = section
    kname$ = varname
    Success = GetPrivateProfileString(Sname$, kname$, "", ret$, Len(ret$), Filename)
    If Success Then
        ret$ = Left$(ret$, Success)
        Read_Ini_Var = ret$
        Message = "[" & section & "]" & " " & varname & "=" & ret$
    Else
        Read_Ini_Var = ""
    End If
End Function

'******** Returns the Full windows directory *************
Function Get_WIN_Dir()

'returns the WINdows directory

Dim Success
Dim ret$

        ret$ = String$(144, 0)

        Success = GetWindowsDirectory(ret$, Len(ret$))
     
   If Success Then
            Get_WIN_Dir = Left$(ret$, Success)
   Else
            Get_WIN_Dir = ""
   End If
   
             
End Function

Function GLD(Old_Date As String) As String
    Dim Test_Date As Date
    If Trim(Old_Date) = "" Then
        GLD = ""
        Exit Function
        End If
    Test_Date = Trim(Old_Date)
    If Test_Date > Date Then
        Test_Date = DateAdd("yyyy", -100, Test_Date)
        End If
    GLD = Format(Test_Date, "dd MMM yyyy")
End Function

Public Function Get_Age(calc_datet, dobt As Date)

Dim calc_date, dob As Date
Dim STemp As Date
calc_date = Format(calc_datet, "dd mmm yyyy")
dob = Format(dobt, "dd mmm yyyy")
If dob > Date Then
    dob = DateAdd("yyyy", -100, dob)
    End If
'On Error GoTo weird_age
 Get_Age = DateDiff("YYYY", dob, calc_date) '- (calc_date < DateSerial(Year(calc_date), Month(dob), Day(dob)))
 STemp = DateAdd("yyyy", Get_Age, dobt)
 If STemp > calc_datet Then
    Get_Age = Get_Age - 1
    End If
  If Get_Age = 0 Then
      Get_Age = ""
      End If
  Exit Function
End Function

Public Function ValidateFilepath(fPath As String) As Boolean
'  Function takes the supplied file path and creates all folders necessary
'  from the root directory downwards, to make the path valid.
'  A false return value indicates the original path was invalid.
   Dim fldr As String
   Dim oldFldr As String
   
   If fs.FolderExists(fPath) Then
      ValidateFilepath = True
   Else
      fldr = fs.GetParentFolderName(fPath)
         
      Do Until fs.FolderExists(fldr)
         Do Until fs.FolderExists(fldr)
            oldFldr = fldr
            fldr = fs.GetParentFolderName(fldr)
         Loop
         fs.CreateFolder (oldFldr)
         fldr = fs.GetParentFolderName(fPath)
      Loop
      fs.CreateFolder fPath
      ValidateFilepath = False
   End If

End Function

Public Function RunningInIDE() As Boolean
   On Error Resume Next
   Debug.Print 1 / 0
   RunningInIDE = (Err.Number <> 0)
   'RunningInIDE = False
End Function

Function GetUNCPath(ByVal strDriveLetter As String, _
                    ByRef strUNCPath As String) As Long

On Local Error GoTo GetUNCPath_Err

    Dim strMsg As String
    Dim lngReturn As Long
    Dim strLocalName As String
    Dim strRemoteName As String
    Dim lngRemoteName As Long

    strLocalName = strDriveLetter
    strRemoteName = String$(255, Chr$(32))
    lngRemoteName = Len(strRemoteName)

    'Attempt to grab UNC
    lngReturn = WNetGetConnection(strLocalName, _
                                  strRemoteName, _
                                  lngRemoteName)

    If lngReturn = NO_ERROR Then
        'No problems - return the UNC
        'to the passed ByRef string
        GetUNCPath = NO_ERROR
        strUNCPath = Trim$(strRemoteName)
        strUNCPath = Left$(strUNCPath, Len(strUNCPath) - 1)
    Else
        'Problems - so return original
        'drive letter and error number
        GetUNCPath = lngReturn
        strUNCPath = strDriveLetter & "\"
    End If
    
GetUNCPath_End:
    Exit Function
    
GetUNCPath_Err:
    GetUNCPath = ERROR_NOT_SUPPORTED
    strUNCPath = strDriveLetter
    Resume GetUNCPath_End
    
End Function

Public Function GetMachineName() As String
'Returns the computername
Dim lngLen As Long, lngX As Long
Dim strCompName As String
    lngLen = 16
    strCompName = String$(lngLen, 0)
    lngX = apiGetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        GetMachineName = Left$(strCompName, lngLen)
    Else
        GetMachineName = ""
    End If
End Function


