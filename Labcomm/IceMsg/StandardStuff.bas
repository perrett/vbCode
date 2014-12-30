Attribute VB_Name = "StandardStuff"
Option Explicit
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const SW_RESTORE = 9
Public RepQue() As Long
Public REPRS As ADODB.Recordset
Public Type RepRec
    EDI_Report_Index As Long
    EDI_Provider_Org As String
    EDI_Loc_Nat_Code_To As String
    EDI_Clin_Nat_Code_To As String
End Type
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
    Dim hWnd As Long
    Dim lRetVal As Long
    Sname = Name.Caption
    If App.PrevInstance Then
        App.Title = "NewCopy"
        Name.Caption = "NewCopy"
        MsgBox "Another Version of '" + Sname + "' is already loaded", vbOKOnly
        hWnd = FindWindow(vbNullString, Sname)
        If hWnd <> 0 Then
            lRetVal = ShowWindow(hWnd, SW_RESTORE)
            lRetVal = SetForegroundWindow(hWnd)
            End If
        End
        End If
End Sub
'Write to a ini File
Function Write_Ini_Var(section, varname As String, Value As String, fileName) As Boolean
Dim Success, Message, Response
    fileName = UCase$(fileName)
        
Writeagain:
  
    Success = WritePrivateProfileString(section, varname, Value, fileName)

  If Success Then
    Write_Ini_Var = True
  Else
    Message = "Unable to Write Profile"
    Response = MsgBox(Message, 277, "Error")
   If Response = 4 Then GoTo Writeagain
    Write_Ini_Var = False
  End If

End Function
'Read the INI file
Function Read_Ini_Var(section, varname, fileName)
 
 'this function gets a profile varable from either the
 '"WIN.INI" or a private "*.INI" file if the profile
 'does not exist an error message is displayed and the
 'function returns a null string

Dim Message, Success
Dim Sname$, Vname$, Fname$, ret$, kname$
Dim Response

    ret$ = String$(255, 0)
    fileName = UCase$(fileName)
    Sname$ = section
    kname$ = varname
    Success = GetPrivateProfileString(Sname$, kname$, "", ret$, Len(ret$), fileName)
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

Public Function ValidateFilepath(fPath As String) As String
   On Error GoTo procEH
'  Function takes the supplied file path and creates all folders necessary
'  from the root directory downwards, to make the path valid.
'  A false return value indicates the original path was invalid.
   Dim driveId As String
   Dim fldr As String
   Dim oldFldr As String
   
   If RunningInIDE Then
      driveId = fs.GetDriveName(fPath)
      If fs.DriveExists(driveId) = False Then
         If Len(driveId) > 3 Then
            fPath = "C:" & Mid(fPath, InStr(3, fPath, "\"))
         ElseIf UCase(driveId) <> "C:" Then
            fPath = "C:" & Mid(fPath, InStr(3, fPath, "\"))
         End If
      ElseIf UCase(driveId) <> "C:" Then
         fPath = "C:" & Mid(fPath, InStr(3, fPath, "\"))
      End If
   End If
   
   eClass.FurtherInfo = "Path = " & fPath
   If fs.FolderExists(fPath) = False And fPath <> "" Then
      fldr = fs.GetParentFolderName(fPath)
      eClass.FurtherInfo = "Validating path: " & fldr
      Do Until fs.FolderExists(fldr)
         Do Until fs.FolderExists(fldr) Or fldr = ""
            oldFldr = fldr
            fldr = fs.GetParentFolderName(fldr)
            If fldr = "" Then
               eClass.FurtherInfo = fPath
               Err.Raise 3465, "ICEMsg", "Invalid drive or UNC path"
            End If
         Loop
         eClass.FurtherInfo = "Creating Folder: " & oldFldr
         fs.CreateFolder (oldFldr)
         fldr = fs.GetParentFolderName(fPath)
      Loop
      eClass.FurtherInfo = "Creating Folder: " & fPath
      fs.CreateFolder fPath
   End If
   ValidateFilepath = fPath
   Exit Function

procEH:
   eClass.CurrentProcedure = "ICEmsg.StandardStuff.ValidateFilepath"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function RunningInIDE() As Boolean
   On Error Resume Next
   Debug.Print 1 / 0
   RunningInIDE = (Err.Number <> 0)
End Function
