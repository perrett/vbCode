Attribute VB_Name = "GlobalData"
Option Explicit

Public Type LogStruct
   DateTime As String
   Procedure As String
   Operation As String
   Description As String
   Number As String
   Source As String
   MsgData As String
   Terminator As String * 2
End Type

Public intBehaviour As Integer  ' Return without failing =  -1, Evaluate Error = 0, Report and stop execution =  1
Public blnLogErrors As Boolean
Public blnLogOverWrite As Boolean
Public intLogMode As Integer
Public errAction As Object
Public strLogSrc As String
Public errNo As Long
Public strDesc As String
Public strSrc As String
Public procName As String
Public extraInfo As String
Public statusFlag As Long
Public modeFlag As Long
Public errorEvaluated As Boolean
Public eData As errorData
Public logBuf As LogStruct
Public dbCon As ADODB.Connection
Public fs As New FileSystemObject
Public LogDirPath As String
Public LogFile As String

Public colErr As New Collection

Public Function ValidateFilepath(SuppliedPath As String) As String
'  Function takes the supplied file path and creates all folders necessary
'  from the root directory downwards, to make the path valid.
   Dim fldr As String
   Dim oldFldr As String
   Dim fPath As String
   
   fPath = SuppliedPath
   If fs.FolderExists(fPath) = False Then
      fldr = fs.GetParentFolderName(fPath)
         
      Do Until fs.FolderExists(fldr) Or fldr = ""
         Do Until fs.FolderExists(fldr)
            oldFldr = fldr
            fldr = fs.GetParentFolderName(fldr)
         Loop
         fs.CreateFolder (oldFldr)
         fldr = fs.GetParentFolderName(fPath)
      Loop
      
      If fldr = "" Then
         fPath = App.Path
      End If
      
      fs.CreateFolder fPath
   End If
   
   ValidateFilepath = fPath
End Function
