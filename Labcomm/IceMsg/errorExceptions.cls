VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "errorExceptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ENUM_ErrTraps
   APP_INVALID_DATABASE = 3263
   VB_PATH_NOT_FOUND = 76
   VB_NO_PATH = 423
   VB_FILE_NOT_FOUND = 53
   UDL_NOFILE = 432
End Enum

Private errPath As String

Friend Property Let ErrorFilePath(strNewValue As String)
   errPath = strNewValue
End Property

Public Function Evaluate(errNum As Long, Data As String, errStatus As Long)
   Dim fd As Folder
   Evaluate = errNum
   
   eClass.FurtherInfo = "Evaluating Error: " & errNum
   Select Case errNum
      Case VB_NO_PATH
         If Data <> "" Then
            Set fd = fs.CreateFolder(Data)
         End If
         Evaluate = 0
      
      Case VB_PATH_NOT_FOUND
         If Data <> "" Then
            Set fd = fs.CreateFolder(Data)
         End If
         Evaluate = 0
      
      Case VB_FILE_NOT_FOUND
'         If Data <> "" Then
'            Set fd = fs.CreateFolder(Data)
'         End If
         Evaluate = 0
              
      Case UDL_NOFILE
         MsgBox "Unable to connect to the database - please check the UDL file in the title bar", vbCritical, UDLPath & "ICE.UDL"
      
      Case 94
         Evaluate = 3263
      
      Case 3021   '  BOF or EOF - Database inconsistent
'         If errStatus = 1 Then   '  Recoverable error
'           Set the error to 'Invalid Database Object'
            Evaluate = 3263
'         End If
   End Select

End Function

Public Sub HandleFailure(errNo As Long, errDetails As String)
   Dim errFile As String
   
   Select Case errNo
'     Put specific failure data here
      Case 3465   '  Invalid UNC
         MsgBox errDetails, vbCritical, "Please check the specified UNC exists"
         
      Case 3234   '  Database timeout
         SendMessage App.EXEName, "Timeout or connection lost processing report"
      
      Case 3263   '  Database inconsistencies
         If tLevel = 1 Then
            iceCon.RollbackTrans
            tLevel = 0
         End If
 '        errFile = RemoveReplistEntry(errDetails)

'        Notify the control application
         SendMessage App.ProductName, "Database inconsistencies found"
         SendMessage App.EXEName, "Report index " & msgControl.CurrentReport & "- entry preserved in " & errFile
      
      Case Else
         If blnErrHalt Then
            MsgBox "Details: " & errDetails, vbCritical, "Unrecoverable Error: " & errNo
         Else
            SendMessage App.EXEName, "Application terminating - unrecoverable error: " & errDetails
         End If
         
   End Select
End Sub
