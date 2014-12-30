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
   VB_PATH_NOT_FOUND = 76
   VB_NO_PATH = 423
   VB_FILE_NOT_FOUND = 53
   CMD_UDL_FILE = -2147467259
   CMD_AUTOMATION = -2147217873
   SQL_AUTOMATION = -2147217900
   APP_OVERFLOW = 3016
   SQL_DUP_PRIMARY_KEY = 2627
   SQL_DUP_INDEX = 2601
   SQL_NO_NULLS = 515
   UDL_NO_FILE = 432
End Enum

Private blnUpdateOk As Boolean

Private Sub Class_Initialize()
   blnUpdateOk = True
End Sub

Public Property Get UpdateDb() As Boolean
   UpdateDb = blnUpdateOk
   blnUpdateOk = True
End Property

Public Function Evaluate(errNum As Long, _
                         data As String, _
                         Optional errStatus As Long = 0) As Long
   Dim fs As New FileSystemObject
   Dim fd As Folder
   Dim frm As Form
   Dim strArray() As String
   Dim adoErr As ADODB.Error
   Dim errId As Long
   Evaluate = errNum
   
   For Each frm In Forms
      If frm.Name <> "frmMain" Then
         Unload frm
      End If
   Next
   
   Select Case errNum
      Case 16001
         MsgBox Err.Description, vbInformation, "IceConfig Validation"
         Evaluate = 0
         
      Case VB_NO_PATH
         If data <> "" Then
            ValidateFilepath data
'            Set fd = fs.CreateFolder(Data)
         End If
         Evaluate = 0
         
      Case VB_FILE_NOT_FOUND
         MsgBox "File: " & eClass.FurtherInfo & vbCrLf & vbCrLf & "no longer held in file archive.", vbExclamation, "ICEConfig - Non critical eror"
         Evaluate = 0
      
      Case VB_PATH_NOT_FOUND
         MsgBox "File: " & eClass.FurtherInfo & vbCrLf & vbCrLf & "no longer held in file archive.", vbExclamation, "ICEConfig - Non critical eror"
         Evaluate = 0
         
'      Case SQL_AUTOMATION
'         For Each adoErr In ICECon.Errors
'            If adoErr.SQLState = "42000" Then
'               eClass.Add adoErr.SQLState, "Syntax or Connection error", "SQL"
'            Else
'               eClass.Add adoErr.SQLState, "Unknown, please look up", "SQL"
'            End If
'         Next
         
      Case CMD_UDL_FILE
         Evaluate = 432
      
      Case CMD_AUTOMATION
         blnUpdateOk = False
         If IsNumeric(data) Then
            errId = data
         Else
            If iceCon.Errors.Count > 0 Then
               errId = iceCon.Errors(0).NativeError
            End If
         End If
            
'        The command object has reported an error. The SQL error code is passed via the data parameter.
         Select Case errId
            Case SQL_DUP_PRIMARY_KEY
               MsgBox "Unable to update the database, probably because this would create a duplicate entry", vbExclamation, "Database Update"
               Evaluate = 0
               
            Case SQL_DUP_INDEX
               MsgBox "Unable to update the database, probably because this would create a duplicate entry", vbExclamation, "Database Update"
               Evaluate = 0
            
            Case SQL_NO_NULLS
               MsgBox "Unable to update the database. One or more of the required fields have not been completed", vbExclamation, "Database Update"
               Evaluate = 0
         
         End Select
   
      Case APP_OVERFLOW
         strArray = Split(data, ":")
         MsgBox "The field " & strArray(0) & " may only be " & strArray(1) & " characters in length", vbExclamation, "ICE Config"
         Evaluate = 0
         
      Case 3263
         MsgBox "The database is inaccessible. Please check that an instance of Iceconfig is not already running", _
                vbExclamation, "Database unavailable"
         End
      
   End Select

End Function
