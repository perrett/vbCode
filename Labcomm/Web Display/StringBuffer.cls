VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STRING_BUFFER_SIZE As Long = 32768

Private mlBufPos As Long
Private mlBufSize As Long
Private msBuf As String

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    msBuf = vbNullString
End Sub

Public Property Get ActualValue() As String
   ActualValue = Left$(msBuf, mlBufPos - 1)
End Property

Public Sub Append(ByVal TheText As String)
    
    On Error GoTo ErrorHandler
        
    Dim BytesToAppend As Long
    Dim TargetSize As Long
    
    BytesToAppend = Len(TheText)
    TargetSize = mlBufPos + BytesToAppend
    If TargetSize < mlBufSize Then
        Mid$(msBuf, mlBufPos, BytesToAppend) = TheText
        mlBufPos = TargetSize
    Else
        msBuf = msBuf & Space(STRING_BUFFER_SIZE)
        mlBufSize = mlBufSize + STRING_BUFFER_SIZE
        Append TheText
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError Err.Number, Err.Description, Err.Source, "StringBuffer_Append"
End Sub

Public Property Get Length() As Long
   Length = mlBufPos - 1
End Property


Public Property Get value() As String
   If mlBufPos > 100 Then
      value = Replace(Left$(msBuf, mlBufPos - 1), "'", vbCrLf)
   Else
      value = Left$(msBuf, mlBufPos - 1)
   End If
End Property

Public Property Let value(newValue As String)
    mlBufPos = 1
    Append newValue
End Property

Public Sub Clear()
    On Error GoTo ErrorHandler
       
    mlBufSize = STRING_BUFFER_SIZE
    msBuf = Space$(mlBufSize)
    mlBufPos = 1
       
    Exit Sub
    
ErrorHandler:
    HandleError Err.Number, Err.Description, Err.Source, "StringBuffer_Append"
End Sub

Private Sub HandleError(lErrNumber As Long, sErrDesc As String, sErrSource As String, sOccuredIn As String)
    App.LogEvent "Error in " & sOccuredIn & " Number = " & lErrNumber & _
        vbCrLf & " Description = " & sErrDesc & vbCrLf & " Source= " & _
        sErrSource
    Err.Raise vbObjectError + lErrNumber, sErrSource & " " & sOccuredIn, sErrDesc
End Sub
