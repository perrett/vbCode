VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBBlob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private dbCon As ADODB.Connection
Private fs As New FileSystemObject
Private tFile As String

Public Property Let DBConnection(objConn As ADODB.Connection)
   Set dbCon = objConn
End Property

Public Function Insert(TemplateFile As String, TemplateType As enum_TEMPLATETYPES)
   Dim RS As New ADODB.Recordset
   Dim TemplateName As String
   Dim status As Integer
   Dim tmpId As Long
   Dim fBuf As String
   Dim buf As String
   Dim bufLen As Long
   Dim pos As Long
   
   TemplateName = fs.GetBaseName(TemplateFile)
   status = 0
   strSQL = "SELECT Id " & _
            "FROM HTML_Templates " & _
            "WHERE Name = '" & TemplateName & "' " & _
               "AND Type = " & TemplateType
   RS.Open strSQL, dbCon, adOpenKeyset, adLockReadOnly
   
   status = CInt(RS.EOF)
   
   If RS.EOF = False Then
      status = MsgBox(TemplateName & " already exists - overwrite original?", vbYesNo, "Existing template")
   End If
      
   RS.Close
      
   If status <> 0 Then
      Open TemplateFile For Binary As #1
      fBuf = String(8192, Chr(0))
      Do Until EOF(1)
         Get #1, , fBuf
         buf = buf & fBuf
      Loop
      buf = Left(buf, LOF(1))
      Close #1
      
      RS.Open "HTML_Templates", dbCon, adOpenKeyset, adLockOptimistic, adCmdTable
      RS.Filter = "Name='" & TemplateName & "' AND Type='" & TemplateType & "'"
      
      If status > 0 Then
         RS.Delete
         RS.Update
      End If
      
      RS.AddNew
      
      RS!Name = TemplateName
      RS!Type = TemplateType
      
      If Len(buf) > 1000 Then
         bufLen = 1000
      Else
         bufLen = Len(buf)
      End If
      
      pos = 1
      Do While pos < Len(buf)
         RS!data.AppendChunk Mid(buf, pos, bufLen)
         pos = pos + bufLen
      Loop
      RS!data.AppendChunk Mid(buf, pos)
'      If pos > Len(buf) Then
'         pos = pos - bufLen
'         bufLen = Len(buf) - pos + 1
'         RS!Data.AppendChunk Mid(buf, pos)
'      End If
      
      RS.Update
      RS.Close
   End If
   Set RS = Nothing
End Function

Public Function Read(TemplateID As Variant, _
                     TemplateType As enum_TEMPLATETYPES) As Variant
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim tLen As Long
   Dim pos As Long
   Dim buf As String
   Dim bufLen As Long
   
   If IsNumeric(TemplateID) Then
      strSQL = "SELECT Data " & _
               "FROM HTML_Templates " & _
               "WHERE ID = '" & TemplateID & "' " & _
                  "AND Type = " & TemplateType
   Else
      strSQL = "SELECT Data " & _
               "FROM HTML_Templates " & _
               "WHERE Name = '" & TemplateID & "' " & _
                  "AND Type = " & TemplateType
   End If
   
   RS.Open strSQL, dbCon, adOpenStatic, adLockReadOnly
   tLen = RS(0).ActualSize
   pos = 0
   
   If tLen > 4000 Then
      bufLen = 4000
   Else
      bufLen = tLen
   End If
   
   Do Until pos >= tLen
      buf = buf & RS(0).GetChunk(bufLen)
      pos = pos + bufLen
      If bufLen > tLen - pos Then
         bufLen = tLen - pos
      End If
   Loop
   RS.Close
   Read = buf
End Function

Private Sub Class_Terminate()
   On Error Resume Next
   fs.DeleteFile tFile, True
   Set dbCon = Nothing
End Sub
