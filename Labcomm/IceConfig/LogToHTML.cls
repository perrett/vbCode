VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogToHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type LogStruct
   DateTime As String
   Procedure As String
   Operation As String
   Description As String
   Number As String
   Source As String
   MsgData As String
   Terminator As String * 2
End Type

Private Type logRead
   readBuffer(1000) As LogStruct
End Type

Private fs As New FileSystemObject
Private preparedFile As String
Private htmlHeader As String
Private htmlTrailer As String
Private htmlRowStart As String
Private htmlRowEnd As String
Private htmlData As String

Private Sub Class_Initialize()
   ReadTemplate
End Sub

Public Property Get HTMLfile() As String
   HTMLfile = preparedFile
End Property

Public Sub PrepareOutput(Filename As String)
   Dim blnHTML As Boolean
   Dim blnXML As Boolean
   Dim fBuf As String
   Dim buf As String
   Dim outfile As String
   Dim htmlHeader As String
   Dim htmlTrailer As String
   
   fBuf = String(8192, Chr(0))
   buf = ""
   Open Filename For Binary As #1
   Do Until EOF(1)
      Get #1, , fBuf
      buf = buf & fBuf
   Loop
   buf = Left(buf, LOF(1))
   Close #1
   
   blnHTML = (InStr(1, Left(buf, 40), "HTML", vbTextCompare)) > 0
   blnXML = (InStr(1, Left(buf, 40), "XML", vbTextCompare)) > 0
   
   If (blnHTML Or blnXML) = False Then
      If Left(buf, 12) = "UNA:+.? 'UNB" Then
         buf = Replace(buf, "'", "</TR><TR>")
      End If
      
      htmlHeader = "<HTML>" & vbCrLf & "<BODY id=" & Chr(34) & "FileData" & Chr(34) & _
                    "bgcolor=" & Chr(34) & "#FFFFCC" & Chr(34) & _
                    "style=" & Chr(34) & "font-family: Courier New; font-size: 8pt" & Chr(34) & ">" & vbCrLf & _
                    "<div style=" & Chr(34) & "WIDTH: 100%" & Chr(34) & " id=" & Chr(34) & "divReport" & Chr(34) & ">" & _
                    "<TABLE>" & vbCrLf & "<TD nowrap=" & Chr(34) & "true" & Chr(34) & ">" & vbCrLf
      htmlTrailer = "</TD></TABLE></Div></BODY></HTML>"
      buf = htmlHeader & buf & htmlTrailer
   End If
   
   outfile = fs.BuildPath(App.Path, fs.GetTempName)
   Open outfile For Output As #1
   Print #1, buf
   Close #1
   
   preparedFile = outfile
End Sub

Private Function PrepareLogHTML(logData As logRead) As String
   Dim i As Integer
   Dim htmlBuf As String
   Dim cellbuf As String
   Dim sBuf As New StringBuffer
   Dim strTemp As String
   Dim strRed As String
   
   sBuf.Clear
   For i = 0 To UBound(logData.readBuffer)
      If logData.readBuffer(i).Procedure = "" Then
         Exit For
      End If
      
      With logData.readBuffer(i)
         cellbuf = htmlData
         
         strTemp = .DateTime & ": "
'         htmlBuf = Replace(cellbuf, "#data#", .DateTime) & vbCrLf
         
         If .Procedure = "IceMsg" Then
            htmlBuf = Replace(cellbuf, "#data#", strTemp & .MsgData) & vbCrLf
            If .MsgData = "Processing acknowledgements" Or .MsgData = "Process Complete" Then
               htmlBuf = Replace(htmlBuf, "#ClassId#", "runstart")
            Else
               htmlBuf = Replace(htmlBuf, "#ClassId#", "log")
            End If
         Else
            strTemp = .MsgData & "<BR> " & _
                      .Number & "<BR>" & _
                      .Description & "<BR>" & _
                      .Procedure & "<BR>" & _
                      .Operation
            htmlBuf = htmlBuf & Replace(cellbuf, "#data#", strTemp) & vbCrLf
            htmlBuf = Replace(htmlBuf, "#ClassId#", "error")
         End If
      End With
      sBuf.Append htmlBuf
   Next i
   
   PrepareLogHTML = sBuf.ActualValue
End Function

Public Function ReadLogFile(logFile As String) As String
   On Error GoTo procEH
   Dim logbuf As logRead
   Dim sBuf As New StringBuffer
   Dim fBuf As String
   Dim buf As String
   Dim outfile As String
      
   fBuf = String(8192, Chr(0))
   buf = ""
   sBuf.Clear
'   sBuf.Append "<HTML><HEAD><STYLE>TABLE,BODY,SELECT{FONT-FAMILY:Arial;FONT-SIZE:12px;}" & _
'               ".Times{FONT-FAMILY:Arial;FONT-SIZE:12px;}</STYLE></HEAD><BODY>"
'
'   sBuf.Append "<TABLE Style=" & Chr(34) & "width=100%" & Chr(34) & ">" '& _
'                      "nowrap=" & Chr(34) & "true" & Chr(34) & ">"

   Open logFile For Binary As #1
   Do Until EOF(1)
      Get #1, , logbuf
      sBuf.Append PrepareLogHTML(logbuf)
   Loop
   Close #1
   
   ReadLogFile = htmlHeader & sBuf.ActualValue & htmlTrailer
   
WriteFile:
   outfile = fs.BuildPath(App.Path, fs.GetTempName)
   Open outfile For Output As #1
   Print #1, htmlHeader & sBuf.ActualValue & htmlTrailer
   Close #1
   ReadLogFile = outfile
   Exit Function
   
procEH:
   sBuf.Append "Error " & Err.Number & " (" & Err.Description & ") while reading log file"
   GoTo WriteFile
End Function

Private Sub ReadTemplate()
   Dim i As Integer
   Dim sBuffer As New StringBuffer
'   Dim blob As New clsDBBlob
   Dim buf As String
   Dim strArray() As String
   'Open template to a string
   
'   blob.dbConnection = iceCon
   htmlHeader = ""
   htmlData = ""
   htmlTrailer = ""
'   buf = blob.Read("log", TT_HTML)
'   strArray = Split(buf, vbCrLf)
   
'   htmlHeader = strArray(0) & vbCrLf & strArray(1) & vbCrLf
'   htmlData = strArray(2) & vbCrLf
'   htmlTrailer = strArray(3) & vbCrLf & strArray(4) & vbCrLf
End Sub

Private Sub Class_Terminate()
'   fs.DeleteFile preparedFile
End Sub
