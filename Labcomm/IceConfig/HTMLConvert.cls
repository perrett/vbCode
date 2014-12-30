VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTMLConvert"
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

Public Sub DeleteTempFile()
   On Error Resume Next
   fs.DeleteFile preparedFile
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

Private Function PrepareLogHTML(logData As LogStruct) As String
   Dim htmlBuf As String
   Dim cellbuf As String
   Dim strTemp As String
   Dim strRed As String
   
   cellbuf = htmlData
   
   htmlBuf = Replace(cellbuf, "@data@", logData.DateTime) & vbCrLf
   
   If logData.Procedure = "IceMsg" Then
      htmlBuf = htmlBuf & Replace(cellbuf, "@data@", logData.MsgData) & vbCrLf
   Else
      strTemp = logData.MsgData & "<BR> " & _
                logData.Number & "<BR>" & _
                logData.Description & "<BR>" & _
                logData.Procedure & "<BR>" & _
                logData.Operation
      htmlBuf = htmlBuf & Replace(cellbuf, "@data@", strTemp) & vbCrLf

'      htmlBuf = htmlBuf & Replace(cellbuf, "@data@", logData.MsgData & "<BR>" & logData.Description & "<BR>" & logData.Procedure) & vbCrLf
'      htmlBuf = htmlBuf & Replace(cellbuf, "@data@", logData.Procedure) & vbCrLf
'      htmlBuf = htmlBuf & Replace(cellbuf, "@data@", logData.Number) & vbCrLf
'      htmlBuf = htmlBuf & Replace(cellbuf, "@data@", logData.Operation) & vbCrLf
'      htmlBuf = htmlBuf & Replace(cellbuf, "@data@", logData.MsgData) & vbCrLf
      htmlBuf = Replace(htmlBuf, "#000000", "#ff0000")
   End If
   PrepareLogHTML = htmlRowStart & htmlBuf & htmlRowEnd
End Function

Public Sub ReadLogFile(logFile As String)
   Dim logbuf As LogStruct
   Dim sBuf As New StringBuffer
   Dim fBuf As String
   Dim buf As String
   Dim outfile As String
      
   fBuf = String(8192, Chr(0))
   buf = ""
   sBuf.Clear
   sBuf.Append "<HTML><HEAD><STYLE>TABLE,BODY,SELECT{FONT-FAMILY:Tahoma;FONT-SIZE:10px;}" & _
               ".Times{FONT-FAMILY:Tahoma;FONT-SIZE:9px;}</STYLE></HEAD><BODY>"
               
   sBuf.Append "<TABLE Style=" & Chr(34) & "width=100%" & Chr(34) & ">" '& _
                      "nowrap=" & Chr(34) & "true" & Chr(34) & ">"

   Open logFile For Binary As #1
   Do Until EOF(1)
      Get #1, , logbuf
      sBuf.Append PrepareLogHTML(logbuf)
   Loop
   Close #1
   
   outfile = fs.BuildPath(App.Path, fs.GetTempName)
   Open outfile For Output As #1
   Print #1, sBuf.ActualValue & "</TABLE></BODY></HEAD></HTML>"
   Close #1
   preparedFile = outfile
End Sub

Private Sub ReadTemplate()
   Dim i As Integer
   Dim oText As TextStream
   Dim sBuffer As New StringBuffer
   
   'Open template to a string
   
   htmlHeader = ""
   htmlData = ""
   htmlTrailer = ""
      
   Set oText = fs.OpenTextFile(App.Path & "\" & "LogTemplate.htm", ForReading, False)
   For i = 0 To 5
      htmlHeader = htmlHeader & oText.ReadLine & vbCrLf
   Next i
   
   htmlRowStart = oText.ReadLine
   htmlData = oText.ReadLine
   htmlRowEnd = oText.ReadLine
   
   Do Until oText.AtEndOfStream
      htmlTrailer = htmlTrailer & oText.ReadLine & vbCrLf
   Loop
   
   oText.Close
End Sub

Private Sub Class_Terminate()
   DeleteTempFile
End Sub
