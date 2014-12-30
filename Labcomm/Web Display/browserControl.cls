VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "browserControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fs As New FileSystemObject
Private wb As wbCtrl

Private qt As String

Private Sub Class_Initialize()
   qt = Chr(34)
End Sub

Public Property Let WebControlId(objNewValue As wbCtrl)
   Set wb = objNewValue
   wb.ObjectName = "objFS"
End Property

Public Sub onDataAvailable(fDir As String, fName As String)
   Dim ev As MSHTML.IHTMLEventObj
'   Dim dNode As MSHTML.IHTMLDOMNode
'   Dim fName As String
'   Dim fDir As String
   
   DoEvents
'   fDir = hDoc.Title
'   Set ev = hDoc.parentWindow.event
'   Set dNode = ev.srcElement
'   fName = ev.srcElement.Children(2).Title 'Mid(ev.srcElement.Title, 3)
   
'   wb.ClearDocumentBody True
   If fs.FileExists(fDir) Then
'     We are viewing a file so show files in the parent directory
   Else
      If fs.GetBaseName(fs.GetParentFolderName(fDir)) = fName Then
'        The 'Up' button was clicked
         ShowSubfolders fs.GetParentFolderName(fDir), False
      ElseIf fs.FolderExists(fs.BuildPath(fDir, fName)) Then
'        A directory was clicked so show the files
         fDir = fs.BuildPath(fDir, fName)
         ShowFiles fDir, True
      Else
'        Display the file
'         hDoc.Close
'         ev.cancelBubble = True
'         ev.returnValue = False
'         Set ev = Nothing
         ShowFileDetails fs.BuildPath(fDir, fName)
      End If
   End If
   wb.BrowserHTML
End Sub

'Public Sub PrepareDirBrowser()
'   wb.AddObject "CLSID:CDC6FB84-0E15-4A8F-993B-8524A9EE452E", 107, 141
'   wb.AddStyleSheet "folderview.css", "SS_Fold"
'End Sub

Public Sub ShowFileDetails(FileName As String)
   Dim buf As String
   Dim tFile As String
   Dim fExt As String
   
   fExt = fs.GetExtensionName(FileName)
   
   If UCase(fExt) = "XML" Or Left(UCase(fExt), 3) = "htm" Then
      wb.NavigateTo FileName, True
   Else
      tFile = fs.BuildPath(App.Path, "dataViewer.html")
      buf = FileToHTML(FileName)
      Open tFile For Output As #1
      Print #1, buf
      Close #1

      wb.NavigateTo tFile, True
   End If
End Sub

Public Sub ShowFiles(DirName As String, _
                     Optional BackButton As Boolean = False)
   Dim buf As String
   Dim tFile As String
   
   tFile = fs.BuildPath(App.Path, "dataViewer.html")
   buf = Files(DirName, False)
   Open tFile For Output As #1
   Print #1, buf
   Close #1
   
   wb.NavigateTo tFile, True
'   PrepareDirBrowser
'   wb.AddScripts "fileview.js;viewcontrol.js", , True
'   wb.FileTitle = DirName
'
'   If BackButton Then
'      wb.SetDataEventParams fs.GetBaseName(fs.GetParentFolderName(DirName))
'   End If
'
'   wb.FireDataEvent
'
'   PrepareDirBrowser
'   wb.AddScripts "fileview.js;viewcontrol.js", , True
'   wb.NavigateTo "C:\ICE\LABCOMM\HISTORY\MSGOUT\030630\SX_XML_1_XML v2.4_1266_17-26-18.xml", True
End Sub

Public Sub ShowSubfolders(DirName As String, _
                          Optional BackButton As Boolean = False)
   Dim buf As String
   Dim tFile As String
   
   tFile = fs.BuildPath(App.Path, "dataViewer.html")
   buf = SubFolders(DirName, False)
   Open tFile For Output As #1
   Print #1, buf
   Close #1
   
   wb.NavigateTo tFile, True
'   wb.FileTitle = DirName
'   PrepareDirBrowser
'   wb.AddScripts "folderview.js;viewcontrol.js", , True
'
'   If BackButton Then
'      wb.SetDataEventParams fs.GetBaseName(fs.GetParentFolderName(DirName))
'   End If
'
'   wb.FireDataEvent
'
'   PrepareDirBrowser
'   wb.AddScripts "folderview.js;viewcontrol.js", , True
''   wb.BrowserHTML
End Sub

Public Function FileToHTML(FileName As String) As String
   Dim strHTML As String
   Dim sBuf As New StringBuffer
   Dim fBuf As String
   Dim fExt As String
   Dim bodyHdr As String
   Dim qt As String
   
'   strHTML = htmlHeader(FileName, "FileView")
   strHTML = strHTML & "<table class=edifact>" & vbCrLf & " <tr>" & vbCrLf & "  <td nowrap>"
      
   Open FileName For Input As #1
   While Not EOF(1)
      Line Input #1, fBuf
      sBuf.Append fBuf & "</td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td nowrap>"
   Wend
   Close #1
      
   fBuf = Replace(sBuf.ActualValue, "?'", "#apos#")
   fBuf = Replace(fBuf, "'", "'</td>" & vbCrLf & " </tr>" & vbCrLf & " <tr>" & vbCrLf & "  <td>")
   fBuf = Replace(fBuf, "#apos#", "'")
   
   fBuf = fBuf & "</td>" & vbCrLf & " </tr>" & vbCrLf & "</table>"
   strHTML = strHTML & fBuf & htmlTrailer
   FileToHTML = strHTML
End Function

Public Function Files(DirName As String, _
                      Optional showBack As Boolean = False) As String
'   Dim strHTML As String
   Dim sBuf As New StringBuffer
   Dim imgSrc As String
   Dim fileImg As String
   Dim rowSrc As String
   Dim colId As Integer
   Dim fl As Files
   Dim fileId As File
   Dim i As Integer
   
'   sBuf.Append htmlHeader(DirName, "folderview", "viewcontrol")
   
   sBuf.Append "<table>" & vbCrLf
   sBuf.Append "<tr>" & _
             "<td class=" & qt & "foldertitle" & qt & " colspan=" & qt & "6" & qt & ">" & DirName & "</td>" & _
             "</tr>"

   If showBack Then
      imgSrc = "<img action=" & qt & "over" & qt & " height=" & qt & "32" & qt & " width=" & qt & "32" & qt & _
               "alt=" & qt & "Click to open" & qt & " border=" & qt & "0" & qt & _
               " src=" & qt & "C:\\ICE\\LABCOMM\\Icons\\back.ico" & qt & ">"
      rowSrc = " <td onclick=" & qt & "handleClick" & qt & " class=" & qt & "folder" & qt & ">" & _
             imgSrc & "<br><span action=" & qt & "over" & qt & " class=" & qt & "fixwidth" & qt & ">Up" & _
             "</span></td>"
      colId = colId + 1
   End If

   imgSrc = "<img action=" & qt & "over" & qt & " height=" & qt & "16" & qt & " width=" & qt & "16" & qt & _
            "alt=" & qt & "Click to open" & qt & " border=" & qt & "0" & qt & " src=" & qt & _
            "C:\\ICE\\LABCOMM\\Icons\\" 'folder.ico" & qt & ">"
   
   Set fl = fs.GetFolder(DirName).Files
   
   For Each fileId In fl
      
      Select Case UCase(fs.GetExtensionName(fileId.Name))
         Case "XMS"
            fileImg = imgSrc & "xms.ico" & qt & ">"
            
         Case "XEN"
            fileImg = imgSrc & "xen.ico" & qt & ">"
            
         Case "XML"
            fileImg = imgSrc & "html.ico" & qt & ">"
                     
         Case "WNG"
            fileImg = imgSrc & "warning.ico" & qt & ">"
         
         Case "ERR"
            fileImg = imgSrc & "error.ico" & qt & ">"
         
         Case Else
            fileImg = imgSrc & "unknown.ico" & qt & ">"
         
      End Select
      
      rowSrc = rowSrc & "   <td onclick=" & qt & "handleClick" & qt & " class=" & qt & "folder" & qt & ">" & _
               fileImg & "<span action=" & qt & "over" & qt & _
               " title=" & qt & fileId.Name & qt & ">" & Replace(fileId.Name, "-", "_") & "</span></td>"
      colId = colId + 1

      If colId = 3 Then
         sBuf.Append "<tr>" & rowSrc & "</tr>" & vbCrLf
         rowSrc = ""
         colId = 0
      End If
   Next
   
   If colId > 0 Then
      sBuf.Append "<tr>" & rowSrc & "</tr>" & vbCrLf
   End If
   
   sBuf.Append "</table>" & vbCrLf & htmlTrailer
   
   Files = sBuf.ActualValue
End Function

Public Function SubFolders(DirName As String, _
                           Optional showBack As Boolean = False) As String
   Dim strHTML As String
   Dim imgSrc As String
   Dim rowSrc As String
   Dim colId As Integer
   Dim fld As Folders
   Dim i As Integer
   
   strHTML = htmlHeader(DirName, fs.BuildPath(App.Path, "folderview.css"), fs.BuildPath(App.Path, "viewcontrol.js"))
   
   strHTML = strHTML & "<table>" & vbCrLf
   strHTML = strHTML & "<tr>" & _
             "<td class=" & qt & "foldertitle" & qt & " colspan=" & qt & "6" & qt & ">" & DirName & "</td>" & _
             "</tr>"

   If showBack Then
      imgSrc = "<img action=" & qt & "over" & qt & " height=" & qt & "32" & qt & " width=" & qt & "32" & qt & _
               "alt=" & qt & "Click to open" & qt & " border=" & qt & "0" & qt & _
               " src=" & qt & "C:\\ICE\\LABCOMM\\Icons\\back.ico" & qt & ">"
      rowSrc = " <td onclick=" & qt & "handleClick" & qt & " class=" & qt & "folder" & qt & ">" & _
             imgSrc & "<br><span action=" & qt & "over" & qt & " class=" & qt & "fixwidth" & qt & ">Up" & _
             "</span></td>"
      colId = colId + 1
   End If

   imgSrc = "<img action=" & qt & "over" & qt & " height=" & qt & "32" & qt & " width=" & qt & "32" & qt & _
            "alt=" & qt & "Click to open" & qt & " border=" & qt & "0" & qt & " src=" & qt & _
            "C:\\ICE\\LABCOMM\\Icons\\folder.ico" & qt & ">"
   
   Set fld = fs.GetFolder(DirName).SubFolders
   
   For i = 0 To fld.Count
      rowSrc = rowSrc & "   <td onclick=" & qt & "handleClick" & qt & " class=" & qt & "folder" & qt & ">" & _
               imgSrc & "<br><span action=" & qt & "over" & qt & " class=" & qt & "fixwidth" & qt & _
               " title=" & qt & fld.Item(i).Name & qt & ">" & fld.Item(i).Name & "</span></td>"
      colId = colId + 1

      If colId = 6 Then
         strHTML = strHTML & "<tr>" & rowSrc & "</tr>" & vbCrLf
         rowSrc = ""
         colId = 0
      End If
   Next i
   
   If colId > 0 Then
      strHTML = strHTML & "<tr>" & rowSrc & "</tr>" & vbCrLf
   End If
   
   strHTML = strHTML & "</table>" & vbCrLf & htmlTrailer
   
   SubFolders = strHTML
End Function

Private Function htmlHeader(Optional Title As String = "", _
                            Optional StyleSheet As String = "", _
                            Optional Script As String = "") As String
   Dim strHTML As String
   Dim dHTML As String
   Dim buf As String
   
   strHTML = "<head>" & vbCrLf
   
   If Title <> "" Then
      strHTML = strHTML & "<title>" & Title & "</title>"
   End If
   
   If StyleSheet <> "" Then
      strHTML = strHTML & "<style>" & vbCrLf
      
      If styleDir = "" Then
         strHTML = strHTML & blob.Read(StyleSheet, TT_Style)
      Else
         dHTML = fs.BuildPath(styleDir, StyleSheet & ".css")
      
         If fs.FileExists(dHTML) Then
            Open dHTML For Input As #1
            Do Until EOF(1)
               Input #1, buf
               strHTML = strHTML & buf & vbCrLf
            Loop
            Close #1
         End If
      End If
      
      strHTML = strHTML & "</style>" & vbCrLf
   End If
   
   If Script <> "" Then
      strHTML = strHTML & "<script>" & vbCrLf
      
      If scriptDir = "" Then
         strHTML = strHTML & blob.Read(Script, TT_Script)
      Else
         dHTML = fs.BuildPath(scriptDir, Script & ".js")
         If fs.FileExists(dHTML) Then
            Open dHTML For Input As #1
            Do Until EOF(1)
               Input #1, buf
               strHTML = strHTML & buf & vbCrLf
            Loop
            Close #1
         End If
      End If
      strHTML = strHTML & "</script>" & vbCrLf
   End If
   
   strHTML = strHTML & "</head>" & vbCrLf & "<body>"
   htmlHeader = strHTML
End Function

Private Function htmlTrailer() As String
   Dim strHTML As String
   
   htmlTrailer = "</body>" & vbCrLf & "</html>"
End Function
