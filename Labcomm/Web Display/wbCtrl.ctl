VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl wbCtrl 
   BackColor       =   &H8000000A&
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   ScaleHeight     =   6390
   ScaleWidth      =   7965
   Begin VB.Frame fraInfo 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   870
      Left            =   90
      TabIndex        =   9
      Top             =   60
      Width           =   7845
      Begin VB.Timer timerDA 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   300
         Top             =   135
      End
      Begin VB.OptionButton optShowFile 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show EDI Output"
         Height          =   225
         Index           =   1
         Left            =   5880
         TabIndex        =   13
         Top             =   450
         Width           =   1725
      End
      Begin VB.OptionButton optShowFile 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show Lab Input"
         Height          =   225
         Index           =   0
         Left            =   5880
         TabIndex        =   12
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox txtDesc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Description"
         Top             =   420
         Width           =   4300
      End
      Begin VB.TextBox txtTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Title"
         Top             =   0
         Width           =   4300
      End
   End
   Begin VB.Frame fraNav 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Caption         =   "Navigation"
      Height          =   870
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   7845
      Begin VB.CommandButton cmdHome 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1290
         Picture         =   "wbCtrl.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Return to the navigation start point"
         Top             =   255
         Width           =   615
      End
      Begin VB.CommandButton cmdUp 
         BackColor       =   &H00C00000&
         Height          =   645
         Left            =   3637
         Picture         =   "wbCtrl.ctx":0C42
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Back to the last view"
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Height          =   645
         Left            =   5985
         MaskColor       =   &H00C00000&
         Picture         =   "wbCtrl.ctx":1884
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Open view in a new Internet Explorer Window"
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3645
         TabIndex        =   7
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
   End
   Begin SHDocVwCtl.WebBrowser wbFolder 
      Height          =   5175
      Left            =   60
      TabIndex        =   4
      ToolTipText     =   "Double click item to navigate or view"
      Top             =   930
      Width           =   7845
      ExtentX         =   13838
      ExtentY         =   9128
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser wbDirect 
      Height          =   6435
      Left            =   15
      TabIndex        =   14
      Top             =   -30
      Width           =   7995
      ExtentX         =   14102
      ExtentY         =   11351
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblLoc 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   290
      Left            =   45
      TabIndex        =   5
      ToolTipText     =   "The current active file or folder"
      Top             =   6100
      Width           =   7845
   End
End
Attribute VB_Name = "wbCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DllVersionInfo
 cbSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformID As Long
End Type

Private Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Enum enum_TEMPLATETYPES
   TT_HTML = 1
   TT_Script = 2
   TT_Style = 3
End Enum

Private blob As New clsDBBlob

Private reportURL As String
Private navBarURL As String
Private logURL As String
Private remoteClass As Object

Private strArray() As String
Private filtArr() As String
Private testStr As String
Private homeDir As String
Private blnDefaultVerb As Boolean
Private regEdit As New clsRegEdit
Private expView(2) As String
Private bgCol As String
Private cObj As Object
Private evObj As Object
Private URLHistory() As String
Private hCount As Integer
Private wbHTMLStart As String
Private wbHTMLEnd As String
Private wbScript As String
Private wbStyle As String
Private wbStyleName As String
Private wbTitle As String

Private objName As String

Private WithEvents hDoc As MSHTMLCtl.HTMLDocument
Attribute hDoc.VB_VarHelpID = -1
'Private wbEv As New wbEvControl

Private scriptNode As MSHTML.HTMLScriptElement
'Private ev As IHTMLEventObj
'Private tr As MSHTMLCtl.IHTMLTxtRange
Private viewClass As Object
Private fs As New FileSystemObject
Private eclass As New AHSLErrorLog.errorControl
Private html As New browserControl
Private tFile As String
Private styleID As String
Private colParams As New Collection

Public Sub AddObject(clsId As String, _
                     clsHeight As String, _
                     clsWidth As String)
   Dim hObj As MSHTML.HTMLObjectElement
   Dim dNode As MSHTML.IHTMLElementCollection
   
   Set hObj = hDoc.getElementById(objName)
   If hObj Is Nothing Then
      Set hObj = hDoc.createElement("object")
      Set dNode = hDoc.getElementsByTagName("head")
      dNode(0).insertBefore hObj, dNode(0).firstChild
      hObj.Id = objName
      hObj.Width = clsWidth
      hObj.Height = clsHeight
      hObj.Attributes("classid").Value = clsId
   End If
End Sub
'
'Public Sub AddEventParam(Name As String, _
'                         value As String)
'   colParams.Add value, Name
'End Sub
'
'Public Sub ClearDocumentBody(Optional allTags As Boolean = False)
'   Dim scriptNode As MSHTML.IHTMLScriptElement
'   Dim styleNode As MSHTML.IHTMLStyleElement
'   Dim dNode As MSHTML.IHTMLDOMNode
'   Dim ss As MSHTML.IHTMLStyleSheet
'   Dim blnPresent As Boolean
'   Dim ssCol As MSHTML.IHTMLElementCollection
'   Dim buf As String
'   Dim cssBuf As String
'   Dim fileId As String
'   Dim strArray() As String
'   Dim i As Integer
'
'   If allTags Then
'      Set scriptNode = hDoc.getElementsByTagName("script")(0)
'      If Not scriptNode Is Nothing Then
'         scriptNode.Text = ""
'         Set dNode = scriptNode
'         dNode.removeNode True
'      End If
'
'      Set ssCol = hDoc.getElementsByTagName("style")
'      For i = 0 To ssCol.Length - 1
'         Set styleNode = ssCol(i)
'         styleNode.StyleSheet.cssText = ""
'         Set dNode = styleNode
'         dNode.removeNode True
'         Set dNode = Nothing
'      Next i
'
'      Set dNode = hDoc.getElementById(objName)
'      If Not dNode Is Nothing Then
'         dNode.removeNode True
'      End If
'   End If
'
'   hDoc.bgColor = "#FFFFFF"
'   hDoc.body.innerHTML = ""
'End Sub
'
'Public Sub ClearEventParams()
'   Set colParams = Nothing
'End Sub
'
'Public Sub AddScripts(scriptData As String, _
'                      Optional FileNotString As Boolean = True, _
'                      Optional ClearExisting As Boolean = False)
'   Dim scriptNode As MSHTML.IHTMLScriptElement
'   Dim dNode As MSHTML.IHTMLDOMNode
'   Dim sEl As MSHTML.IHTMLElementCollection
'   Dim scriptBuf As String
'   Dim fileId As String
'   Dim buf As String
'   Dim strArray() As String
'   Dim i As Integer
'
'   Set scriptNode = hDoc.getElementsByTagName("script")(0)
'   If scriptNode Is Nothing Then
'      Set scriptNode = hDoc.createElement("script")
'
'      Set dNode = hDoc.getElementsByTagName("head")(0)
'      dNode.insertBefore scriptNode, dNode.firstChild
'   End If
'
'   If ClearExisting Then
'      scriptNode.Text = ""
'   End If
'
'   If FileNotString Then
'      strArray = Split(scriptData, ";")
'      For i = 0 To UBound(strArray())
'         fileId = fs.BuildPath(scriptDir, strArray(i))
'         If fs.FileExists(fileId) Then
'            Open fileId For Input As #1
'            Do Until EOF(1)
'               Line Input #1, buf
'               scriptBuf = scriptBuf & buf
'            Loop
'            Close #1
'            scriptNode.Text = scriptNode.Text & scriptBuf & vbCrLf
'            scriptBuf = ""
'         End If
'      Next i
'   Else
'      scriptNode.Text = scriptData
'   End If
'End Sub
'
'Public Sub AddStyleSheet(cssData As String, _
'                         cssId As String, _
'                         Optional FileNotString As Boolean = True)
'   Dim ss As MSHTML.IHTMLStyleSheet
'   Dim blnPresent As Boolean
'   Dim ssCol As MSHTML.IHTMLElementCollection
'   Dim buf As String
'   Dim cssBuf As String
'   Dim fileId As String
'   Dim strArray() As String
'   Dim i As Integer
'
'   Set ssCol = hDoc.getElementsByTagName("style")
'   For i = 0 To ssCol.Length - 1
'      blnPresent = (blnPresent Or (ssCol(i).Title = cssId))
'   Next i
'
'   If blnPresent = False Then
'      Set ss = hDoc.createStyleSheet
'      ss.Title = cssId
'      If FileNotString Then
'         strArray = Split(cssData, ";")
'         For i = 0 To UBound(strArray())
'            fileId = fs.BuildPath(styleDir, strArray(i))
'            If fs.FileExists(fileId) Then
'               Open fileId For Input As #1
'               Do Until EOF(1)
'                  Line Input #1, buf
'                  cssBuf = cssBuf & buf
'               Loop
'               Close #1
'               ss.cssText = ss.cssText & cssBuf
'            End If
'         Next i
'      Else
'         ss.cssText = cssData
'      End If
'   End If
'End Sub

Public Sub BrowserHTML()
   Dim dNode As MSHTML.IHTMLDOMNode
   
   If fs.FolderExists("C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests") Then
      Open "C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests\wb.html" For Output As #1
      Set dNode = hDoc.firstChild

      Print #1, hDoc.body.parentElement.outerHTML
      Close #1
   End If
End Sub

Public Property Let BrowserToolTip(strNewValue As String)
   wbFolder.ToolTipText = strNewValue
End Property

Public Property Let DBConnection(objCon As ADODB.Connection)
   blob.DBConnection = objCon
End Property
'
'Public Property Let HomeDirectory(strNewValue As String)
'   homeDir = strNewValue
'   cmdUp.Enabled = False
'   ReDim URLHistory(1)
'   hCount = -1
'   URLHistory(0) = homeDir
''   regEdit.WriteRegistry HKEY_LOCAL_MACHINE, "Software\Anglia\IceConfig\HistoryURL", "URL_" & hCount, ValString, strNewValue
'End Property

Public Sub ShowDirInfo(DirName As String)
   Dim sBuf As New StringBuffer
   Dim buf As String
   Dim tFile As String
   Dim imgSrc As String
   Dim fileImg As String
   Dim rowSrc As String
   Dim colId As Integer
   Dim noCols As Integer
   Dim fl As Files
   Dim fileId As File
   Dim i As Integer
   
'   tFile = fs.BuildPath(App.Path, "dataViewer.html")
'   buf = html.Files(DirName, False)
   
'   sBuf.Clear
   sBuf.Append "<table>" & vbCrLf
   sBuf.Append "<tr>" & _
             "<td class=foldertitle colspan=6>" & DirName & "</td>" & vbCrLf & _
             "</tr>" & vbCrLf
   imgSrc = "<img action=over height=16 width=16 alt='Click to open' border=0 src='C:\\ICE\\LABCOMM\\Icons\\"
   
   Set fl = fs.GetFolder(DirName).Files
   If fl.Count Mod 25 > 0 Then
      noCols = Int(fl.Count / 20) + 1
   Else
      noCols = (fl.Count / 20)
   End If
   
   For Each fileId In fl
      
      Select Case UCase(fs.GetExtensionName(fileId.Name))
         Case "XMS"
            fileImg = imgSrc & "xms.ico'>"
            
         Case "XEN"
            fileImg = imgSrc & "xen.ico'>"
            
         Case "XML"
            fileImg = imgSrc & "html.ico'>"
                     
         Case "WNG"
            fileImg = imgSrc & "warning.ico'>"
         
         Case "ERR"
            fileImg = imgSrc & "error.ico'>"
         
         Case Else
            fileImg = imgSrc & "unknown.ico'>"
         
      End Select
      
      rowSrc = rowSrc & _
               "   <td onclick=handleClick class=folder>" & fileImg & "<span action=over>&nbsp;" & fileId.Name & "</span></td>"
      colId = colId + 1

      If colId = noCols Then
         sBuf.Append "<tr>" & rowSrc & "</tr>" & vbCrLf
         rowSrc = ""
         colId = 0
      End If
   Next
   
   If colId > 0 Then
      sBuf.Append "<tr>" & rowSrc & "</tr>" & vbCrLf
   End If
   
   sBuf.Append "</table>" & vbCrLf
   
   Style = "Folderview"
   Script = "viewcontrol"
   
   txtTitle.Text = DirName
   NavigateTo sBuf.ActualValue
'   Open tFile For Output As #1
'   Print #1, buf
'   Close #1
'
'   wbFolder.Navigate2 tFile
End Sub

Public Property Let StartHTML(strNewValue As String)
   wbHTMLStart = strNewValue
End Property

Public Property Let EndHTML(strNewValue As String)
   wbHTMLEnd = strNewValue
End Property

Public Property Let LocationToolTip(strNewValue As String)
   lblLoc.ToolTipText = strNewValue
End Property

Public Property Let NavigationBar(intNewValue As Integer)
   Select Case intNewValue
      Case 0
         fraNav.Visible = False
         fraInfo.Visible = False
         wbFolder.Top = 60
         wbFolder.Height = UserControl.Height - 375
         
      Case 1
         fraNav.Visible = True
         fraInfo.Visible = False
         wbFolder.Top = 930
         wbFolder.Height = UserControl.Height - 1245
      
      Case 2
         fraNav.Visible = False
         fraInfo.Visible = True
         wbFolder.Top = 930
         wbFolder.Height = UserControl.Height - 1245
   
   End Select
'   If blnNewValue Then
'      fraNav.Visible = True
'      fraInfo.Visible = False
'   Else
'      fraNav.Visible = False
'      fraInfo.Visible = True
'   End If
End Property

Public Function RunAgainstIEVersion(Optional MinVersion As String = "") As Boolean
   Dim VersionInfo As DllVersionInfo
  
   If Not IsNumeric(MinVersion) Then
      MinVersion = 6
   End If
   
   VersionInfo.cbSize = Len(VersionInfo)
   Call DllGetVersion(VersionInfo)
   RunAgainstIEVersion = (VersionInfo.dwMajorVersion >= Val(MinVersion))
End Function

'Public Sub SetHTMLTemplate(TemplateID As enum_TEMPLATETYPES, _
'                           TemplateInfo As String)
'   Dim blnFile As Boolean
'
'   blnFile = (fs.GetFile(TemplateInfo) Or fs.GetFile(fs.BuildPath(App.Path, TemplateInfo)))
'
'   Select Case TemplateID
'      Case 1
'         If blnFile Then
'            reportURL = TemplateInfo
'         Else
'            reportURL = fs.BuildPath(App.Path, "repTemplate.html")
'            Open reportURL For Output As #1
'            Print #1, TemplateInfo
'            Close #1
'         End If
'
'      Case 2
'         If blnFile Then
'            navBarURL = TemplateInfo
'         Else
'            navBarURL = fs.BuildPath(App.Path, "navBarTemplate.html")
'            Open navBarURL For Output As #1
'            Print #1, TemplateInfo
'            Close #1
'         End If
'
'      Case 3
'         If blnFile Then
'            logURL = TemplateInfo
'         Else
'            logURL = fs.BuildPath(App.Path, "logTemplate.html")
'            Open logURL For Output As #1
'            Print #1, TemplateInfo
'            Close #1
'         End If
'
'   End Select
'End Sub

Public Property Let InfoBGColour(strNewValue As String)
   bgCol = strNewValue
End Property

Public Property Let InfoDescription(strNewValue As String)
   txtDesc.Text = strNewValue
End Property

Public Property Let InfoCallBack(strNewObj As Object)
   Set cObj = strNewObj
End Property

Public Property Let InfoTitle(strNewValue As String)
   txtTitle.Text = strNewValue
End Property

Public Property Let ViewerModule(objNewValue As Object)
   Set viewClass = objNewValue
End Property
'
'Private Sub cmdNew_Click()
'   Dim ie As Object
'
'   Set ie = CreateObject("InternetExplorer.Application")
'   ie.Navigate2 lblLoc.Caption
'   ie.Visible = True
'End Sub
'
'Private Sub cmdHome_Click()
'  NavigateTo homeDir, True
'End Sub
'
'Private Sub cmdUp_Click()
'   On Error GoTo procEH
'   Dim URL As String
'
'   hCount = hCount - 1
'   URL = URLHistory(hCount)
'   If hCount = 0 Then
'      cmdUp.Enabled = False
'   End If
'   wbFolder.Navigate2 URL
''   wbFolder.GoBack
''   If wbFolder.LocationURL = homeDir Then
''      cmdUp.Enabled = False
''   End If
'   Exit Sub
'
'procEH:
'   cmdUp.Enabled = False
'End Sub

Private Function hDoc_oncontextmenu() As Boolean
   hDoc_oncontextmenu = False
End Function

Public Property Let LocationTitle(strNewValue As String)
   lblLoc.Caption = strNewValue
   If Not cObj Is Nothing Then
      lblLoc.Tag = cObj.CallBack("LocationFileName")
   End If
End Property

Public Property Let CallBack(oMod As Object)
   Set remoteClass = oMod
End Property
'
'Public Property Let EventHandlerClass(evClass As Object)
'   Set evObj = evClass
'End Property

Public Sub FireDataEvent()
   hDoc.FireEvent "ondatasetchanged"
End Sub
'
'Public Property Let ScriptDirectory(strNewValue As String)
'   If fs.FolderExists(strNewValue) Then
'      scriptDir = strNewValue
'   Else
'      MsgBox "Invalid script directory: (" & strNewValue & ")", vbExclamation, "Control validation"
'      scriptDir = ""
'   End If
'End Property
'
'Public Property Let StyleDirectory(strNewValue As String)
'   If fs.FolderExists(strNewValue) Then
'      styleDir = strNewValue
'   Else
'      MsgBox "Invalid style directory: (" & strNewValue & ")", vbExclamation, "Control validation"
'      styleDir = ""
'   End If
'End Property
'
'Public Sub SetDataEventParams(paramData As String)
'   Dim body As MSHTML.HTMLBody
'
'   If paramData <> "" Then
'      Set body = hDoc.body
'      body.setAttribute "param", paramData
'   End If
'End Sub

Public Sub NavigateTo(Destination As Variant, _
                      Optional FileName As Boolean = False)
   On Error GoTo procEH
   Dim ss As MSHTML.IHTMLStyleSheet
   Dim ssheet As New MSHTML.HTMLStyleSheetsCollection
   Dim styleNode As MSHTML.HTMLStyleElement
   Dim tnode As MSHTML.HTMLTextElement
   Dim dNode As MSHTML.IHTMLDOMNode
   Dim ec As MSHTML.IHTMLElementCollection
   Dim impName As String
   Dim fTxt As TextStream
   Dim i As Integer
   Dim stylePresent As Boolean
   Dim fss As MSHTML.IHTMLStyleSheet
   Dim bp As MSHTML.IHTMLScriptElement
   
   DoEvents
   lblLoc.Visible = False
   
   If fs.FileExists(Destination) Or _
      fs.FolderExists(Destination) Or _
      Destination = "" Then
      
'      Set wbEv.evSink = wbDirect
      If Destination = "" Then
         wbDirect.Navigate2 fs.BuildPath(App.Path, "logo1.jpg")
         Do Until wbDirect.Document.ReadyState = "complete" '.ReadyState = READYSTATE_COMPLETE
            DoEvents
         Loop
      Else
         wbDirect.Navigate2 Destination
      End If
      
'      Set dDoc = wbDirect.Document
      
      wbFolder.Visible = False
      wbDirect.Visible = True
   Else
      With wbFolder
         Set hDoc = Nothing
         Set hDoc = wbFolder.Document
         
         hDoc.body.innerHTML = ""
         
         Do Until .Document.ReadyState = "complete" '.ReadyState = READYSTATE_COMPLETE
            DoEvents
         Loop
      End With
               
      hDoc.bgColor = bgCol
      hDoc.Title = txtTitle.Text
      
      hDoc.body.innerHTML = Destination
      Do Until wbFolder.ReadyState = READYSTATE_COMPLETE Or hDoc.ReadyState = "interactive"
         DoEvents
      Loop
      
      If Not cObj Is Nothing Then
         cObj.CallBack "IceConfig_ShowSourceFile", Abs(CInt(optShowFile(1).Value)) + 1
      End If
      
'**************************************
'     Any special start or end HTML?
'**************************************
      If wbHTMLStart <> "" Then
         hDoc.body.innerHTML = wbHTMLStart & hDoc.body.innerHTML
      End If
      
      If wbHTMLEnd <> "" Then
         hDoc.body.innerHTML = hDoc.body.innerHTML & wbHTMLEnd
      End If
         
'*************************
'     Add a stylesheet?
'*************************
      RemoveStyling
      
      If wbStyle <> "" Then
'        Create as file to link to
         impName = fs.BuildPath(App.Path, wbStyleName & ".css")
         
         If fs.FileExists(impName) = False Then
            Set fTxt = fs.OpenTextFile(impName, ForWriting, True)
            fTxt.Write wbStyle
            fTxt.Close
         End If
         wbStyle = ""
         
'        Is the top level link already present?
         If hDoc.styleSheets.Length = 0 Then
            Set ss = hDoc.createStyleSheet(styleID)
         Else
            Set ss = hDoc.styleSheets(0)
         End If
         
'        Import the child link under the top link
         For i = 0 To ss.imports.Length - 1
            Set fss = ss.imports(i)
            If fss.href = impName Then
               stylePresent = True
               Exit For
            End If
         Next i
         
         If stylePresent = False Then
            ss.addImport impName
         End If
            
      End If
         
'*********************
'     Add a script?
'*********************
      Set dNode = hDoc.getElementsByTagName("script")(0)

      If Not dNode Is Nothing Then
         dNode.removeNode
      End If
      
      If wbScript <> "" Then
'        A single script file is linked. Write any scripts to this
         impName = fs.BuildPath(App.Path, "AHSL_Script.js")
         
         Set scriptNode = hDoc.createElement("script")
         Set dNode = hDoc.getElementsByTagName("body")(0)
         dNode.insertBefore scriptNode, dNode.firstChild
         
'        Set the link
         scriptNode.src = impName
            
         Set fTxt = fs.OpenTextFile(impName, ForWriting, True)
         fTxt.Write wbScript
         wbScript = ""
         fTxt.Close
      End If
         
'************************************************
'     Write a diagnostic copy of the document
'************************************************
      If fs.FolderExists("C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests") Then
         Open "C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests\wb.html" For Output As #1
         Set dNode = hDoc.firstChild

         Print #1, hDoc.body.parentElement.outerHTML
         Close #1
      End If
   
      wbDirect.Visible = False
      wbFolder.Visible = True
   End If
                        
   lblLoc.Visible = True
   Exit Sub
   
procEH:
   If Err.Description = "Automation error" Then
      Resume Next
   End If
   If eclass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eclass.CurrentProcedure = "wbCtrl.NavigateTo"
   eclass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Public Property Let ObjectName(strNewValue As String)
   objName = strNewValue
End Property
'
'Public Function ReadIdField(fieldId As String) As String
'   Dim hDoc As HTMLDocument
'   Dim ElId As IHTMLElement
'   Dim dStr As String
'
'   With wbFolder
'      Set hDoc = wbFolder.Document
'      Set ElId = hDoc.getElementById(fieldId)
'      dStr = ElId.innerText
'   End With
'   ReadIdField = dStr
'End Function
'
'Public Sub RemoveScript(scriptId As String)
'   Dim scriptNode As MSHTML.IHTMLScriptElement
'   Dim dNode As MSHTML.IHTMLDOMNode
'   Dim sNode As MSHTML.IHTMLElement
'   Dim scriptBuf As String
'   Dim fileId As String
'   Dim buf As String
'   Dim strArray() As String
'   Dim i As Integer
'
'   Set dNode = hDoc.getElementById(scriptId)
'   If Not dNode Is Nothing Then
'      dNode.removeNode False
'   End If
'End Sub

Public Sub RemoveStyling()
   Dim i As Integer
   Dim ss As MSHTML.IHTMLStyleSheet
   Dim css As MSHTML.HTMLStyleSheetsCollection
   
   If hDoc.styleSheets.Length > 0 Then
      Set ss = hDoc.styleSheets(0)
      For i = 0 To ss.imports.Length - 1
         ss.removeImport i
      Next i
   End If
End Sub

Public Property Let Script(scriptId As String)
   wbScript = blob.Read(scriptId, TT_Script)
End Property

Public Property Let Style(styleID As String)
   wbStyle = blob.Read(styleID, TT_Style)
   wbStyleName = styleID
End Property

Public Property Let FileTitle(strNewValue As String)
   wbTitle = strNewValue
   If Not hDoc Is Nothing Then
      hDoc.Title = strNewValue
'      Debug.Print strNewValue
   End If
End Property
'
'Public Sub SetBrowserType(strNewValue As String)
'   If strNewValue = "Classic" Then
'      expView(0) = regEdit.ReadRegistry(HKEY_CURRENT_USER, _
'                                        "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", _
'                                        "WebView")
'      expView(1) = regEdit.ReadRegistry(HKEY_CURRENT_USER, _
'                                        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
'                                        "ClassicShell")
'
'      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
'                              "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", _
'                              "WebView", _
'                              ValDWord, _
'                              0
'
'      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
'                              "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
'                              "ClassicShell", _
'                              ValDWord, _
'                              1
'   Else
'      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
'                              "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", _
'                              "WebView", _
'                              ValDWord, _
'                              IIf(expView(0) = "", 0, expView(0))
'
'      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
'                              "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
'                              "ClassicShell", _
'                              ValDWord, _
'                              IIf(expView(1) = "", 0, expView(1))
'   End If
'End Sub

Private Sub hDoc_ondataavailable()
   Dim ev As MSHTML.IHTMLEventObj
   Dim dNode As MSHTML.IHTMLDOMNode
   Dim fName As String
   Dim fDir As String
   Dim buf As String
   Dim fExt As String
   
   Set ev = hDoc.parentWindow.event
   Set dNode = ev.srcElement
   fName = dNode.firstChild.nextSibling.firstChild.nodeValue 'ev.srcElement.children(1).Title
   fDir = hDoc.Title
   ev.cancelBubble = True
   ev.returnValue = False
   
   DoEvents
   
   If fs.FileExists(fDir) Then
'     We are viewing a file so show files in the parent directory
   Else
      If hDoc.Title = fName Then
'        The 'Up' button was clicked
         html.ShowSubfolders fs.GetParentFolderName(fDir), False
      ElseIf fs.FolderExists(fs.BuildPath(fDir, fName)) Then
'        A directory was clicked so show the files
         fDir = fs.BuildPath(fDir, fName)
         buf = html.Files(fDir, False)
         Open tFile For Output As #1
         Print #1, buf
         Close #1
         wbDirect.Navigate2 tFile
      Else
'        Display the file
         
         fExt = fs.GetExtensionName(fName)
         
         If UCase(fExt) = "XML" Or Left(UCase(fExt), 3) = "htm" Then
'            fName = "file:///" & Replace(fs.BuildPath(fDir, fName), "\", "/")
            wbDirect.Navigate2 "file:///" & fs.BuildPath(Trim(fDir), Mid(fName, 2))
         
         Else
            fDir = fs.BuildPath(fDir, Mid(fName, 2))
            buf = html.FileToHTML(fDir)
            Open tFile For Output As #1
            Print #1, buf
            Close #1
            wbDirect.Navigate2 tFile
         End If
      End If
   End If
'   Do Until wbFolder.ReadyState = READYSTATE_COMPLETE
'      DoEvents
'   Loop
'   BrowserHTML
   
'   timerDA.Tag = fName
'   timerDA.Enabled = True
End Sub

Private Sub lblLoc_DblClick()
   If Not cObj Is Nothing Then
      cObj.CallBack "FileToNotepad", lblLoc.Tag
   End If
End Sub

Private Sub optShowFile_Click(Index As Integer)
   If Not cObj Is Nothing Then
      cObj.CallBack "IceConfig_ShowSourceFile", Index + 1
   End If
End Sub

Public Function TemplateToArray(TemplateID As String, _
                                TemplateType As enum_TEMPLATETYPES) As Variant
   Dim strArray() As String

   strArray = Split(blob.Read(TemplateID, TemplateType), vbCrLf)
   TemplateToArray = strArray
End Function

Private Sub UserControl_Initialize()
   NavigationBar = 0
   lblLoc.Caption = "Anglia Healthcare Systems Ltd."
   optShowFile(0).Value = False
   optShowFile(1).Value = True
'   wbFolder.Navigate2 "C:\ICE\LABCOMM\HISTORY\MSGOUT\030630\SX_XML_1_XML v2.4_1266_17-26-18.xml"
   wbDirect.Navigate2 fs.BuildPath(App.Path, "logo1.jpg")
   wbFolder.Navigate2 "about:blank"
   tFile = fs.BuildPath(App.Path, "dataViewer.html")
   styleID = fs.BuildPath(App.Path, "AHSL_Style.css")
   If fs.FileExists(styleID) = False Then
      fs.CreateTextFile (styleID)
   End If
End Sub

Private Sub UserControl_InitProperties()
   optShowFile(0).Value = False
   optShowFile(1).Value = True
   eclass.Behaviour = 0
End Sub

Private Sub UserControl_Resize()
   With UserControl
      fraInfo.Width = .Width - 120
      optShowFile(0).Left = .Width - 1965
      optShowFile(1).Left = .Width - 1965
      fraNav.Width = .Width - 120
      
      If (fraInfo.Visible Or fraNav.Visible) Then
         wbFolder.Height = .Height - 1245
         lblLoc.Top = wbFolder.Height + fraInfo.Height + 55
      Else
         wbFolder.Height = .Height - 375
         lblLoc.Top = wbFolder.Height + 55
      End If
      
      wbFolder.Width = .Width - 120
'      lblLoc.Top = wbFolder.Height + fraInfo.Height + 55
      lblLoc.Width = .Width - 120
   End With
End Sub

Private Sub UserControl_Terminate()
   Set evObj = Nothing
   Set cObj = Nothing
End Sub

Private Sub wbDirect_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
'   Stop

End Sub
