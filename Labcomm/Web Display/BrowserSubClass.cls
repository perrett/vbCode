VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrowserSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private BrowserId As SHDocVwCtl.WebBrowser
Private mainHWnd As Long
Private firstWndProc As Long
Private classType As String

Private collPWP As New Collection

Friend Property Get BrowserClassType() As String
   BrowserClassType = classType
End Property

Public Property Let BrowserType(strNewValue As String)
   If strNewValue = "Classic" Then
      expView(0) = regEdit.ReadRegistry(HKEY_CURRENT_USER, _
                                        "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", _
                                        "WebView")
      expView(1) = regEdit.ReadRegistry(HKEY_CURRENT_USER, _
                                        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
                                        "ClassicShell")
      
      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
                              "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", _
                              "WebView", _
                              ValDWord, _
                              0
                              
      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
                              "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
                              "ClassicShell", _
                              ValDWord, _
                              1
   Else
      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
                              "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", _
                              "WebView", _
                              ValDWord, _
                              IIf(expView(0) = "Not Found", 0, expView(0))
                              
      regEdit.EditRegistryKey HKEY_CURRENT_USER, _
                              "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
                              "ClassicShell", _
                              ValDWord, _
                              IIf(expView(1) = "Not Found", 0, expView(1))
   End If
End Property

Sub EnumWindows(Level As Integer, ByVal HWnd As Long)
   Dim S As String
   Dim prevProc As Long
   Dim wndClass As String
   Dim CurrWnd As Long, x
   Dim count%
   Dim bData As New clsBrowserData
   
   RecordWndProc HWnd
   
   HWnd = GetWindow(HWnd, GW_CHILD)

   count = 0
'  Loop while the hWnd returned by GetWindow is valid.
   While HWnd <> 0
      count = count + 1
      RecordWndProc HWnd
'      Call EnumWindows(Level + 1, HWnd)
'     Debug.Print String$(Level, Chr$(9)); count; ". "; Hex$(HWnd)
    
'     Get the next task list item in the master list.
      HWnd = GetWindow(HWnd, GW_HWNDNEXT)
   Wend
End Sub

Friend Function FindBrowserWndProc(HWnd As Long) As Long
   On Error GoTo procEH
   
   FindBrowserWndProc = collwndProc("W_" & HWnd)
   Exit Function
   
procEH:
   FindBrowserWndProc = 0
End Function
'
'Friend Sub NavigateTo(Destination As Variant, _
'                      Optional HTMLString As Boolean = False)
'   Dim hDoc As MSHTMLCtl.HTMLDocument
'
'   With BrowserId
'      If HTMLString Then
'         .Navigate2 "about:blank"
'         Do Until .ReadyState = READYSTATE_COMPLETE
'            DoEvents
'         Loop
'         Set hDoc = wbFolder.Document
'         hDoc.body.innerHTML = Destination
'      Else
'         .Navigate2 Destination
'         Do Until .ReadyState = READYSTATE_COMPLETE
'            DoEvents
'         Loop
'      End If
'   End With
'End Sub

Friend Property Let OriginalWndProc(lngNewValue As Long)
   Dim bData As New clsBrowserData
   
   Debug.Print "Original let: " & lngNewValue
   bData.OriginalWndProc = lngNewValue
   collPWP.Add bData
   Set bData = Nothing
End Property

Friend Property Get OriginalWndProc() As Long
   Dim bData As clsBrowserData
   
   Set bData = collPWP(collPWP.count)
   collPWP.Remove collPWP.count
   OriginalWndProc = bData.OriginalWndProc
   Debug.Print "Original get: " & bData.OriginalWndProc
   Set bData = Nothing
End Property

Friend Property Let OldHWnd(lngNewValue As Long)
   Dim bData As New clsBrowserData
   
   bData.OldHWnd = lngNewValue
   collPWP.Add bData
   Set bData = Nothing
End Property

Friend Property Get OldHWnd() As Long
   Dim bData As clsBrowserData
   
   Set bData = collPWP(collPWP.count)
   collPWP.Remove collPWP.count
   OldHWnd = bData.OldHWnd
   Set bData = Nothing
End Property

Friend Function ProcBindToBrowser() As Long
   EnumWindows 0, mainHWnd
   Debug.Print "clsBrowser.ProcBindToBrowser (" & mainHWnd & ")"
   ProcBindToBrowser = mainHWnd
End Function

Friend Function ProcUnBindFromBrowser() As Long
   Dim i As Integer
   Dim bData As clsBrowserData
   
   Debug.Print "ProcUnBindFromBrowser"
   For i = 1 To collPWP.count
      Set bData = collPWP(i)
      SetWindowLong bData.OldHWnd, GWL_WNDPROC, bData.OriginalWndProc
      collPWP.Remove i
      collwndProc.Remove "W_" & bData.OldHWnd
   Next i
End Function

Public Sub RecordWndProc(HWnd As Long)
   Dim S As String
   Dim wndClass As String
'   Dim CurrWnd As Long, x
   Dim bData As New clsBrowserData
   
'  Get the hWnd of the first item in the master list
'  so we can process the task list entries (top-level only).

   wndClass = vbGetWindowTextLine(HWnd)
   If wndClass = "" Then
      wndClass = vbGetClassName(HWnd)
   End If
   
'   Debug.Print "Class: " & wndClass
   S = "{" & Hex$(HWnd) & "}" & "[" & wndClass & "]"
'   Debug.Print String$(Level * 2, " ") & S
  
'  wndClass = "HTML_Internet Explorer" for IE3
'  wndClass = "Internet Explorer_Server" for IE4
   If HWnd <> 0 And (wndClass = "HTML_Internet Explorer" Or _
      wndClass = "Internet Explorer_Server" Or _
      wndClass = "FolderView") Then
      bData.OriginalWndProc = GetWindowLong(HWnd, GWL_WNDPROC)
      bData.OldHWnd = HWnd
      SetWindowLong HWnd, GWL_WNDPROC, AddressOf HTMLWndProc
      Debug.Print "clsBrowser.RecordWndProc: " & wndClass & "(hWnd = " & bData.OldHWnd & " wndProc = " & bData.OriginalWndProc & ")"
      collPWP.Add bData
      collwndProc.Add bData.OriginalWndProc, "W_" & HWnd
      classType = wndClass
'   Else
'      Debug.Print "clsBrowser.RecordWndProc: wndClass = " & wndClass
   End If
   
   HWnd = GetWindow(HWnd, GW_CHILD)
   Set bData = Nothing
End Sub

Friend Function Register(objNewValue As SHDocVwCtl.WebBrowser) As Long
   Set BrowserId = objNewValue
   BrowserId.SetFocus
   mainHWnd = GetFocus
   SetFocusToBrowser mainHWnd
   
   firstWndProc = GetWindowLong(mainHWnd, GWL_WNDPROC)
   Debug.Print "clsBrowser.Browser: MainHwnd = " & mainHWnd & " + OldWndProc: " & firstWndProc
   SetWindowLong mainHWnd, GWL_WNDPROC, AddressOf HTMLWndProc
   collwndProc.Add firstWndProc, "W_" & mainHWnd
End Function

Function SetFocusToBrowser(hBrowserHwnd As Long) As Long
   Dim lStyle As Long
   Dim lResult As Long
   Dim HWnd As Long
   
   HWnd = hBrowserHwnd
   While (lResult = 0) And (HWnd <> 0)
      HWnd = GetWindow(HWnd, GW_CHILD)
      lStyle = GetWindowLong(HWnd, GWL_STYLE)
      lResult = lStyle And WS_VSCROLL
   Wend
   SetFocusAPI (HWnd)
   SetFocusToBrowser = HWnd
End Function

Function vbGetWindowText(ByVal HWnd) As String
   Dim c As Integer
   Dim S As String
   
   c = GetWindowTextLength(HWnd)
   If c <= 0 Then Exit Function
   S = String$(c, 0)
   c = GetWindowText(HWnd, S, c + 1)
   vbGetWindowText = S
End Function

Function vbGetWindowTextLine(ByVal HWnd) As String
   Dim sTitle As String
   Dim cTitle As Integer
   
   sTitle = vbGetWindowText(HWnd)

'  Chop off end of multiline captions
   cTitle = InStr(sTitle, vbCr)
   vbGetWindowTextLine = IIf(cTitle, Left$(sTitle, cTitle), sTitle)
End Function

Function vbGetClassName(ByVal HWnd) As String
   Dim sName As String
   Dim cName As Integer
   
   sName = String$(41, 0)
   cName = GetClassName(HWnd, sName, 41)
   vbGetClassName = Left$(sName, cName)
End Function

Friend Property Get WndHandle() As Long
   WndHandle = mainHWnd
End Property

