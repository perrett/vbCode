Attribute VB_Name = "browserMod"
Option Explicit

Public collwndProc As New Collection

'--------------------------------------------------------------------------------
'This sample will show you how to use the WebBrowser Control and
'how to disable right click context menu for browser's window.
'
'Requires installed Internet Explorer 3.xx or Internet Explorer 4.xx.
'Works with both versions.
'
'This sample works with All version of VB5.
'--------------------------------------------------------------------------------
'Author   : Serge Baranovsky
'Email    : sergeb@vbcity.com
'Web      : http://www.vbcity.com/
'Date     : 16-07-98
'--------------------------------------------------------------------------------

'#Const bAllowRightClick = True
#Const bAllowRightClick = False

Public encount As Long
Public hwnds() As Long

Public Declare Function GetWindow Lib "user32" (ByVal HWnd As Long, _
    ByVal wCmd As Long) As Long

Public Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal HWnd As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal HWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetFocusAPI Lib "user32" _
    Alias "SetFocus" (ByVal HWnd As Long) As Long

Public Declare Function GetFocus Lib "user32" () As Long

'GetWindow constants
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5

' Window field offsets for GetWindowLong() and GetWindowWord()
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)

Public Const WS_VSCROLL = &H200000

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Const WM_PARENTNOTIFY = &H210

Public mainHWnd As Long
Public prevMainWndProc As Long

Public PrevWndProc() As Long
Public prevWndProcCount As Integer

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal HWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal HWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' ShellExecute Declarations ...
Public Const SW_SHOWDEFAULT = 10

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Const WM_USER = &H400
Const TB_SETSTYLE = WM_USER + 56
Const TB_GETSTYLE = WM_USER + 57
Const TBSTYLE_FLAT = &H800
Const TBSTYLE_ALTDRAG = &H400

Public regEdit As New clsRegEdit
Public expView(2) As String
Public ackFile As String
Public bCtrl As New BrowserSubClass

Public Function HTMLWndProc(ByVal hw As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
   Dim wndProc As Long
   
   Select Case uMsg
      Case WM_RBUTTONDOWN
         Debug.Print "Yeaaa !!! WM_RBUTTONDOWN"
         
'        eat message
      
      Case WM_RBUTTONUP
         Debug.Print "Yeaaa !!! WM_RBUTTONUP"
'        eat message
    
      Case Else
'        check if messages captured for hw
         wndProc = bCtrl.FindBrowserWndProc(hw)
         If wndProc <> 0 Then
'            Debug.Print "Passed to: Old WndProc: " & hw
            
'           handle captured windows messages
            HTMLWndProc = CallWindowProc(wndProc, hw, uMsg, wParam, lParam)
         End If
         
   End Select
End Function
