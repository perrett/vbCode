VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
                                         ByVal wCmd As Long) _
                                         As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                                                    ByVal lpString As String, ByVal cch As Long) _
                                                                    As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const WM_ACTIVATE = &H6
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2

Private retryFreq As String
Private failAfter As Integer
Private writeBack As Object

Private Sub Class_Initialize()
   failAfter = 10
   retryFreq = 30000
   Set writeBack = Nothing
End Sub

Public Property Let Attempts(intNewValue As Integer)
   failAfter = intNewValue
End Property

Public Property Let CheckingFrequency(lngNewValue As Long)
   If lngNewValue < 1000 Then
      retryFreq = lngNewValue * 1000
   Else
      retryFreq = lngNewValue
   End If
End Property

Private Function ProcessRunning(windowId As String)
   Dim lngDeskTopHandle As Long    'Used to hold the value of the Desktop handle.
   Dim lngHand As Long             'Used to hold each windows handle as it loops.
   Dim strName As String * 255     'Fixed length string passed to GetWindowText API call.
   Dim lngWindowCount As Long      'Counter used to return the numberof open windows in the system.
   Dim blnRunning As Boolean
   
'  Get the handle for the desktop.
   lngDeskTopHandle = GetDesktopWindow()
   
'  Get the first child of the desktop window.
'  (Note: The desktop is the parent of all windows in the system.
   lngHand = GetWindow(lngDeskTopHandle, GW_CHILD)
   
'  set the window counter to 1.
   lngWindowCount = 1
   
'  Loop while there are still open windows.
   Do While lngHand <> 0
'     Get the title of the next window in the window list.
      GetWindowText lngHand, strName, Len(strName)
        
'     Get the sibling of the current window.
      lngHand = GetWindow(lngHand, GW_HWNDNEXT)
        
'     Make sure the window has a title; and if it does see if it is the one we are looking for.
      If Left$(strName, 1) <> vbNullChar Then
         If UCase(Left(strName, InStr(1, strName, vbNullChar) - 1)) = UCase(windowId) Then
            blnRunning = True
            Exit Do
         End If
      End If
   Loop
   
   ProcessRunning = blnRunning
End Function

Public Property Let FeedbackObject(oNewValue As Object)
  Set writeBack = oNewValue
End Property

Public Function waitFor(windowId As String) As Boolean
   Dim retryCount As Integer
   Dim blnFinished As Boolean
   
   blnFinished = True
   Do Until ProcessRunning(windowId) = False
      If retryCount >= failAfter Then
         blnFinished = False
         Exit Do
      End If
      
      If Not writeBack Is Nothing Then
         writeBack.Caption = "Waiting - " & windowId & " process running"
         DoEvents
      End If
      
      Sleep retryFreq
      retryCount = retryCount + 1
   Loop
   
   waitFor = blnFinished
End Function


