VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SegmentValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strInvalid As String
Private strReplace As String

Public Property Get InvalidChars() As String
   InvalidChars = strInvalid
End Property

Public Property Get ReplaceChars()
   ReplaceChars = strReplace
End Property

Public Sub CheckData(ByRef buf As String, _
                     MsgType As String)
   Dim mStart As Long
   Dim mEnd As Long
   Dim iPoint As String
   Dim mLen As Long
   Dim sPos As Long
   Dim sTot As Long
   Dim tBuf As String
   Dim i As Long
   Dim aVal As Integer
   Dim Char As String
   Dim newbuf As String
   Dim repStr As String
   Dim iLen As Long
   Dim RejectCRLF As Boolean
   
   RejectCRLF = (MsgType <> "ASTM")
   i = 1
   Do Until i > Len(buf)
      Char = Mid(buf, i, 1)
      If Asc(Char) < 32 Or Asc(Char) > 126 Then
         If (Asc(Char) = 10 Or Asc(Char) = 13) Then
            If RejectCRLF Then
               strInvalid = strInvalid & Char
               strReplace = strReplace & Chr(0)
            End If
         Else
            strInvalid = strInvalid & Char
            strReplace = strReplace & Chr(0)
         End If
      End If
      i = i + 1
   Loop
   
   If MsgType <> "ASTM" Then
      mStart = InStr(1, buf, "UNH+")
      Do Until mStart = 0
'        Find end of message
         mEnd = InStr(mStart, buf, "'UNT+")
'        Set the insertion point & length for the segment count
         iPoint = mEnd + 5
         iLen = InStr(iPoint, buf, "+") - iPoint
'        Set up the string to be replaced
         repStr = Mid(buf, iPoint, iLen)
      
'        Now find the end and subsequent length of this message
         mEnd = InStr(mEnd + 1, buf, "'") + 1
         mLen = mEnd - mStart
      
         tBuf = Mid(buf, mStart, mLen)
         sPos = InStr(1, tBuf, "'")
         sTot = 0
      
'        Count the segment terminators
         Do Until sPos = 0
            sPos = InStr(sPos + 1, tBuf, "'")
            sTot = sTot + 1
         Loop
      
'        Insert this figure at the relevant point
         buf = Left(buf, iPoint - 1) & Replace(buf, repStr, CStr(sTot), iPoint, 1)
         mStart = InStr(mEnd - 3, buf, "UNH+")
      Loop
   End If
End Sub

