Attribute VB_Name = "Routines"
'--------------------------------------------------------------
' Copyright ©1996-2001 VBnet, Randy Birch, All Rights Reserved.
' Terms of use http://www.mvps.org/vbnet/terms/pages/terms.htm
'--------------------------------------------------------------


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)


Public Function GetArrayDimensions(ByVal arrPtr As Long) As Integer
   Dim address As Long
   CopyMemory address, ByVal arrPtr, ByVal 4
   If address <> 0 Then
      CopyMemory GetArrayDimensions, ByVal address, 2
   End If
End Function


Public Function VarPtrArray(arr As Variant) As Long
   CopyMemory VarPtrArray, ByVal VarPtr(arr) + 8, ByVal 4
End Function

Public Function New_NHSno_Check(Nhsno) As Integer

    Dim Total As Long
    Dim F As Long
    Dim checkdigit As Long
    
    New_NHSno_Check = 2
    If Len(Nhsno) > 0 Then
        If Len(Nhsno) <> 10 Then
            Exit Function
        End If
    End If
    Total = 0
    For F = 1 To 9
      Total = Total + (Val(Mid(Nhsno, F, 1)) * (11 - F))
    Next F
    checkdigit = Total Mod 11
    checkdigit = 11 - checkdigit
    If checkdigit = 11 Then
       checkdigit = 0
    End If
    If checkdigit = 10 Then
        Exit Function
    End If
    If checkdigit = Val(Right(Nhsno, 1)) Then
      New_NHSno_Check = 1
   End If
    
End Function

Public Sub OpenInNotepad(Filename As String, Optional AddLineFeeds As Boolean = False)
   Dim tFile As String
   Dim fStream As TextStream
   Dim Tstr As String
   Dim tBuf As New StringBuffer
   
   On Error GoTo procEH
   Open Filename For Input As #1
   tFile = fs.BuildPath(App.Path, fs.GetTempName)
   Set fStream = fs.CreateTextFile(tFile)
   While Not EOF(1)
      Line Input #1, Tstr
      tBuf.Append Tstr
      If Len(Tstr) > 0 Then
         tBuf.Append vbCrLf
      End If
   Wend
   
   If AddLineFeeds Then
      Tstr = Replace(tBuf.value, "?'", "#apos#")
      Tstr = Replace(Tstr, "'", "'" & vbCrLf)
      Tstr = Replace(Tstr, "#apos#", "'")
   Else
      Tstr = tBuf.value
   End If
   
   fStream.Write tBuf.value
   fStream.Close
   Close #1
   Shell "Notepad.EXE " & tFile, vbNormalFocus '  fraPanel(5).Caption, vbNormalFocus
   Exit Sub
   
procEH:
   Exit Sub
End Sub

Public Sub ResetPropList()

   With frmMain.edipr
      .PropertyItems.Clear
      .Pages.Clear
      .UsePageKeys = False
      .HideColumnHeaders = True
   End With

End Sub


