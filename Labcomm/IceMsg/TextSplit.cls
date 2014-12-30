VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strTxtArr() As String
Private strBuf As String
Private strSubsChars As String
Private strInvalid As String
Private arraySize As Integer
Private RangeHi As String
Private RangeLo As String

Private Sub Class_Initialize()
'   strSubsChar = " "
'   strInvalid = "#$@[\]^_`{|}~"
End Sub

Public Property Let SubstituteCharacter(strNewValue As String)
   strSubsChars = strNewValue
End Property

Public Property Let Invalid(strNewValue As String)
   strInvalid = strNewValue
End Property

Public Function EscapedData(strBuf As String, _
                            Optional Apostrophe As Boolean = True) As String
   Dim i As Integer
   Dim strTemp As String

'  Space fill any shortfall in the replacement string
   If Len(strSubsChars) < Len(strInvalid) Then
      strSubsChars = strSubsChars & Space(Len(strInvalid) - Len(strSubsChars))
   End If
   strTemp = strBuf
   For i = 1 To Len(strInvalid)
      If Apostrophe And Mid(strInvalid, i, 1) = "`" Then
         strTemp = Replace(strTemp, Mid(strInvalid, i, 1), "?'")
      ElseIf Mid(strSubsChars, i, 1) <> Chr(0) Then
         strTemp = Replace(strTemp, Mid(strInvalid, i, 1), Mid(strSubsChars, i, 1))
      Else
         strTemp = Replace(strTemp, Mid(strInvalid, i, 1), "")
      End If
   Next i
   EscapedData = RTrim(strTemp)
      
End Function

Public Function DataChunk(bufIn As String) As Variant

   Dim strBuf As String
   Dim strTxtArr() As String
   Dim i As Integer
   
'  Set i to -1. If it has not been changed then do not return any data.
   i = -1
   ReDim strTxtArr(0)
   strBuf = EscapedData(bufIn, False)  '  Do not replace ` with ?'
'   strBuf = bufIn
   If Len(strBuf) <= 70 Then
      If Len(RTrim(strBuf)) > 0 Then
         i = 0
         ReDim strTxtArr(0)
         strTxtArr(0) = RTrim(strBuf)
      End If
   Else
      i = 0
      Do
'        Do not include blank lines in the data
         If Len(RTrim(Left(strBuf, 70))) > 0 Then
            ReDim Preserve strTxtArr(i)
            strTxtArr(i) = RTrim(Left(strBuf, 70))
            i = i + 1
         End If
         strBuf = Mid(strBuf, 71)
      Loop Until Len(strBuf) <= 70
      
'     Append the last chunk
      If Len(RTrim(strBuf)) > 0 Then
         ReDim Preserve strTxtArr(i)
         strTxtArr(i) = RTrim(strBuf)
      End If
   End If
   
'  Any data to return?
   If i > -1 And strTxtArr(0) <> "" Then
      DataChunk = strTxtArr
   End If
End Function

Public Sub AdjustRangeData(LoRange As String, _
                           HiRange As String)
   On Local Error GoTo procEH
   Dim ptr As Integer
   Dim bLen As Integer
   Dim intRLen As Integer
   Dim intSPos As Integer
   
   RangeLo = ""
   RangeHi = ""

'  If a 'Greater than' appears in the low data, strip it out and place the value in the high field
   bLen = Len(Trim(LoRange))
   ptr = InStr(Trim(LoRange), ">")
   If ptr = 1 Then
      RangeHi = Mid(Trim(LoRange), 2)
   ElseIf ptr > 1 Then
      RangeHi = Left(Trim(LoRange), bLen - 1)
   Else
      RangeLo = Trim(LoRange)
   End If
   
'  If a 'Less than' appears in the high data, strip it out and place the value in the low field
   bLen = Len(Trim(HiRange))
   ptr = InStr(Trim(HiRange), "<")
   If ptr = 1 Then
      RangeLo = Mid(Trim(HiRange), 2)
   ElseIf ptr > 1 Then
      RangeLo = Left(Trim(HiRange), bLen - 1)
   Else
      RangeHi = Trim(HiRange)
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "AHSLMessaging.AHSLMsgDb.AdjustRangeData"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function AnalyseRangeData(ResultData As String)
   Dim searchStr(5) As String
   Dim i As Integer
   Dim pos As Integer
   Dim strRes As String
   Dim strRI As String
   Dim strForm As String
   
   searchStr(0) = "~="
   searchStr(1) = "<="
   searchStr(2) = ">="
   searchStr(3) = ">"
   searchStr(4) = "<"
   
   strRes = ""
   strRI = ""
   For i = 0 To 4
      pos = InStr(ResultData, searchStr(i))
      If pos > 0 Then
         strRes = Mid(ResultData, pos + 1)
         Exit For
      End If
   Next i
   If strRes <> "" Then
      Select Case i
         Case 0
            strRI = "3"
         
         Case 1
            strRI = "8"
            
         Case 2
            strRI = "5"
         
         Case 3
            strRI = "6"
         
         Case 4
            strRI = "7"
         
         Case Else
            strRI = ""
            
      End Select
'      AnalyseRangeData = strRes & "|" & strRI
   Else
      strRes = ResultData
   End If
   
   pos = InStr(strRes, ".")
   If pos > 0 Then
      If pos < Len(strRes) Then
         strForm = "0.0" & String(Len(strRes) - (pos + 1), "0")
      Else
         strForm = "#0"
      End If
   Else
      strForm = "#0"
   End If
   AnalyseRangeData = Format(strRes, strForm) & "|" & strRI
End Function

Public Sub FreeText(TextDest As Object, _
                    RS As ADODB.Recordset, _
                    txtType As String, _
                    Optional CommentsSoFar As Integer = 0, _
                    Optional msgId As String = "NHS003", _
                    Optional MaxComments As Integer = 100)
   Dim vTxt As Variant
   Dim cmtType As String
   Dim txtBuf() As String
   Dim j As Integer
   Dim k As Integer
   Dim lCount As Integer
   Dim totComments As Integer
   Dim segCnt As Integer
      
'  Add the comments so far to maintain the FTX positional counter
   totComments = RS.RecordCount
   With TextDest
      cmtType = txtType
      segCnt = .FTX.Count
      
      If RS.RecordCount < MaxComments Then
         Do Until RS.EOF
            eClass.FurtherInfo = "SG18.FTX(" & segCnt & ")"
            vTxt = DataChunk(RS(0))
'           Check some data was returned
            If VarType(vTxt) > 0 Then
               .FTX.Append
'               cmtType = txtType
               
'              Complex Reference range test
               If InStr(1, RS.Source, "ICELABCOMM_Report_Results") > 0 Then
                  If RS(1) = "C" Then
                     If msgId = "NHS003" Then
                        cmtType = "CRR"
                     End If
                  End If
               End If
               
               .FTX.Item(segCnt).TextSubjectQualifier.Value = cmtType
               For k = 0 To UBound(vTxt)
                  Select Case k
                     Case 0
                        .FTX.Item(segCnt).TextLiteral.FreeText1.Value = vTxt(k)
                     Case 1
                        .FTX.Item(segCnt).TextLiteral.FreeText2.Value = vTxt(k)
                     Case 2
                        .FTX.Item(segCnt).TextLiteral.FreeText3.Value = vTxt(k)
                     Case 3
                        .FTX.Item(segCnt).TextLiteral.FreeText4.Value = vTxt(k)
                     Case 4
                        .FTX.Item(segCnt).TextLiteral.FreeText5.Value = vTxt(k)
                  End Select
               Next k
            ElseIf msgId = "NHS003" Then
'              A blank line is represented by placing nothing in FreeTextn.value
'              so append outside the check
               .FTX.Append
               .FTX.Item(segCnt).TextSubjectQualifier.Value = cmtType
            End If
            RS.MoveNext
            segCnt = segCnt + 1
         Loop
         
      Else
         lCount = 0
         Do Until RS.EOF
            eClass.FurtherInfo = "SG18.FTX(" & segCnt & ")"
            vTxt = DataChunk(RS(2))
            If VarType(vTxt) > 0 Then
               For k = 0 To UBound(vTxt)
                  ReDim Preserve txtBuf(lCount)
                  txtBuf(lCount) = vTxt(k)
                  lCount = lCount + 1
               Next k
            ElseIf msgId = "NHS003" Then  '  Output a blank line
               ReDim Preserve txtBuf(lCount)
               lCount = lCount + 1
            End If
         Loop
   
         lCount = UBound(txtBuf)
         For j = 0 To lCount Step 5
            .FTX.Append
            .FTX.Item(segCnt).TextSubjectQualifier.Value = cmtType
            
            .FTX.Item(segCnt).TextLiteral.FreeText1.Value = txtBuf(j)
            
            If (j + 1) < lCount Then
               .FTX.Item(segCnt).TextLiteral.FreeText2.Value = txtBuf(j + 1)
            End If
            
            If (j + 2) < lCount Then
               .FTX.Item(segCnt).TextLiteral.FreeText3.Value = txtBuf(j + 2)
            End If
            
            If (j + 3) < lCount Then
               .FTX.Item(segCnt).TextLiteral.FreeText4.Value = txtBuf(j + 3)
            End If
            
            If (j + 4) < lCount Then
               .FTX.Item(segCnt).TextLiteral.FreeText5.Value = txtBuf(j + 4)
            End If
            
            segCnt = segCnt + 1
         Next j
      End If
   End With
End Sub

Friend Property Get HighRange() As String
   HighRange = RangeHi
End Property

Friend Property Get LowRange() As String
   LowRange = RangeLo
End Property

Public Sub RemoveInvalidChars(ByRef buf As String)
   Dim i As Long
   Dim Char As String
   
   i = 1
   Do Until i > Len(buf)
      Char = Mid(buf, i, 1)
      If Asc(Char) < 32 Or Asc(Char) > 126 Then
         buf = Replace(buf, Char, "")
      End If
      i = i + 1
   Loop
End Sub
