VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgRoutines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strInvalid As String
Private strReplace As String
Private RangeHi As String
Private RangeLo As String
Private outRS As New ADODB.Record

Public Function AddImpExpMessage(ReportId As String, _
                             HeaderRef As Long, _
                             Discipline As String, _
                             dateTimeStamp As String, _
                             MsgStat As Long, _
                             Optional MsgComment As String = "", _
                             Optional reject As Long = 0, _
                             Optional fileId As String = "", _
                             Optional MsgPosn As Long = 0, _
                             Optional LTSIndex As Long = 0)

'  Write new Service_ImpExp_Message for all affected messages

   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim iceCmd As New ADODB.Command
   Dim svcMessId As Long
   Dim patLocId As String
   Dim patName As String
   Dim svcId As String
   Dim patSex As Long
   Dim patDOB As Variant
   Dim svcDate As Variant
   Dim dest As String
   Dim svcRepInd As Long
   Dim RepId As String
   Dim indTo As Long
   Dim rejectCnt As Long
   Dim Fname As String

'  Select the original message
   strSQL = "SELECT SM.* " & _
            "FROM Service_ImpExp_Messages SM " & _
               "INNER JOIN Service_ImpExp_Headers SH " & _
               "ON SM.Service_ImpExp_Id = SH.Service_ImpExp_Id " & _
            "WHERE Service_Type = 1 " & _
               "AND Service_Report_Index = " & ReportId
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly

   If LogLevel = 2 Then
      If RS.BOF And RS.EOF Then
         eClass.LogMessage "No records found"
      Else
         eClass.LogMessage RS.RecordCount & " records found"
      End If
   End If

   eClass.FurtherInfo = "ICEMSG_Ins_Service_ImpExp_Message"

   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_Ins_Service_ImpExp_Message"
      .CommandTimeout = 1
      .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
      .Parameters.Append .CreateParameter("Service_ImpExp_Id", adInteger, adParamInput, , HeaderRef)
      .Parameters.Append .CreateParameter("Service_Message_Id", adVarChar, adParamInput, 50, MsgPosn)
      .Parameters.Append .CreateParameter("Patient_Local_Id", adVarChar, adParamInput, 30, RS!Patient_Local_ID)
      .Parameters.Append .CreateParameter("Service_Id", adVarChar, adParamInput, 40, RS!Service_Id)
      .Parameters.Append .CreateParameter("Patient_Name", adVarChar, adParamInput, 70, RS!Patient_Name)
      .Parameters.Append .CreateParameter("Patient_Sex", adInteger, adParamInput, , RS!Patient_Sex)
      .Parameters.Append .CreateParameter("Patient_DOB", adDate, adParamInput, , RS!Patient_DOB)
      .Parameters.Append .CreateParameter("Service_Date_Of_Service", adDate, adParamInput, , RS!Service_Date_Of_Service)
      .Parameters.Append .CreateParameter("Service_Message_Comment", adVarChar, adParamInput, 120, IIf(MsgComment = "", RS!Service_Message_Comment, MsgComment))
      .Parameters.Append .CreateParameter("Rejected_Count", adInteger, adParamInput, , IIf(reject = 0, RS!Rejected_Count, reject))
      .Parameters.Append .CreateParameter("Rejected_File", adVarChar, adParamInput, 128, IIf(fileId = "", RS!Rejected_File, fileId))
      .Parameters.Append .CreateParameter("Date_Added", adDate, adParamInput, , dateTimeStamp)
      .Parameters.Append .CreateParameter("Destination", adVarChar, adParamInput, 30, RS!Destination)
      .Parameters.Append .CreateParameter("Service_Report_Index", adInteger, adParamInput, , RS!Service_Report_Index)
      .Parameters.Append .CreateParameter("Discipline", adInteger, adParamInput, , Discipline)
      .Parameters.Append .CreateParameter("To_EDI", adBoolean, adParamInput, , True)
      .Parameters.Append .CreateParameter("Report_Identifier", adVarChar, adParamInput, 35, RS!Report_Identifier)
      .Parameters.Append .CreateParameter("EDI_Individual_Index_To", adInteger, adParamInput, , RS!EDI_Individual_Index_To)
      .Parameters.Append .CreateParameter("LTSINDEX", adInteger, adParamInput, , IIf(LTSIndex = 0, Null, LTSIndex))
      .Parameters.Append .CreateParameter("MsgStat", adInteger, adParamInput, , MsgStat)
      .Execute
      AddImpExpMessage = .Parameters("Return").Value
   End With
   RS.Close
   Set RS = Nothing
End Function

Public Sub AddImpExpComment(HeaderRef As Long, _
                            Comment As String, _
                            dateTimeStamp As String, _
                            Optional Code As Long = 0, _
                            Optional Source As String = "IceMsg", _
                            Optional MsgRef As Long = 0)
   Dim iceCmd As New ADODB.Command
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_Ins_Service_ImpExp_Comment"
      .Parameters.Append .CreateParameter("ImpExpRef", adInteger, adParamInput, , HeaderRef)
      .Parameters.Append .CreateParameter("Comment", adVarChar, adParamInput, 350, Comment)
      .Parameters.Append .CreateParameter("Source", adVarChar, adParamInput, 10, Source)
      .Parameters.Append .CreateParameter("Code", adInteger, adParamInput, , Code)
      .Parameters.Append .CreateParameter("When", adDate, adParamInput, , dateTimeStamp)
      If MsgRef > 0 Then
         .Parameters.Append .CreateParameter("MsgRef", adInteger, adParamInput, , MsgRef)
      End If
      .Execute
   End With
   Set iceCmd = Nothing
End Sub

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
   eClass.CurrentProcedure = "MsgRoutines.AdjustRangeData"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function AnalyseResultString(ResultData As String) As String
   On Error GoTo procEH
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
         If Len(Trim(Left(ResultData, pos - 1))) = 0 Then
'           There are no characters before the seachstr
            strRes = Mid(ResultData, pos + 1)
            Exit For
         End If
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
   Else
      strRes = ResultData
   End If
   
   If IsNumeric(strRes) Then
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
      
      AnalyseResultString = Format(strRes, strForm) & "|" & strRI
   Else
      AnalyseResultString = strRes & "|" & strRI
   End If
   
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "MsgRoutines.AnalyseResultString"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function DataToArray(bufIn As String, _
                            Optional ArrayLength As Integer = 70) As Variant
   On Error GoTo procEH
   Dim strBuf As String
   Dim strTxtArr() As String
   Dim i As Integer
   
'  Set i to -1. If it has not been changed then do not return any data.
   i = -1
   ReDim strTxtArr(0)
'   strBuf = EscapeData(bufIn, False)  '  Do not replace ` with ?'
   strBuf = bufIn
   If Len(strBuf) <= ArrayLength Then
      If Len(RTrim(strBuf)) > 0 Then
         i = 0
         ReDim strTxtArr(0)
         strTxtArr(0) = RTrim(strBuf) & Chr(255)
      End If
   Else
      i = 0
      Do
'        Do not include blank lines in the data
         If Len(RTrim(Left(strBuf, ArrayLength))) > 0 Then
            ReDim Preserve strTxtArr(i)
            strTxtArr(i) = RTrim(Left(strBuf, ArrayLength)) & Chr(255)
            i = i + 1
         End If
         strBuf = Mid(strBuf, ArrayLength + 1)
      Loop Until Len(strBuf) <= ArrayLength
      
'     Append the last chunk
      If Len(RTrim(strBuf)) > 0 Then
         ReDim Preserve strTxtArr(i)
         strTxtArr(i) = RTrim(strBuf) & Chr(255)
      End If
   End If
   
'  Any data to return?
   If i > -1 And strTxtArr(0) <> "" Then
      DataToArray = strTxtArr
   End If
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "MsgRoutines.DataToArray"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function


Public Function EscapeData(strBuf As String, _
                           Optional ReleaseChar As String = "?", _
                           Optional RetainGrave As Boolean = False) As String
   On Error GoTo procEH
   Dim i As Integer
   Dim strTemp As String
   Dim escChar As String
   
   escChar = ReleaseChar & "'"
   
   strTemp = strBuf
   
   If RetainGrave = False Then
      strTemp = Replace(strBuf, "`", escChar)
   End If
   
'  Space fill any shortfall in the replacement string
   If Len(strReplace) < Len(strInvalid) Then
      strReplace = strReplace & Space(Len(strInvalid) - Len(strReplace))
   End If
   
   For i = 1 To Len(strInvalid)
      If Mid(strReplace, i, 1) = Chr(0) Then
         strTemp = Replace(strTemp, Mid(strInvalid, i, 1), "")
      Else
         strTemp = Replace(strTemp, Mid(strInvalid, i, 1), Mid(strReplace, i, 1))
      End If
   Next i
   
   EscapeData = RTrim(strTemp)
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "MsgRoutines.EscapeData"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Sub FreeText(TextDest As Object, _
                    RS As ADODB.Recordset, _
                    txtType As String, _
                    Optional CommentsSoFar As Integer = 0, _
                    Optional msgId As String = "NHS003", _
                    Optional MaxComments As Integer = 100)
   On Error GoTo procEH
   Dim sbuf As New StringBuffer
   Dim vTxt As Variant
   Dim cmtType As String
   Dim txtBuf() As String
   Dim j As Integer
   Dim k As Integer
   Dim lCount As Integer
   Dim totComments As Integer
   Dim SegCnt As Integer
   
'  Add the comments so far to maintain the FTX positional counter
   totComments = RS.RecordCount
   With TextDest
      cmtType = txtType
      SegCnt = .FTX.Count
      
      If RS.RecordCount < MaxComments Then
         Do Until RS.EOF
            SegCnt = .FTX.Count
            eClass.FurtherInfo = "SG18.FTX(" & SegCnt & ")"
            vTxt = DataToArray(RS(0))
'           Check some data was returned
            If VarType(vTxt) > 0 Then
               .FTX.Append
'               cmtType = txtType
               
'              Complex Reference range test
               If InStr(1, RS.Source, "ICELABCOMM_Report_ResultComment") > 0 Then
                  If RS(1) = "C" Then
                     If msgId = "NHS003" Then
                        cmtType = "CRR"
                     End If
                  ElseIf RS(1) = "S" Then
                     cmtType = "SPC"
                  End If
               End If
               
               .FTX.Item(SegCnt).TextSubjectQualifier.Value = cmtType
               For k = 0 To UBound(vTxt)
                  vTxt(k) = Replace(vTxt(k), "'", "`")
                  Select Case k
                     Case 0
                        .FTX.Item(SegCnt).TextLiteral.FreeText1.Value = vTxt(k)
                     Case 1
                        .FTX.Item(SegCnt).TextLiteral.FreeText2.Value = vTxt(k)
                     Case 2
                        .FTX.Item(SegCnt).TextLiteral.FreeText3.Value = vTxt(k)
                     Case 3
                        .FTX.Item(SegCnt).TextLiteral.FreeText4.Value = vTxt(k)
                     Case 4
                        .FTX.Item(SegCnt).TextLiteral.FreeText5.Value = vTxt(k)
                  End Select
               Next k
            ElseIf msgId = "NHS003" Then
'              A blank line is represented by placing nothing in FreeTextn.value
'              so append outside the check
               .FTX.Append
               .FTX.Item(SegCnt).TextSubjectQualifier.Value = cmtType
            End If
            RS.MoveNext
'            segCnt = segCnt + 1
         Loop
         
      Else
         lCount = 0
         Do Until RS.EOF
            eClass.FurtherInfo = "SG18.FTX(" & SegCnt & ")"
            vTxt = DataToArray(RS(0))
            If VarType(vTxt) > 0 Then
               For k = 0 To UBound(vTxt)
                  ReDim Preserve txtBuf(lCount)
                  txtBuf(lCount) = Replace(vTxt(k), "'", "`") & Chr(255)
                  lCount = lCount + 1
               Next k
'            ElseIf msgId = "NHS003" Then  '  Output a blank line
'               ReDim Preserve txtBuf(lCount)
'               lCount = lCount + 1
            End If
            RS.MoveNext
         Loop
   
         lCount = UBound(txtBuf)
         For j = 0 To lCount Step 5
            .FTX.Append
            .FTX.Item(SegCnt).TextSubjectQualifier.Value = cmtType
            
            .FTX.Item(SegCnt).TextLiteral.FreeText1.Value = txtBuf(j)
            
            If (j + 1) < lCount Then
               .FTX.Item(SegCnt).TextLiteral.FreeText2.Value = txtBuf(j + 1)
            End If
            
            If (j + 2) < lCount Then
               .FTX.Item(SegCnt).TextLiteral.FreeText3.Value = txtBuf(j + 2)
            End If
            
            If (j + 3) < lCount Then
               .FTX.Item(SegCnt).TextLiteral.FreeText4.Value = txtBuf(j + 3)
            End If
            
            If (j + 4) < lCount Then
               .FTX.Item(SegCnt).TextLiteral.FreeText5.Value = txtBuf(j + 4)
            End If
            
            SegCnt = SegCnt + 1
         Next j
      End If
   End With
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "MsgRoutines.FreeText"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Friend Property Get HighRange() As String
   HighRange = RangeHi
End Property

Public Property Let InvalidChars(vNewValue As Variant)
   strInvalid = vNewValue(0)
   strReplace = vNewValue(1)
End Property
'
'Public Property Get InvalidChars() As String
'   InvalidChars = strInvalid
'End Property

Public Property Get LineCount(ByRef fileBuf As String)
   Dim lc As Long
   Dim pos As Long
   
   lc = 0
   pos = InStr(pos + 1, fileBuf, vbCrLf)
   Do Until pos = 0
      lc = lc + 1
      pos = InStr(pos + 1, fileBuf, vbCrLf)
   Loop
   LineCount = lc
End Property

Friend Property Get LowRange() As String
   LowRange = RangeLo
End Property

Public Sub RangeFreeText(TextOrigin As Object, _
                         RS As ADODB.Recordset)
   Dim k As Integer
   Dim ftxCnt As Integer
   
   With TextOrigin
      .TextSubjectQualifier.Value = "RPD"
      ftxCnt = 0
      Do Until RS.EOF
         Select Case ftxCnt
            Case 0
               .TextLiteral.FreeText1.Value = RTrim(RS(0))
            
            Case 1
               .TextLiteral.FreeText2.Value = RTrim(RS(0))
            
            Case 2
               .TextLiteral.FreeText3.Value = RTrim(RS(0))
            
            Case 3
               .TextLiteral.FreeText4.Value = RTrim(RS(0))
            
            Case 4
               .TextLiteral.FreeText5.Value = RTrim(RS(0))
            
         End Select
            
         ftxCnt = ftxCnt + 1
         RS.MoveNext
      Loop
   End With
End Sub

Public Property Let ReplaceChars(strNewValue As String)
   strReplace = strNewValue
End Property

Public Property Get ReplaceChars() As String
   ReplaceChars = strReplace
End Property

Public Sub RecordInvalid(ByRef buf As String)
   Dim Char As String
   Dim i As Long
   
   i = 1
   Do Until i > Len(buf)
      Char = Mid(buf, i, 1)
      Select Case Asc(Char)
         Case Is < 14
'           CR or LF rejections are set by the message class
            If (Asc(Char) <> 10 And Asc(Char) <> 13) Then
               strInvalid = strInvalid & Char
               strReplace = strReplace & Chr(0)
            End If
         
         Case Is < 32
            strInvalid = strInvalid & Char
            strReplace = strReplace & Chr(0)
         
         Case Is > 126
            strInvalid = strInvalid & Char
            strReplace = strReplace & Chr(0)
         
      End Select
      
'      If Asc(Char) < 32 Or Asc(Char) > 126 Then
'         If (Asc(Char) = 10 Or Asc(Char) = 13) Then
'            If RejectCRLF Then
'               strInvalid = strInvalid & Char
'               strReplace = strReplace & Chr(0)
'            End If
'         Else
'            strInvalid = strInvalid & Char
'            strReplace = strReplace & Chr(0)
'         End If
'      End If
      i = i + 1
   Loop
End Sub

Public Sub SegmentCount(ByRef segData As String)
   On Error GoTo procEH
   Dim pos As Long
   Dim segs As Long
   
   pos = 1
   Do Until pos > Len(segData)
      pos = InStr(pos, segData, "'")
      If pos > 0 Then
         segs = segs + 1
         pos = pos + 1
      Else
         pos = Len(segData) + 1
      End If
   Loop
   pos = InStrRev(segData, "'UNT+") + 4
   segData = Left(segData, pos) & CStr(segs) & Mid(segData, pos + 2) '  Use 2 to remove the zero placed in here
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "MsgRoutines.SegmentCount"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
