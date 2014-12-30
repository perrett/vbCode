VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private logMsg As String
Private LogStatus As Long
Private ReportDiscipline As String
Private ReportIndex As Long
Private ReportLTSIndex As Long
Private ReportId As String
Private ReportStatus As Long
Private ReportMsgFormat As String
Private ReportNatCode As String
Private ExtraCopyTo As String
Private impExpRef As Long
Private errFile As String

Private colLogged As New Collection

Friend Sub AndStatus(sVal As Long)
   ReportStatus = (ReportStatus And sVal)
End Sub

Friend Property Let CopyTo(strNewValue As String)
   ExtraCopyTo = strNewValue
End Property

Friend Property Get CopyTo() As String
   CopyTo = ExtraCopyTo
End Property

Friend Property Let Discipline(strNewValue As String)
   ReportDiscipline = strNewValue
End Property

Friend Property Get Discipline() As String
   Discipline = ReportDiscipline
End Property

Friend Property Let EDIIndex(lngNewValue As Long)
   ReportIndex = lngNewValue
End Property

Friend Property Get EDIIndex() As Long
   EDIIndex = ReportIndex
End Property

Friend Property Let ErrorReport(fileName As String)
   errFile = fileName
End Property

Friend Property Get ErrorReport() As String
   ErrorReport = errFile
End Property

Friend Property Let LogData(Status As Long, msgData As String)
   Dim lMsg As New RepLogData
   
   lMsg.MessageData(Status) = msgData  '  ReportId & ": " & msgData
   colLogged.Add lMsg
   Set lMsg = Nothing
End Property

Friend Property Get LoggedMessages(LogLevel As Long) As String
   Dim i As Integer
   Dim lMsg As RepLogData
   Dim msgStr As String
   
   For i = 1 To colLogged.Count
      Set lMsg = colLogged(i)
      msgStr = msgStr & lMsg.MessageData(LogLevel)
   Next i
   
   If Len(msgStr) > 0 Then
      LoggedMessages = Left(msgStr, Len(msgStr) - 1)
   End If
End Property

Friend Property Let LTSIndex(lngNewValue As Long)
   ReportLTSIndex = lngNewValue
End Property

Friend Property Get LTSIndex() As Long
   LTSIndex = ReportLTSIndex
End Property

Friend Property Let MessageFormat(strNewValue As String)
   ReportMsgFormat = strNewValue
End Property

Friend Property Get MessageFormat() As String
   MessageFormat = ReportMsgFormat
End Property

Friend Property Let MessageImpExp(lngNewValue As Long)
   impExpRef = lngNewValue
End Property

Friend Property Get MessageImpExp() As Long
   MessageImpExp = impExpRef
End Property

Friend Property Let NationalCode(strNewValue As String)
   ReportNatCode = strNewValue
End Property

Friend Property Get NationalCode() As String
   NationalCode = ReportNatCode
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = colLogged.[_NewEnum]
End Function

Friend Sub OrStatus(sVal As Long)
   ReportStatus = (ReportStatus Or sVal)
End Sub

Friend Property Let ReportName(vNewValue As String)
   ReportId = vNewValue
End Property

Friend Property Get ReportName() As String
   ReportName = ReportId
End Property

Friend Property Get Status() As Long
   Status = ReportStatus
End Property
