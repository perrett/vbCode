VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfirmationControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Type00
   recId As String * 2
   StartSignal As String * 20
   mWare_Ver As String * 3
   mWare_Code As String * 3
   TelNo As String * 20
   transNo As String * 6
   dateTime As String * 12
End Type

Private Type Type80
   recId As String * 2
   ReportID As String * 35
   practiceId As String * 6
   successFlag As String * 1
   finalState As String * 35
   fsDateTime As String * 12
   errMessage As String * 70
   errDateTime As String * 12
End Type

Private Type Type85
   recId As String * 2
   deptCode As String * 3
   transNo As String * 6
   batchDateTime As String * 12
   errMessage As String * 70
   errDateTime As String * 12
End Type

Private Type Type90
   recId As String * 2
   eobSignal As String * 20
End Type

Private Type PMEPAck
   batchStart As Type00
   fileSuccess As Type80
   fileFailure As Type85
   batchEnd As Type90
End Type
   
Private Type MRHdr
   Trigger As String
   dateTime As String
   Control As String
End Type


Private iSoftHdr As Type00
Private iSoftBody1 As Type80
Private iSoftBody2 As Type85
Private iSoftTrail As Type90
Private masterRec As MRHdr
Private fileSrc As String
Private FILENAME As String
Private practiceId As String
Private locTrader As Integer
Private ControlId As Long
Private i As Integer
Private blnRequired As Boolean

Private colReps As New Collection

Friend Sub AddReport(ReportID As String, _
                     Qualifier As Integer)
   Dim repSrc As IceMsgReps
   
   If fileSrc > 0 Then
      Set repSrc = New IceMsgReps
      repSrc.ReportID = Qualifier
      repSrc.ReportName = ReportID
      colReps.Add repSrc
      Set repSrc = Nothing
   End If
End Sub

Private Function ControlRef(LTIndex As Integer) As String
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "SELECT * " & _
            "FROM EDI_Local_Trader_Settings " & _
            "WHERE EDI_LTS_Index = " & locTrader
   RS.Open strSQL, ICEcon, adOpenKeyset, adLockOptimistic
   ControlId = RS!Confirmation_Control
   RS!Confirmation_Control = RS!Confirmation_Control + 1
   RS.Update
   RS.Close
   Set RS = Nothing
   ControlRef = ControlId
End Function

Friend Sub GenerateConfirmationFile(FileLocation As String)
   Dim repSrc As IceMsgReps
   Dim Fname As String
   
   If fileSrc > 0 Then
      If colReps.Count > 0 Then
         Select Case fileSrc
            Case 1
               With masterRec
                  fileData = .Trigger & .dateTime & .Control
               End With
                  
               For i = 1 To colReps.Count
                  Set repSrc = colReps(i)
                  fileData = fileData & "REPORT" & repSrc.ReportName & repSrc.ReportID & Format(Now(), "yyyymmddhhnn")
               Next i
               Set repSrc = Nothing
               
               fileData = fileData & "ENDRFMESS" & ControlId
               Fname = fs.BuildPath(FileLocation, "rfmess." & ControlId)
            
            Case 2
'               fileData = iSoftHdr
               
            
         End Select
         
         Open Fname For Output As #1
         Print #1, fileData
         Close #1
      End If
   End If
End Sub

Friend Sub SetFileSource(SourceId As String, _
                         EDIFile As String)
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim pos As Integer
   
   strSQL = "SELECT * " & _
            "FROM EDI_Local_Trader_Settings " & _
            "WHERE EDI_LTS_Index = " & SourceId
   RS.Open strSQL, ICEcon, adOpenKeyset, adLockOptimistic
   fileSrc = RS!Confirmation_Type
   locTrader = SourceId
   FILENAME = EDIFile
   pos = InStr(1, EDIFile, "_") - 2
   practiceId = Mid(EDIFile, 2, pos)
   RS.Close
   Set RS = Nothing
End Sub

Friend Sub WriteHeader()
   If fileSrc > 0 Then
      Select Case fileSrc
         Case 1
            If masterRec.Trigger = "" Then
               masterRec.Trigger = "STARTRFMESS"
               masterRec.dateTime = Format(Now(), "yyyymmddhhnn")
               masterRec.Control = ControlRef(locTrader)
            End If
            
         Case 2
            With iSoftHdr
               If .recId = "" Then
                  .recId = "80"
                  .StartSignal = "PMEPACK"
                  .transNo = ControlRef(locTrader)
                  .dateTime = Format(Now(), "yyyymmddhhnn")
               End If
            End With
         
      End Select
   End If
End Sub

