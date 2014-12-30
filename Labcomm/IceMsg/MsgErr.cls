VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private errFile As String
Private msgTot As Long
Private errStatus As Long

Private colReps As New Collection

Private Sub Class_Initialize()
   errFile = fs.BuildPath(PendingPath, "Run_" & Format(Now(), "yyyymmdd_hhnnss") & ".err")
   Open errFile For Binary As #2
End Sub

Public Sub Add(ByRef fileBuf As String, _
               ByRef RepId As RepData)
   colReps.Add RepId
   Put #2, , fileBuf
   msgTot = msgTot + 1
   errStatus = errStatus Or RepId.Status
End Sub

Private Function AddErrorHeader(ErrorFile As String, _
                                Status As Long)
   Dim iceCmd As New ADODB.Command
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICEMSG_Ins_Service_ImpExp_Headers"
      .Parameters.Append .CreateParameter("recId", adInteger, adParamReturnValue)
      .Parameters.Append .CreateParameter("Organisation", adVarChar, adParamInput, 6, msgControl.Trust)
      .Parameters.Append .CreateParameter("Trader_Code", adVarChar, adParamInput, 15, 0)
      .Parameters.Append .CreateParameter("ImpExp", adChar, adParamInput, 1, "O")
      .Parameters.Append .CreateParameter("ImpExp_File", adVarChar, adParamInput, 120, ErrorFile)
      .Parameters.Append .CreateParameter("Provider_Id", adInteger, adParamInput, , 1)
      .Parameters.Append .CreateParameter("Service_Type", adInteger, adParamInput, , 2)
      .Parameters.Append .CreateParameter("Messages", adInteger, adParamInput, , msgTot)
      .Parameters.Append .CreateParameter("Date_Added", adDate, adParamInput, , Now())
      .Parameters.Append .CreateParameter("Control_Ref", adVarChar, adParamInput, 14, 0)
      .Parameters.Append .CreateParameter("Errors_Found", adInteger, adParamInput, , 1)
      .Parameters.Append .CreateParameter("Comment_Marker", adBoolean, adParamInput, , True)
      .Parameters.Append .CreateParameter("Status_Flag", adInteger, adParamInput, , 1)
      .Parameters.Append .CreateParameter("EDI_LTS_Index", adInteger, adParamInput, , 0)
      .Parameters.Append .CreateParameter("Warning_Flag", adBoolean, adParamInput, , True)
      .Parameters.Append .CreateParameter("HeaderFlag", adInteger, adParamInput, , Status)
      .Parameters.Append .CreateParameter("NatCode", adVarChar, adParamInput, 10, "")
      .Parameters.Append .CreateParameter("FormatIndex", adInteger, adParamInput, , 0)
      .Execute
   End With
   AddErrorHeader = iceCmd.Parameters("RecId")
   Set iceCmd = Nothing
   
End Function

Public Sub WriteFile()
   On Error GoTo procEH
   Dim I As Integer
   Dim RepId As RepData
   Dim repLogId As RepLogData
   Dim errHist As String
   Dim impExpRef As Long
   Dim msgId As Long
   Dim sqlBuf As New StringBuffer
   
   Close #2
   
   If msgTot > 0 Then
      iceCon.BeginTrans

      errHist = fs.BuildPath(ErrorPath, fs.GetFileName(errFile))
      fs.CopyFile errFile, errHist
      impExpRef = AddErrorHeader(errHist, errStatus)
      
      For I = 1 To colReps.Count
         Set RepId = colReps(I)
         
         If (RepId.Status And TF_IS_SUPPRESS) = TF_IS_SUPPRESS Then
            RepId.OrStatus RS_SUPPRESSION
         End If
         
         msgId = msgData.AddImpExpMessage(RepId.EDIIndex, _
                                          impExpRef, _
                                          RepId.Discipline, _
                                          Now(), _
                                          RepId.Status, _
                                          "", _
                                           0, _
                                          "", _
                                          RepId.MessageImpExp, _
                                          RepId.LTSIndex)
         
         LogStatus = &HFFFFFFDE
         
         For Each repLogId In RepId
            If (repLogId.LogStatus And LogStatus) = repLogId.LogStatus Then
               msgData.AddImpExpComment impExpRef, repLogId.LogData, Format(Now(), "dd/mm/yyyy hh:nn"), repLogId.LogStatus, "IceMsg", msgId
            End If
         Next
         
         sqlBuf.Append "DELETE FROM EDI_Rep_List WHERE EDI_Report_Index = " & RepId.EDIIndex & "; "
         
      Next
      
      iceCon.Execute sqlBuf.Value
      
      If blnRollBack = False Then
         iceCon.CommitTrans
      Else
         iceCon.RollbackTrans
      End If
   End If
   
   fs.DeleteFile errFile
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "MsgErr.WriteFile"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
