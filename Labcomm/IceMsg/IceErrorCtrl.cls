VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IceErrorCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private errPend As String

Private colErrRep As New Collection

Public Sub Add(RepId As RepData)
   colErrRep.Add RepId
End Sub

Public Sub CreateFile()
   errPend = fs.BuildPath(PendingPath, "IceMsg_" & Format(Now(), "yymmddhhnnss") & ".err")
End Sub

Public Sub WriteFile()
   Dim errFile As String
   Dim RepId As RepData
   Dim ImpExpRef  As Long
   
   errFile = fs.BuildPath(ErrorPath, fs.GetFileName(errPend))
   ImpExpRef = msgData.AddImpExpHeader("N/A", _
                                       "N/A", _
                                       errFile, _
                                       colErrRep.Count, _
                                       "000000000000000", _
                                       0, _
                                       Now(), _
                                       errStatus, _
                                       -1, _
                                       False, _
                                       "Many", _
                                       "Many")
   strSQL = ""
   
   For i = 1 To colErrRep.Count
      Set RepId = colErrRep(i)
      msgData.AddImpExpMessage RepId.EDIIndex, ImpExpRef, RepId.Discipline, dateAdded, "", RepId.Status
      
      strArray = Split(RepId.LoggedMessages(&HFFFFFFDE), "|")
      For j = 0 To UBound(strArray)
         msgData.AddImpExpComment ImpExpRef, strArray(j), Format(Now(), "dd/mm/yyyy hh:nn"), RepId.Status
      Next j
                  
'     Delete the record from EDI_Rep_List if the conformance target has not been reached
      eClass.FurtherInfo = "Delete after success - report: " & RepId.ReportName & " - index " & RepId.EDIIndex
      strSQL = strSQL & _
               "DELETE FROM EDI_Rep_List " & _
               "WHERE EDI_Report_Index = " & RepId.EDIIndex & "; "
   Next
   
   ICEcon.BeginTrans
   ICEcon.Execute strSQL
'   ICEcon.CommitTrans
   ICEcon.RollbackTrans
   fs.MoveFile errPend, errFile
End Sub
