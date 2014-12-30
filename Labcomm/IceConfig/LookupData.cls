VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LookupData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Criteria() As String
Private cTot As Integer

Friend Sub AddCriteria(FieldName As String, _
                                     Comparator As String, _
                                     FieldValue As String)
   Criteria(cTot) = FieldName & ": " & Comparator & " :" & FieldValue
   cTot = cTot + 1
End Sub

Friend Function GetCriteria() As String
   Dim i As Integer
   Dim strTemp As String
   Dim strArray() As String
   
   strTemp = " WHERE "
   For i = 0 To cTot - 1
      strArray = Split(Criteria(i), ":")
      strTemp = strTemp & strArray(0) & strArray(1) & strArray(2) & " AND "
   Next i
   strArray = Split(Criteria(i), ":")
   strTemp = strTemp & strArray(0) & strArray(1) & strArray(2)
   GetCriteria = strTemp
End Function

