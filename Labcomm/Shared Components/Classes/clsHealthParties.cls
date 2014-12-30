VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHealthParties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
   
Private clinSpec As String
Private blnNat902 As Boolean
Private gp902Code As String
Private gp902Name As String
Private gp906Code As String
Private gp906Name As String
Private blnNat906 As Boolean
Private indIndex As Long
Private ltIndex As Long
Private downLoadId As Integer
Private PracticeId As String

Public Property Get ClinicianSpeciality() As String
   ClinicianSpeciality = clinSpec
End Property

Public Property Get HP902Name() As String
   HP902Name = Replace(gp902Name, "'", "`")
End Property

Public Property Get HP902Code() As String
   HP902Code = Replace(gp902Code, "'", "`")
End Property

Public Property Get HP902Nat() As Boolean
   HP902Nat = blnNat902
End Property

Public Property Get HP906Name() As String
   HP906Name = Replace(gp906Name, "'", "`")
End Property

Public Property Get HP906Code() As String
   HP906Code = Replace(gp906Code, "'", "`")
End Property

Public Property Get HP906Nat() As Boolean
   HP906Nat = blnNat906
End Property

Public Property Let IndividualIndex(lngNewValue As Long)
   indIndex = lngNewValue
End Property

Public Property Let LTSIndex(lngNewValue As Long)
   ltIndex = lngNewValue
End Property

Public Property Get Practice() As String
   Practice = PracticeId
End Property

Public Property Get DownloadRequest() As Integer
   DownloadRequest = downLoadId
End Property

Public Property Let DownloadRequest(intValue As Integer)
   downLoadId = intValue
End Property

Public Sub Read(ReportIndex As Long, _
                Read906 As Boolean, _
                Optional ValidNatCode As Boolean = True, _
                Optional blnClinNameUseAll As Boolean = False)
'   On Error GoTo procEH
   Dim strSQL As String
   Dim iceCmd As New ADODB.Command
   Dim pLTS As ADODB.Parameter
   Dim pMatch As ADODB.Parameter
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim nat902 As String
   Dim hp902NatCode As String
   
   If indIndex = -1 Then
      strSQL = "SELECT EDI_Individual_Index_To, " & _
                  "Case " & _
                     "When sm.EDI_LTS_Index is null Then sh.EDI_LTS_Index " & _
                     "When sm.EDI_LTS_Index = 0 then sh.EDI_LTS_Index " & _
                     "When sm.EDI_LTS_Index <> sh.EDI_LTS_Index then sm.EDI_LTS_Index " & _
                     "Else sm.EDI_LTS_Index " & _
                  "End As EDI_LTS_Index " & _
               "FROM Service_ImpExp_Messages sm " & _
                  "INNER JOIN Service_ImpExp_Headers sh " & _
                  "ON sm.Service_ImpExp_Id = sh.Service_ImpExp_Id " & _
               "WHERE Service_Report_Index = " & ReportIndex
      RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      ltIndex = RS!EDI_LTS_Index
      
      If downLoadId = 0 Then
         indIndex = RS!EDI_Individual_Index_To
      End If
      
      RS.Close
   End If
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportIndex)
      .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , 0)
   End With
   
'**************************************
'     Evalute GP 902/906 Records
'**************************************
   blnNat902 = False
   blnNat906 = False
   gp906Code = ""
   gp906Name = ""
      
   If indIndex > 0 Then
'     Importer has matched 902
      iceCmd.CommandText = "ICELABCOMM_Report_HP_Individual"
      iceCmd.Parameters("Index").value = indIndex
      
      Set RS = iceCmd.Execute
      clinSpec = RS!Clinician_Speciality_Code & ""
      RS.Filter = "EDI_HP_Type = '902'"
      nat902 = RS!EDI_NatCode & ""
      gp902Name = Trim(RS!EDI_OP_Name & "")
      
      hp902NatCode = RS!EDI_HP_Nat_Code
      
      If (Left(RS!EDI_NatCode, 1) = "G" Or _
         Left(RS!EDI_NatCode, 1) = "C") And _
         IsNumeric(Right(RS!EDI_NatCode, 7)) Or _
         ValidNatCode = False Then
         blnNat902 = True
         gp902Code = RS!EDI_NatCode
      Else
         gp902Code = RS!EDI_Local_Key3 & ""
      End If
      
      RS.Filter = ""
      RS.MoveFirst
      RS.Filter = "EDI_HP_Type = '906'"
      
      If RS.EOF Then
         blnNat906 = blnNat902
         gp906Code = gp902Code
         gp906Name = gp902Name
      
      ElseIf hp902NatCode = RS!EDI_HP_Nat_Code Then
         blnNat906 = blnNat902
         gp906Code = gp902Code
         gp906Name = gp902Name
      Else
         iceCmd.CommandText = "ICELABCOMM_Report_HP_Details"
         iceCmd.Parameters("index").value = ltIndex
         RS.Close
         Set RS = iceCmd.Execute
         
         iceCmd.CommandText = "ICELABCOMM_Report_HP_Key3"
         
         Set RS2 = iceCmd.Execute   '  Read Key3 details for the key1/LTS Index
         RS.Filter = "EDI_HP_Type='906'"
         
         Do Until RS.EOF
            RS2.Find "Practice_Key3 = '" & Replace(RS!Clinician_Local_Code, "'", "''") & "'"
            If Not RS2.EOF Then
'              We have a key 3 match
               gp906Name = Trim(RS2!EDI_OP_Name & "")
               If (Left(RS2!GP_NatCode, 1) = "G" Or _
                  Left(RS2!GP_NatCode, 1) = "C") And _
                  IsNumeric(Right(RS2!GP_NatCode, 7)) Then

'                 This is a valid National Code
                  blnNat906 = True
                  gp906Code = RS!EDI_HP_Nat_Code
               Else
'                 Set the code to the Local Key 3
                  gp906Code = RS2!Practice_Key3
               End If
               Exit Do
               
            End If
            RS.MoveNext
         Loop
         
         RS.MoveFirst
         
         If gp906Code = "" Then
            gp906Name = Trim(RS!Clinician_Surname & "")
            If (Left(RS!EDI_HP_Nat_Code, 1) = "G" Or _
               Left(RS!EDI_HP_Nat_Code, 1) = "C") And _
               IsNumeric(Right(RS!EDI_HP_Nat_Code, 7)) Then
   
'              This is a valid national code - even though is is not matched
               blnNat906 = True
               gp906Code = RS!EDI_HP_Nat_Code
            Else
               If RS.RecordCount > 1 Then
'                 We cannot determine which Local code to use, so use the name
                  gp906Code = Left(RS!Clinician_Surname, 10)
               Else
                  gp906Code = RS!Clinician_Local_Code
               End If
            End If
         End If
      
         RS2.Close
      End If
      
      RS.Close
      
   Else
      '  Extra code added to process downloaded reports - Get the details from Reports_For_Download.
      
      If downLoadId > 0 Then '  Download report
         iceCmd.CommandText = "IceLabcomm_GetDownloadGP"
         iceCmd.Parameters("Index").value = ltIndex
         
         Set RS2 = iceCmd.Execute
         
         PracticeId = RS2!Practice
         gp902Name = RS2!hp902_Name
         gp902Code = RS2!hp902_Code
         
         If (Left(RS2!hp902_Code, 1) = "G" Or _
            Left(RS2!hp902_Code, 1) = "C") And _
            IsNumeric(Right(RS2!hp902_Code, 7)) Then
'              This is a valid National Code
            blnNat902 = True
         Else
            blnNat902 = False
         End If
         
         gp906Name = RS2!hp906_Name
         gp906Code = RS2!hp906_Code
         
         If (Left(RS2!hp906_Code, 1) = "G" Or _
            Left(RS2!hp906_Code, 1) = "C") And _
            IsNumeric(Right(RS2!hp906_Code, 7)) Then
'              This is a valid National Code
            blnNat906 = True
         Else
            blnNat902 = False
         End If
         
      Else
'     Importer hasn't matched, so read Health Party records
         iceCmd.CommandText = "ICELABCOMM_Report_HP_Details"
         iceCmd.Parameters("Index").value = ltIndex
         
         Set RS = iceCmd.Execute '  Read Health Party records
         
         iceCmd.CommandText = "ICELABCOMM_Report_HP_Key3"
         
         Set RS2 = iceCmd.Execute   '  Read Key3 details for the key1/LTS Index
         
         RS.Filter = "EDI_HP_Type = '902'"
         nat902 = RS!EDI_HP_Nat_Code
         
   '     Try to match the Clinician_Local_Code with a EDI_Local_Key3 for this practice.
   '     ICELABCOMM_Report_HP_Key3 returns a list of GP's for the practice with the relevant LTS Index
         Do Until RS.EOF
            RS2.Find "Practice_Key3 = '" & Replace(RS!Clinician_Local_Code & "", "'", "''") & "'"
            If Not RS2.EOF Then
   '           We have a key 3 match
               gp902Name = Trim(RS2!EDI_OP_Name & "")
               
               If (Left(RS2!GP_NatCode, 1) = "G" Or _
                  Left(RS2!GP_NatCode, 1) = "C") And _
                  IsNumeric(Right(RS2!GP_NatCode, 7)) Then
   '              This is a valid National Code
                  blnNat902 = True
                  gp902Code = RS2!GP_NatCode
               Else
   '              Use the Local Key 3
                  gp902Code = RS2!Practice_Key3
               End If
               
               Exit Do
            End If
            RS.MoveNext
            If RS2.EOF = False Then
               RS2.MoveFirst
            End If
         Loop
         
         RS.MoveFirst
         
         If gp902Code = "" Then
   '        No Key 3 found
            gp902Name = Trim(RS!Clinician_Surname & "")
            
            '  Has the clinician name been parsed in the database (and does the query return the fields)?
            If RS.Fields.Count > 4 Then
               If blnClinNameUseAll Then
                   gp902Name = Trim(RS!Clinician_Title & "") & " " & _
                               Trim(RS!Clinician_Forename & "") & " " & _
                               Trim(RS!Clinician_Title & "")
               End If
            End If
            
            If (Left(RS!EDI_HP_Nat_Code, 1) = "G" Or _
               Left(RS!EDI_HP_Nat_Code, 1) = "C") And _
               IsNumeric(Right(RS!EDI_HP_Nat_Code, 7)) Then
   
   '           This is a valid national code - even though is is not matched
               blnNat902 = True
               gp902Code = RS!EDI_HP_Nat_Code
            Else
               If RS.RecordCount > 1 Then
   '              We cannot determine which Local code to use, so use the name
                  gp902Code = Left(RS!Clinician_Surname, 10)
               Else
                  gp902Code = RS!Clinician_Local_Code
               End If
            End If
            
         End If
         
         RS.Filter = ""
         RS.MoveFirst
         RS.Filter = "EDI_HP_Type = '906'"
         
         If RS.EOF Then
   '        No need for further processing - use 902 values
            blnNat906 = blnNat902
            gp906Code = gp902Code
            gp906Name = gp902Name
            
         ElseIf (RS!EDI_HP_Nat_Code = nat902) Or Read906 = False Then
   '        No need for further processing - use 902 values
            blnNat906 = blnNat902
            gp906Code = gp902Code
            gp906Name = gp902Name
         
         Else
         
            RS2.MoveFirst
            
            Do Until RS.EOF
               RS2.Find "Practice_Key3 = '" & Replace(RS!Clinician_Local_Code, "'", "''") & "'"
               If Not RS2.EOF Then
   '              We have a key 3 match
                  gp906Name = Trim(RS2!EDI_OP_Name & "")
                  If (Left(RS2!GP_NatCode, 1) = "G" Or _
                     Left(RS2!GP_NatCode, 1) = "C") And _
                     IsNumeric(Right(RS2!GP_NatCode, 7)) Then
   
   '                 This is a valid National Code
                     blnNat906 = True
                     gp906Code = RS!EDI_HP_Nat_Code
                  Else
   '                 Set the code to the Local Key 3
                     gp906Code = RS2!Practice_Key3
                  End If
                  Exit Do
                  
               End If
               RS.MoveNext
            Loop
            
            RS.MoveFirst
            
            If gp906Code = "" Then
               gp906Name = Trim(RS!Clinician_Surname & "")
               
               '  Has the clinician name been parsed in the database (and does the query return the fields)?
               If RS.Fields.Count > 4 Then
                  If blnClinNameUseAll Then
                      gp906Name = Trim(RS!Clinician_Title & "") & " " & _
                                  Trim(RS!Clinician_Forename & "") & " " & _
                                  Trim(RS!Clinician_Title & "")
                  End If
               End If
            
               If (Left(RS!EDI_HP_Nat_Code, 1) = "G" Or _
                  Left(RS!EDI_HP_Nat_Code, 1) = "C") And _
                  IsNumeric(Right(RS!EDI_HP_Nat_Code, 7)) Then
      
   '              This is a valid national code - even though is is not matched
                  blnNat906 = True
                  gp906Code = RS!EDI_HP_Nat_Code
               Else
                  If RS.RecordCount > 1 Then
   '                 We cannot determine which Local code to use, so use the name
                     gp906Code = Left(RS!Clinician_Surname, 10)
                  Else
                     gp906Code = RS!Clinician_Local_Code
                  End If
               End If
            End If
         End If
         RS2.Close
         RS.Close
      End If '  Report Class = 3
   End If
   
   indIndex = -1
   Set RS2 = Nothing
   Set RS = Nothing
End Sub
