VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HL7_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Reports As ICEReporting.Reports
'Public Report As ICEReporting.Report
'Public ReportList As ICEReporting.ReportList
'Public ICECon As ADODB.Connection
'Public IceRst As ADODB.Recordset
'Public IceRstD As ADODB.Recordset
Private fileName As String
Private outDir As String
Private XMLOutDir As String
Private Extent As String
Private XMLExtent As String
Private CDasNTE As Boolean
Private RITasNTE As Boolean
Private SCasNTE As Boolean
Private ORCOUT As Boolean
Private ANON As Boolean
Private blnXMLReqd As Boolean
'private HPcnt As Integer
'private ClinDetcnt As Integer
'private RepCmtcnt As Integer
'private sampcnt As Integer
'private SampCmtcnt As Integer
'private INVcnt As Integer
'private INVCmtcnt As Integer
'private REScnt As Integer
'private ResCmtcnt As Integer

Private som As String
Private eom As String
Private cs As String
Private RS As String
Private ec As String
Private ss As String
Private fs As String
Private st As String
Private rtfNumLines As Integer
Private orgCode As String
Private SegCnt As Integer
Private outStr() As String
Private procID As String
Private Vers As String
Private istr() As String
Private XMLCnt As Integer
Private RepId() As Integer
Private RepIDCnt As Integer
Private Country As String
Private xmlSender As String
Private xmlSendFacility As String

''User defined types for files section
''Service_Requests
'Type Service_Requests
'    Service_Request_Index As Long
'    Service_Request_ID As String
'    Service_Provider_ID As Long
'    DateTime_Of_Request As Date
'    SRQDateTime_Received As Date
'    SRQPatient_ID As String
'    SRQHospital_Number As String
'    SRQClinical_Information As Boolean
'    SRDate_Added As Date
'    SRQComment_Marker As Boolean
'    SRQDate_Added As Date
'    Request_Type As String
'    SRQImported_To_LIS As String
'End Type
'Public Service_Request As Service_Requests
''
'
'Type Service_Requests_Comments
'    Request_Comment_Index As Long
'    Service_Request_Index As Long
'    Service_Request_Comment As String
'End Type
'Public Service_Request_Comment() As Service_Requests_Comments
'Public Num_Request_Comments As Long
'
'
'Type Service_Invest_Requests
'    Invest_Req_Index As Long
'    Service_Request_Index As Long
'    Sample_Id As String
'    Investigation_Requested As String
'    Specialty As String
'    Date_Added As Date
'End Type
'
'Public Service_Invest_Request() As Service_Invest_Requests
'Public Num_Request_Invest As Long
'
''Patient_Local_IDs
'Type Patient_Local_Ids
'    Patient_Id_Key As Long
'    Patient_Local_ID As String
'    PLIOrg_Code As String
'    PLIDate_Added As Date
'End Type
'Public Patient_Local_ID As Patient_Local_Ids
''
''Patient
'Type Patients
'    Patient_Id_Key As Long
'    New_NHS_No As String
'    Old_NHS_No As String
'    Date_Of_Birth As Date
'    Date_Of_Birth_Format As Integer
'    Registered_GP As String
'    Registered_Practice As String
'    Sex As Long
'    Ethnic_Group As String
'    Marital_Status As Integer
'    Title As String
'    Surname As String
'    Midname As String
'    Forename As String
'    Pat_Addr_Line1 As String
'    Pat_Addr_Line2 As String
'    Pat_Addr_Line3 As String
'    Pat_Addr_Line4 As String
'    Pat_Addr_Line5 As String
'    Temporary_Address_Flag As Boolean
'    Comment_Marker As Boolean
'    Pat_Postcode As String
'    Pat_Tel_No As String
'    Date_Added As Date
'End Type
'Public Patient As Patients
'
''Patient Event History
'Type Patient_Event_History
'    Pat_Event_Index As Long
'    Pat_Msg_Type As String
'    Pat_Event_Type As String
'    Pat_Service_Type As Integer
'    Pat_Location As String
'    Pat_Event_ID As String
'    Pat_Local_ID As String
'    Pat_Status As String
'    Pat_Event_Date As Date
'    Date_Added As Date
'End Type
'Public Pat_Event_History As Patient_Event_History
'
''
''Clinical Details
'Type Clinical_Details
'    Clin_Detail_Key As String
'    Free_Text As String
'    Detail_Title As String
'    CLDDateTime_of_Detail As Date
'    CLDComments_Marker As Boolean
'    CLDDate_Added As Date
'End Type
'Public Clinical_Detail() As Clinical_Details
'Public Num_Clinical_Details As Long
''
''Service_Samples
'Type Service_Samples
'    Sample_Index As Long
'    Service_Request_Index As Long
'    SSSample_ID As String
'    Sample_Code As String
'    Sample_Text As String
'    Collection_Code As String
'    Collection_Text As String
'    AnatOrig_Code As String
'    Collection_DateTime As Date
'    Collection_Amount As String
'    Collection_Units As String
'    Collection_Tube As String
'    Collection_DateTimeReceived As Date
'    SSComment_Marker As Boolean
'    SSDate_Added As Date
'End Type
'Public Service_Sample() As Service_Samples
'Public Samplelk(60, 60) As Long
'
'Type Service_Samples_Comments
'    Sample_Comment_Index As Long
'    Sample_Index As Long
'    Sample_Comment As String
'End Type
'Public Service_Sample_Comment(30, 60) As Service_Samples_Comments
'Public Num_Sample_Comments(30) As Long
'
'
''
''Service_Health_Parties
'Type Service_Health_Parties
'    EDI_HP_Type As String
'    EDI_NatCode As String
'    EDI_HP_Nat_Code As String
'    Clinician_National_Code As String
'    Clinician_Local_Code As String
'    Clinician_Surname As String
'    Clinician_Specialty_Code As String
'    EDI_OP_Name As String
'End Type
'Public EDI_Health_Party(10) As Service_Health_Parties
'
''Service_Reports
'Type Service_Reports
'    Service_Report_Index As Long
'    Service_Report_ID As String
'    Service_Request_Index As Long
'    Service_Report_Type As String
'    SRDateTime_Of_Report As Date
'    Urgency As String
'    Status As String
'    Auth_DateTime As Date
'    SRComment_Marker As Boolean
'    SRDate_Added As Date
'    Alert As String
'    Review_Level As Integer
'    Comment_Level As Integer
'End Type
'Public Service_Report As Service_Reports
'Public Num_Service_Reports As Long
''
''Report_Comments
'Type Report_Comments
'    Comment As String
'    Type As String
'End Type
'Public Report_Comment() As Report_Comments
'Public Num_Report_Comments As Long
''
''Service_Investigations
'Type Service_Investigations
'    Investigation_Index As Long
'    Service_Report_Index As Long
'    Investigation_Code As String
'    Investigation_Requested As String
'    Sample_Index As Long
'    SIComment_Marker As Boolean
'    SIDate_Added As Date
'End Type
'Public Service_Investigation() As Service_Investigations
'Public Num_Service_Investigations As Long
''
''Service_Comments
'Type Service_Invest_Comments
'    Comment_Index As Long
'    Comment_Orig_Index As Long
'    Comment_Orig_type As String
'    Comment As String
'End Type
'Public Service_Invest_Comment(60, 160) As Service_Invest_Comments
'Public Num_Service_Invest_Comments(60) As Long
''
''Service_Results
'Type Service_Results
'    Result_Index As Long
'    Investigation_Index As Long
'    Result_Type As String
'    Sample_Index As Long
'    Result_Code As String
'    Result_Rubric As String
'    Comparator As String
'    Deviation As String
'    Abnormal_Flag As Boolean
'    Abnormal_Code As String
'    Result As String
'    UOM_Code As String
'    UOM_Text As String
'    Status As String
'    RESDate_Added As Date
'    RESComment_Marker As Boolean
'End Type
'Public Service_Result(60, 60) As Service_Results
'Public Num_Service_Results(60) As Long
''
''Service_Comments
'Type Service_Comments
'    Comment_Index As Long
'    Comment_Orig_Index As Long
'    Comment_Orig_type As String
'    Comment_Type As String
'    Comment As String
'End Type
'Public Service_Comment(60, 60, 200) As Service_Comments
'Public Num_Service_Comments(60, 60) As Long
''
''Service_Ranges
'Type Service_Ranges
'    Range_Index As Long
'    Result_Index As Long
'    Lower_range As String
'    Upper_Range As String
'    Type As String
'    Unit_Code As String
'    Unit_Text As String
'    RANComment_Marker As Boolean
'    RANDate_Added As Date
'End Type
'Public Service_Range(60, 60) As Service_Ranges
'Public Num_Service_Ranges(60, 60) As Long
'
''Service_ImpExp_Headers
'Type Service_ImpExp_Headers
'    Service_ImpExp_ID As Long
'    Organisation As String
'    Imp_Exp As String
'    ImpExp_File As String
'    Provider_ID As Integer
'    Service_Type As Integer
'    Messages As Long
'    Date_Added As Date
'    Error_Found As Long
'    Comment_Marker As Boolean
'    Warning_Flag As Boolean
'    File_Identifier As String
'    EDI_LTS_Index As Long
'End Type
'Public Service_ImpExp_Header As Service_ImpExp_Headers
'Public blnDoneImpExpHeader As Boolean
'
''Service_ImpExp_Messages
'Type Service_ImpExp_Messages
'    Service_ImpExp_Message_ID As Long
'    Service_ImpExp_ID As Long
'    Service_Message_ID As String
'    Patient_Local_ID As String
'    Service_ID As String
'    Patient_Name As String
'    Patient_Sex As Integer
'    Patient_DoB As Date
'    Service_Date_of_Service As Date
'    Service_Message_Comment As String
'    Rejected_Count As Integer
'    Rejected_File As String
'    Date_Added As Date
'    Destination As String
'    Service_Report_Index As Long
'    Discipline As Long
'    To_EDI As Boolean
'    Report_Identifier As String
'    EDI_Individual_Index_To As Long
'End Type
'Public Service_ImpExp_Message As Service_ImpExp_Messages
'Public Num_ImpExp_Messages As Long
'
''Service_ImpExp_Comments
'Type Service_ImpExp_Comments
'    Service_ImpExp_Comment_ID As Long
'    Service_ImpExp_ID As Long
'    Service_ImpExp_Comment As String
'    Service_ImpExp_Process As String
'    Code As Integer
'    Date_Added As Date
'End Type
'Public Service_ImpExp_Comment() As Service_ImpExp_Comments
'Public Num_ImpExp_Comments As Long
'
'Type EDI_Rep_Lists
'  EDI_Report_Index As Long
'  EDI_Provider_Org As String
'  EDI_Loc_Nat_Code_To As String
'  EDI_Individual_Index_To As Long
'  EDI_Service_ID  As String
'  Date_Added As Date
'  EDI_LTS_Index As Long
'  EDI_Rep_Specialty As Long
'End Type
'Public EDI_Rep_List As EDI_Rep_Lists
'

Private Sub Class_Initialize()
   Dim iVar As String
   
   If Read_Ini_Var("Data", "ClinDets", INIFile) = "OBR" Then CDasNTE = False
   If Read_Ini_Var("Data", "TextResults", INIFile) = "OBX" Then RITasNTE = False
   If Read_Ini_Var("Data", "SampleComments", INIFile) = "OBR" Then SCasNTE = False
   If Read_Ini_Var("Data", "OutputORC", INIFile) = "1" Then ORCOUT = True
   If Read_Ini_Var("Data", "Anon", INIFile) = "1" Then ANON = True
   
   iVar = Read_Ini_Var("Control", "Start", INIFile)
   If IsNumeric(iVar) Then
      som = Chr(iVar)
   Else
      som = Chr(11)
   End If
   
   iVar = Read_Ini_Var("Control", "End", INIFile)
   If IsNumeric(iVar) Then
      eom = Chr(iVar)
   Else
      eom = Chr(28)
   End If
   
   iVar = Read_Ini_Var("Control", "Escape", INIFile)
   If IsNumeric(iVar) Then
      ec = Chr(iVar)
   Else
      ec = Chr(92)
   End If
   
   iVar = Read_Ini_Var("Control", "Subcomponent", INIFile)
   If IsNumeric(iVar) Then
      ss = Chr(iVar)
   Else
      ss = Chr(38)
   End If
   
   iVar = Read_Ini_Var("Control", "Field", INIFile)
   If IsNumeric(iVar) Then
      fs = Chr(iVar)
   Else
      fs = Chr(124)
   End If
   
   iVar = Read_Ini_Var("Control", "Component", INIFile)
   If IsNumeric(iVar) Then
      cs = Chr(iVar)
   Else
      cs = Chr(94)
   End If
   
   iVar = Read_Ini_Var("Control", "Repetition", INIFile)
   If IsNumeric(iVar) Then
      RS = Chr(iVar)
   Else
      RS = Chr(126)
   End If
   
   procID = Read_Ini_Var("Control", "ProcID", INIFile)
   Vers = Read_Ini_Var("Control", "Version", INIFile)
   Country = Read_Ini_Var("Control", "Country", INIFile)
   xmlSender = Read_Ini_Var("XML", "Sender", INIFile)
   xmlSendFacility = Read_Ini_Var("XML", "SenderFacility", INIFile)
   
   st = vbCr & vbLf
   rtfNumLines = 0
End Sub

Public Sub Output_HL7_Array()
Dim cnt As Integer
Dim I As Integer

'              FILENAME = Format(Date, "DD") + Format(Time, "HHMMSS")
'              FILENAME = Trim(outDir) + "\" + Trim(FILENAME) + "." + Trim(Extent)
'              cnt = 1
'              While Dir(FILENAME) > ""
'                FILENAME = Format(Date, "DD") + Format(Time, "HHMMSS") + Trim(Str(cnt))
'                FILENAME = Trim(outDir) + "\" + Trim(FILENAME) + "." + Trim(Extent)
'                cnt = cnt + 1
'              Wend
'              Open FILENAME For Output As #10
'                For i = 1 To SegCnt
'                    Print #10, outstr(i)
'                    'rtfDisplay.Text = rtfDisplay.Text & vbCrLf & Outstr(i)
'                    'rtnumlines = rtfNumLines + 1
'                    If Len(frmMain.rtfDisplay.Text) > 4000 Then frmMain.rtfDisplay.Text = ""
'                    frmMain.rtfDisplay.Text = frmMain.rtfDisplay.Text & Format(Date, "short date") & " " & Format(Time, "hh:mm") & " " & outstr(i) & Chr(10)
'                    frmMain.rtfDisplay.SelStart = Len(frmMain.rtfDisplay.Text)
'                    frmMain.rtfDisplay.Refresh
'                Next i
'              'FileOpen = True
'              'AddMessage ("Open " + FILENAME + " for output")
'              'Print #10, Obuffer;
'              Close #10
'              'AddMessage "Message Saved File " + FILENAME
'              'Obuffer = ""

End Sub

Public Sub Output_XML_Array()
Dim cnt As Integer
Dim I As Integer

              fileName = Format(Date, "DD") + Format(Time, "HHMMSS")
              fileName = Trim(XMLOutDir) + "\" + Trim(fileName) + "." + Trim(XMLExtent)
              cnt = 1
              While Dir(fileName) > ""
                fileName = Format(Date, "DD") + Format(Time, "HHMMSS") + Trim(Str(cnt))
                fileName = Trim(XMLOutDir) + "\" + Trim(fileName) + "." + Trim(XMLExtent)
                cnt = cnt + 1
              Wend
              Open fileName For Output As #10
                For I = 1 To XMLCnt
                    Print #10, istr(I)
                    'rtfDisplay.Text = rtfDisplay.Text & vbCrLf & Outstr(i)
                    'rtnumlines = rtfNumLines + 1
'                    If Len(frmMain.rtfDisplay.Text) > 4000 Then frmMain.rtfDisplay.Text = ""
'                    frmMain.rtfDisplay.Text = frmMain.rtfDisplay.Text & Format(Date, "short date") & " " & Format(Time, "hh:mm") & " " & iStr(i) & Chr(10)
'                    frmMain.rtfDisplay.SelStart = Len(frmMain.rtfDisplay.Text)
'                    frmMain.rtfDisplay.Refresh
                Next I
              'FileOpen = True
              'AddMessage ("Open " + FILENAME + " for output")
              'Print #10, Obuffer;
              Close #10
              'AddMessage "Message Saved File " + FILENAME
              'Obuffer = ""

End Sub

Public Function FileHeader() As String
   Dim mshHeader As String
   Dim xmlHeader As String
   Dim tStr() As String
   Dim uStr() As String
   Dim I As Integer
   Dim j As Integer
   Dim n As Integer
   
   If blnXMLReqd Then

      If xmlSender = "" Then
         xmlSender = "ICE"
      End If
      
      xmlSendFacility = Read_Ini_Var("XML", "SenderFacility", INIFile)
      If xmlSendFacility = "" Then
         xmlSendFacility = orgCode
      Else
         uStr = Split(xmlSendFacility, "+")
      End If
      
      mshHeader = "MSH" & fs & cs & RS & ec & ss & fs & xmlSender & fs & uStr(0)
      
      For j = 1 To UBound(uStr)
         mshHeader = mshHeader & cs & uStr(j)
      Next j
      
      mshHeader = mshHeader & fs _
                  & fs _
                  & msgControl.NationalCode & fs _
                  & Mid$(Str$(Date), 7, 4) & Mid$(Str$(Date), 4, 2) & Left$(Str$(Date), 2) & Left$(Str$(Time), 2) & Mid$(Str$(Time), 4, 2) & fs _
                  & fs _
                  & "ORU" & cs & "R01" & fs _
                  & Trim(Str(msgControl.CurrentInterchange)) & fs _
                  & Trim(procID) & fs _
                  & Vers
            
      xmlHeader = "<MSH><MSH.1>" & fs & "</MSH.1>"
      
      mshHeader = Replace(mshHeader, "&", "&amp;")
      mshHeader = Replace(mshHeader, "<", "&lt;")
      mshHeader = Replace(mshHeader, ">", "&gt;")
      mshHeader = Replace(mshHeader, Chr(34), "&quot;")
      mshHeader = Replace(mshHeader, "'", "&apos;")
      
      tStr = Split(mshHeader, fs, -1)
      
      For j = 1 To UBound(tStr)
         n = j + 1
         tStr(j) = Replace(tStr(j), "<", "&lt;")
         If Trim(tStr(j)) <> "" Then
            xmlHeader = xmlHeader & "<MSH." & Trim(Str(n)) & ">"
            
            Select Case n
               Case 2, 10
                  xmlHeader = xmlHeader & tStr(j)
                      
               Case 3, 5, 6
                  xmlHeader = xmlHeader & vbCrLf & "<HD.1>" & tStr(j) & "</HD.1>" & vbCrLf
                      
               Case 4
                  uStr = Split(tStr(j), cs)
                  xmlHeader = xmlHeader & vbCrLf & "<HD.1>" & uStr(0) & "</HD.1>" & vbCrLf
                  If UBound(uStr) > 1 Then
                     For I = 1 To UBound(uStr)
                        xmlHeader = xmlHeader & vbCrLf & "<HD." & I + 1 & ">" & uStr(I) & "</HD." & I + 1 & ">" & vbCrLf
                     Next I
                  End If
                     
               Case 7
                  xmlHeader = xmlHeader & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
               
               Case 9
                  uStr = Split(tStr(j), cs, -1)
                  If Trim(uStr(0)) <> "" Then
                     xmlHeader = xmlHeader & vbCrLf & "<MSG.1>" & uStr(0) & "</MSG.1>"
                  End If
                  
                  If UBound(uStr) > 0 Then
                     If Trim(uStr(1)) <> "" Then
                        xmlHeader = xmlHeader & vbCrLf & "<MSG.2>" & uStr(1) & "</MSG.2>" & vbCrLf
                     End If
                  End If
               
               Case 11
                  xmlHeader = xmlHeader & vbCrLf & "<PT.1>" & tStr(j) & "</PT.1>" & vbCrLf
                      
               Case 12
                  xmlHeader = xmlHeader & vbCrLf & "<VID.1>2.4</VID.1>" & vbCrLf
                  
            End Select
            
            xmlHeader = xmlHeader & "</MSH." & Trim(Str(n)) & ">"
         End If
      Next j
            
      FileHeader = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf & _
                   "<ORU_R01 xmlns=" & Chr(34) & "urn:hl7-org:v2xml" & Chr(34) & ">" & xmlHeader & "</MSH>"
   
   Else
      FileHeader = som
   End If
End Function

Public Function FileTrailer() As String
   If blnXMLReqd Then
      FileTrailer = "</ORU_R01>"
   Else
      FileTrailer = eom & st
   End If
End Function

'Public Function Get_LocalCode(Local_Rubric As String, Sample_Code As String) As String
'Get_LocalCode = ""
'strSQL = "SELECT EDI_Local_Test_Code FROM EDI_InvTest_Codes WHERE EDI_Local_Rubric = '" & Local_Rubric & "'" _
'         & " AND EDI_Sample_TypeCode = '" & Sample_Code & "'"
' Set IceRstD = New ADODB.Recordset
' IceRstD.Open strSQL, ICEcon, adOpenKeyset, adLockReadOnly
' If Not (IceRstD.BOF And IceRstD.EOF) Then Get_LocalCode = Trim(IceRstD.Fields("EDI_Local_test_Code") & "")
' IceRstD.Close
'End Function

Public Function Get_DiscCode(txtDiscipline As String) As String
   Dim RS As New ADODB.Recordset

   Get_DiscCode = txtDiscipline
   strSQL = "SELECT Specialty_Text " & _
            "FROM Service_Discipline_Map " & _
            "WHERE Specialty_Code = '" & txtDiscipline & " '"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.EOF Then
      Get_DiscCode = txtDiscipline
   Else
      Get_DiscCode = Trim(RS!Specialty_Text & "")
   End If
   
   RS.Close
   Set RS = Nothing
End Function

Public Function GetSampleLab(SampleText As String, _
                             Optional ReturnDiscipline As Boolean = False) As String
   Dim RS As New ADODB.Recordset
   
   strSQL = "SELECT * " & _
            "FROM Service_Sample_Labs " & _
            "WHERE Left(Sample_LabId,"
   
   If Left(SampleText, 1) = "M" Or _
      Left(SampleText, 1) = "K" Then
      strSQL = strSQL & "1) = '" & Left(SampleText, 1) & "'"
   Else
      strSQL = strSQL & "2) = '" & Left(SampleText, 2) & "'"
   End If
   
   RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
   
   If RS.EOF Then
      GetSampleLab = ""
   Else
      If ReturnDiscipline Then
         GetSampleLab = RS!Discipline_Code
      Else
         GetSampleLab = RS!Lab_Expansion
      End If
   End If
   
   RS.Close
   Set RS = Nothing
End Function


Public Property Get InvalidChars() As Variant
   ReDim strArray(2)
   strArray(0) = ""
   strArray(1) = ""
   InvalidChars = strArray
End Property

Public Property Get Notification() As Boolean
   Notification = False
End Property

Public Property Get ReleaseChar() As String
   ReleaseChar = ""
End Property

Public Property Get RetainGrave() As Boolean
   RetainGrave = False
End Property

Public Sub ProcessReport(ID As Long)
   On Local Error GoTo procEH:
   Dim iceCmd As New ADODB.Command
   Dim cRS As New ADODB.Recordset
'   Dim RS As New ADODB.Recordset
   Dim sampRS As New ADODB.Recordset
   Dim invRS As New ADODB.Recordset
   Dim resRS As New ADODB.Recordset
   Dim segCheck As Long
   Dim blnAbnormal As Boolean
   Dim istr As String
   Dim jstr As String
   Dim Requestor As String
   Dim Astr As String
   Dim RecipOrg As String
   Dim Recipient As String
   Dim ReqOrg As String
   Dim DiscCode As String
   Dim InitRepLine As Boolean
   Dim InitDetLine As Boolean
   Dim InitSampLine As Boolean
   Dim InitRes As Boolean
   Dim ResultCmt As String
   Dim rCode As String
   Dim rRubric As String
   Dim I As Long
   Dim ClinDets As String
   Dim nat902 As String
   Dim gp902Code As String
   Dim gp902Name As String
   Dim gp906Code As String
   Dim gp906Name As String
   Dim blnNat902 As Boolean
   Dim blnNat906 As Boolean
   Dim SampleCmts As String
   Dim ReportCmts As String
   Dim ResultValue As String
   Dim NTEcnt As Integer
   Dim OBXcnt As Integer
   Dim repClin As String
   Dim strTemp As String
   Dim P_nhs_No As String
   Dim P_Surname As String
   Dim P_Midname As String
   Dim P_Forename As String
   Dim P_Pat_Addr_Line1 As String
   Dim P_Pat_Addr_Line2 As String
   Dim P_Pat_Addr_Line3 As String
   Dim P_Pat_Addr_Line4 As String
   Dim P_Pat_Addr_Line5 As String
   Dim P_Pat_PostCode As String
   Dim Status As enumInvResStatus
   Dim vData As Variant
   Dim rData As Variant
   Dim vFields As Variant
   Dim testDesc As String
   Dim testCount As Integer
   Dim rcIndex As Long
   Dim iPosn As Integer
   Dim strArray() As String
   Dim rPosn As Integer
   Dim ServiceId As String
   Dim datStr As String
   Dim rangeStr As String
   Dim obr2Str As String
   Dim PIDStr As String
   Dim strSQL As String
   Dim hpData As New clsHealthParties

   ServiceId = repListRS!Service_Id
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ID)
      'Set RS = .Execute
   End With
   
   RecipOrg = ""
   Recipient = ""
   Requestor = ""
   ReqOrg = ""
   
'**************************************
'  Evalute GP 902/906 Records
'**************************************
   hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
   hpData.LTSIndex = repListRS!EDI_LTS_Index
   hpData.Read ID, True
   
   RecipOrg = repListRS!EDI_Loc_Nat_Code_To
   
   Recipient = hpData.HP902Code
   repClin = hpData.HP902Name
   
   orgCode = msgControl.Organisation
   
   Requestor = hpData.HP906Code
             
   strSQL = "SELECT * " & _
            "FROM EDI_Health_Parties " & _
            "WHERE EDI_Report_Index = " & ID & _
               " AND EDI_HP_Type = '907'"
   cRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   If Not cRS.EOF Then
'     Requester Location
      ReqOrg = Mid(cRS!EDI_HP_Nat_Code, 7, 8)
   End If
   
   cRS.Close
   
   If ReqOrg = "" Then
      If Requestor = Recipient Then
         ReqOrg = RecipOrg
      Else
         ReqOrg = orgCode
      End If
   End If
   
   eClass.FurtherInfo = "Get Discipline Code"
   If blnUseSampleForDisc Then
      DiscCode = ""
   Else
      DiscCode = Get_DiscCode(repListRS!service_Report_Type)
   End If
'   DiscCode = IIf(IsNull(repListRS!Discipline_Text), repListRS!Service_Report_Type, repListRS!Discipline_Text)
   ReDim outStr(0)
'   SegCnt = SegCnt + 1
   ReDim Preserve outStr(0 To SegCnt)
   
   If xmlSender = "" Then
      xmlSender = "ICE"
   End If
   
   xmlSendFacility = Read_Ini_Var("XML", "SenderFacility", INIFile)
   If xmlSendFacility = "" Then
      xmlSendFacility = orgCode
   End If
   
   datStr = Format(Date, "dd/mm/yyyy")
   outStr(SegCnt) = "MSH" & fs & cs & RS & ec & ss & fs & xmlSender & fs & xmlSendFacility & fs _
                          & fs & RecipOrg & fs & Mid$(datStr, 7, 4) & Mid$(datStr, 4, 2) & Left$(datStr, 2) & Left$(Str$(Time), 2) & Mid$(Str$(Time), 4, 2) _
                          & fs & fs & "ORU" & cs & "R01" & fs & Trim(Str(ID)) & fs & Trim(procID) & fs & Vers
   SegCnt = SegCnt + 1
   
   
   eClass.FurtherInfo = "Patient Details"
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Patient"
      Set cRS = .Execute
   End With
   
   ReDim Preserve outStr(0 To SegCnt)
   
   istr = Format(cRS!Date_Of_Birth, "yyyymmdd")
   Astr = ""
   
'   If Trim(cRS!Hospital_Number & "") <> "" Then
'      Astr = Trim(cRS!Hospital_Number) & cs & cs & cs & cs & "MRN"
'   End If
   If msgControl.AnonymizeData Then
      If (cRS!New_NHS_No <> "") Then
         P_nhs_No = "9999999603"
      End If
      
      P_Surname = "Surname"
      P_Midname = "M"
      P_Forename = "Forename"
      P_Pat_Addr_Line1 = "Address 1"
      P_Pat_Addr_Line2 = "Address 2"
      P_Pat_Addr_Line3 = "Address 3"
      P_Pat_Addr_Line4 = "Address 4"
      P_Pat_Addr_Line5 = "Address 5"
      P_Pat_PostCode = "PC1 1CP"
   Else
      P_nhs_No = Trim(cRS!New_NHS_No & "")
      P_Surname = cRS!Surname
      P_Midname = cRS!Midname
      P_Forename = cRS!Forename
      P_Pat_Addr_Line1 = cRS!Pat_Addr_Line1 & ""
      P_Pat_Addr_Line2 = cRS!Pat_Addr_Line2 & ""
      P_Pat_Addr_Line3 = cRS!Pat_Addr_Line3 & ""
      P_Pat_Addr_Line4 = cRS!Pat_Addr_Line4 & ""
      P_Pat_Addr_Line5 = cRS!Pat_Addr_Line5 & ""
      P_Pat_PostCode = cRS!Pat_PostCode & ""
   End If
        
   If XMLMRNPrefix = "" Then
      If Trim(cRS!Hospital_Number & "") <> "" Then
         If UCase(Left(Trim(cRS!Hospital_Number), 3)) = "APX" Then
            Astr = Trim(cRS!Hospital_Number) & cs & cs & cs & cs & "LAB"
         Else
            Astr = Trim(cRS!Hospital_Number) & cs & cs & cs & cs & "MRN"
         End If
      End If
   
      If Trim(P_nhs_No & "") <> "" Then
         If Country = "IE" Then
             Astr = Astr & RS & Trim(P_nhs_No) & cs & cs & cs & cs & "MRN"
         Else
             Astr = Astr & RS & Trim(P_nhs_No) & cs & cs & cs & cs & "NHS"
         End If
      End If
   Else
      PIDStr = IIf(P_nhs_No = "", cRS!Hospital_Number, P_nhs_No)
      
      If PIDStr <> "" Then
         If InStr(1, XMLMRNPrefix, Left(PIDStr, 1), vbTextCompare) > 0 Then
            Astr = Trim(PIDStr) & cs & cs & cs & cs & "MRN"
         ElseIf InStr(1, XMLLABPrefix, Left(PIDStr, 1), vbTextCompare) > 0 Then
            Astr = Trim(PIDStr) & cs & cs & cs & cs & "LAB"
         End If
      End If
   End If
   
   If Astr = "" Then
      Astr = ServiceId & cs & cs & cs & cs & "MRN"
   End If
   
   Select Case cRS!Sex
      Case 0
         strTemp = "U"
      
      Case 1
         strTemp = "M"
         
      Case 2
         strTemp = "F"
         
      Case 9
         strTemp = "S"
   End Select
      
   outStr(SegCnt) = "PID" & fs & _
                    "1" & fs & _
                    fs & _
                    Astr & fs & _
                    fs & _
                    Trim(P_Surname & "") & cs & Trim(P_Forename & "") & cs & Trim(P_Midname & "") & _
                    cs & cs & _
                    cRS!Title & cs & cs & "L" & fs & _
                    fs & _
                    Format(cRS!Date_Of_Birth, "yyyymmdd") & fs & _
                    strTemp & fs & _
                    fs & _
                    fs & _
                    Trim(P_Pat_Addr_Line1 & "") & cs & _
                    Trim(P_Pat_Addr_Line2 & "") & cs & _
                    Trim(P_Pat_Addr_Line3 & "") & cs & _
                    Trim(P_Pat_Addr_Line4 & "") & cs & _
                    Trim(P_Pat_Addr_Line5 & "") & fs & _
                    Trim(P_Pat_PostCode & "") & fs
     
   ClinDets = ""
   SampleCmts = ""
   ReportCmts = ""
   NTEcnt = 0
   InitRepLine = False
   InitDetLine = False
   InitSampLine = False
     
   cRS.Close
   
   eClass.FurtherInfo = "Report Level Comments"
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Comments"
      Set cRS = .Execute
   End With
      
   Do Until cRS.EOF
      Select Case cRS!Comment_Type
         Case "R"
            SegCnt = SegCnt + 1
            NTEcnt = NTEcnt + 1
            ReDim Preserve outStr(0 To SegCnt)
            outStr(SegCnt) = "NTE" & fs & Trim(Str(NTEcnt)) & fs & "C" & fs & cRS!Service_Report_Comment
            InitRepLine = True
            
         Case "D"
                'Clinical Details
            If CDasNTE Then
               SegCnt = SegCnt + 1
               NTEcnt = NTEcnt + 1
               ReDim Preserve outStr(0 To SegCnt)
               outStr(SegCnt) = "NTE" & fs & Trim(Str(NTEcnt)) & fs & "C" & fs & cRS!Service_Report_Comment
               InitRepLine = True
            Else
               If ClinDets = "" Then
                  ClinDets = cRS!Service_Report_Comment
               Else
                   ClinDets = ClinDets & "\.br\" & cRS!Service_Report_Comment
               End If
            End If
         
         Case "S"
            'Sample Comments
            If SCasNTE Then
               SegCnt = SegCnt + 1
               NTEcnt = NTEcnt + 1
               ReDim Preserve outStr(0 To SegCnt)
               outStr(SegCnt) = "NTE" & fs & Trim(Str(NTEcnt)) & fs & "C" & fs & cRS!Service_Report_Comment
               InitRepLine = True
            Else
               If SampleCmts = "" Then
                   SampleCmts = cRS!Service_Report_Comment
               Else
                  SampleCmts = SampleCmts & "\.br\" & cRS!Service_Report_Comment
               End If
            End If
                
      End Select
      cRS.MoveNext
   Loop
    
   eClass.FurtherInfo = "Requestor for PV1"
   SegCnt = SegCnt + 1
   ReDim Preserve outStr(0 To SegCnt)
    
   outStr(SegCnt) = "PV1" & fs & _
                    fs & _
                    "O" & fs & _
                    ReqOrg & fs & _
                    fs & _
                    fs & _
                    fs
   If Requestor <> "" Then
      outStr(SegCnt) = outStr(SegCnt) _
                       & Requestor & cs & repClin & fs _
                       & Requestor & cs & repClin & fs
   Else
      outStr(SegCnt) = outStr(SegCnt) _
                       & cs & repClin & fs _
                       & cs & repClin & fs
   End If
   
   If blnIncludeRepId Then
      For I = 9 To 18
         outStr(SegCnt) = outStr(SegCnt) & fs
      Next I
      outStr(SegCnt) = outStr(SegCnt) & ID
   End If
   
   I = 0
   
   If ORCOUT Then
      SegCnt = SegCnt + 1
      ReDim Preserve outStr(0 To SegCnt)
      outStr(SegCnt) = "ORC" & fs _
                       & "RE" & fs _
                       & fs _
                       & ServiceId & fs _
                       & fs _
                       & fs _
                       & fs _
                       & fs _
                       & fs _
                       & fs _
                       & fs _
                       & fs & Requestor & cs & repClin & fs _
                       & Trim(ReqOrg)
   End If
        
   cRS.Close
   
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Sample"
      Set sampRS = .Execute
   End With
   
   If DiscCode = "" Then
      DiscCode = GetSampleLab(sampRS!Sample_Id, True)
   End If
   
   segCheck = SegCnt
   Do Until sampRS.EOF
      With iceCmd
         If blnUseRCIndex Then
            .CommandText = "ICELABCOMM_Report_Invest_By_Index"
         Else
            .CommandText = "ICELABCOMM_Report_Invest_By_Code"
         End If
         
         .Parameters(0).Value = ID
         .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
         .Parameters.Append .CreateParameter("Samp", adInteger, adParamInput, , sampRS!Sample_Index)
         .Parameters.Append .CreateParameter("Ab", adBoolean, adParamOutput, , blnAbnormal)
         Set invRS = .Execute
         .Parameters.Delete ("Ab")
         .Parameters.Delete ("Samp")
         .Parameters.Delete ("LTSIndex")
      End With
      
'      If IsNull(invRS!EDI_RC_Index) Then
'         vFields = Array(9, 10, 12, 13, 18, 19, 20, 21, 22, 23, 8)
'      Else
'         vFields = Array(14, 15, 16, 17, 24, 25, 26, 27, 28, 29, 8)
'      End If
'
'      iPosn = 0
'      vData = invRS.GetRows(-1, adBookmarkFirst, vFields)
      invRS.MoveFirst
      
      Do Until invRS.EOF
         eClass.FurtherInfo = "Sample Index = " & sampRS!Sample_Index & " Investigation Index = " & invRS!Investigation_Index
         If IsNull(invRS!EDI_Local_Test_Code) Then
            testDesc = invRS!Investigation_Code & " (" & invRS!Investigation_Requested & ") "
         Else
            testDesc = Replace(invRS!EDI_Local_Test_Code, Chr(0), "") & " (" & invRS!Investigation_Requested & ") "
         End If
         
         If invRS!EDI_OP_Suppress Then
'         If vData(2, iPosn) Then
            Status = IS_INV_SUPPRESSED
            msgControl.LogReportMessage IS_INV_SUPPRESSED, testDesc & "Battery Header suppressed - no output for this investigation"
         Else
         
            If IsNull(invRS!EDI_RC_Index) Then
'            If IsNull(vData(10, iPosn)) Then
               rcIndex = -1
            Else
               rcIndex = invRS!EDI_RC_Index 'vData(10, iPosn)
            End If
            
            SegCnt = SegCnt + 1
            ReDim Preserve outStr(0 To SegCnt)
            If IsNull(invRS!EDI_Local_Test_Code) Then
               istr = invRS!Investigation_Code
            Else
               istr = Replace(invRS!EDI_Local_Test_Code, Chr(0), "")
            End If
            
            If sampRS!Collection_Text = "Fasting" Then
               Astr = "F"
            End If
            
'            If blnIncludeRepId Then
'               obr2Str = ID
'            Else
'               obr2Str = ""
'            End If
            
            outStr(SegCnt) = "OBR" & fs _
                                   & Trim(Str(I + 1)) & fs _
                                   & fs _
                                   & sampRS!Sample_Id & fs _
                                   & istr & cs & invRS!Investigation_Requested _
                                   & cs & "LC"
            If Not IsNull(invRS!Read_V2RC) And rcIndex > 0 Then
               outStr(SegCnt) = outStr(SegCnt) & cs & invRS!Read_V2RC & cs & invRS!Read_V2Rubric & cs & "RC"
'               outStr(SegCnt) = outStr(SegCnt) & cs & vData(4, iPosn) & cs & vData(5, iPosn) & cs & "RC"
            End If
            
            outStr(SegCnt) = outStr(SegCnt) & fs _
                                   & fs _
                                   & fs
            If SampleCmts <> "" Then
                If ClinDets <> "" Then
                    ClinDets = ClinDets & "\.br\" & SampleCmts
                Else
                    ClinDets = SampleCmts
                End If
            End If
            
            outStr(SegCnt) = outStr(SegCnt) _
                             & Format(sampRS!Collection_DateTime, "yyyymmddhhmm") & fs _
                             & fs _
                             & fs _
                             & fs _
                             & fs _
                             & fs _
                             & ClinDets & fs
            
            outStr(SegCnt) = outStr(SegCnt) _
                             & Format(sampRS!Collection_DateTimeReceived, "yyyymmddhhmm") & fs _
                             & Trim(sampRS!Sample_Code) & cs & sampRS!Sample_Text & fs _
                             & Recipient & cs & repClin & fs _
                             & fs _
                             & fs _
                             & fs _
                             & fs _
                             & fs
'                             & sampRS!Sample_ID & fs _
                             & fs
            
            outStr(SegCnt) = outStr(SegCnt) _
                             & Format(repListRS!Report_Date, "yyyymmddhhmm") & fs _
                             & fs _
                             & DiscCode & fs
            istr = IIf((Mid(sampRS!Status, 1, 1) = "" Or Mid(sampRS!Status, 1, 1) = "U"), "", Mid(sampRS!Status, 1, 1))
            outStr(SegCnt) = outStr(SegCnt) _
                                   & istr
            If Requestor <> Recipient Then
               outStr(SegCnt) = outStr(SegCnt) & fs & fs & fs & Recipient & cs & repClin
            End If
            
            NTEcnt = 0
            
            With iceCmd
               .CommandText = "ICELABCOMM_Report_InvestComments"
               .Parameters(0) = invRS!Investigation_Index
               Set cRS = .Execute
            End With
            
            Do Until cRS.EOF
            
               If Trim(cRS!Service_Investigation_Comment) <> "" Then
                  SegCnt = SegCnt + 1
                  NTEcnt = NTEcnt + 1
                  ReDim Preserve outStr(0 To SegCnt)
                  outStr(SegCnt) = "NTE" & fs & Trim(Str(NTEcnt)) & fs & "I" & fs & cRS!Service_Investigation_Comment
               End If
               cRS.MoveNext
            Loop
            
            cRS.Close
            
            With iceCmd
               .CommandText = "ICELABCOMM_Report_Results"
               .Parameters(0).Value = invRS!Investigation_Index
               .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
               Set resRS = .Execute
               .Parameters.Delete ("LTSIndex")
            End With
            
            resRS.Filter = "EDI_OP_Suppress = False"
            testCount = resRS.RecordCount
         
'            resRS.Filter = "Direct_EDI_OP_Suppress = False"
'            testCount = testCount + resRS.RecordCount
            
            If testCount = 0 Then
               Status = IS_INV_SUPPRESSED
               msgControl.LogReportMessage Status, invRS!Investigation_Code & " (" & invRS!Investigation_Requested & _
                                           ") ALL the tests in this battery have their output suppressed"
            Else
               resRS.Filter = ""
               
'               If IsNull(resRS!EDI_RC_Index) Then
'                  vFields = Array(16, 17, 19, 20, 21, 28, 29, 30, 31, 32, 33, 34, 5)
'               Else
'                  vFields = Array(22, 23, 24, 25, 26, 35, 36, 37, 38, 39, 40, 41, 5)
'               End If
'
'               rData = resRS.GetRows(-1, adBookmarkFirst, vFields)
               resRS.MoveFirst
'               rPosn = 0
                  
               OBXcnt = 0
               Do Until resRS.EOF
                  eClass.FurtherInfo = "Sample Index = " & sampRS!Sample_Index & _
                                       " Investigation Index = " & invRS!Investigation_Index & _
                                       " Result Index = " & resRS!Result_Index
                  If IsNull(resRS!EDI_Local_Test_Code) Then
                     testDesc = resRS!Result_Code & " (" & resRS!Result_Rubric & ") "
                  Else
                     testDesc = Replace(resRS!EDI_Local_Test_Code, Chr(0), "") & " (" & resRS!EDI_Local_Rubric & ") "
                  End If
                  
                  If resRS!EDI_OP_Suppress Then
                     Status = IS_TEST_SUPPRESSED
                     msgControl.LogReportMessage Status, testDesc & "Test suppressed"
                  Else
               
                     If IsNull(resRS!EDI_RC_Index) Then
                        rcIndex = -1
                     Else
                        rcIndex = resRS!EDI_RC_Index
                     End If
                     
                     istr = "ST"
                     jstr = Replace(resRS!EDI_Local_Test_Code, Chr(0), "")
                     strArray = Split(msgData.AnalyseResultString(resRS!Result), "|")
                     If IsNumeric(strArray(0)) Then
                        istr = "NM"
                     End If
                  
                     msgData.AdjustRangeData Trim(resRS!Lower_Range & ""), Trim(resRS!Upper_Range & "")
                     
                     Astr = ""
                     If resRS!Abnormal_Flag Then
                        If Val(strArray(0)) < Val(msgData.LowRange) Then
                           Astr = "L"
                        ElseIf Val(strArray(0)) > Val(msgData.HighRange) Then
                           Astr = "H"
                        End If
                     End If
                                     
                     ResultValue = ""
                     ResultCmt = ""
                     InitRes = False
                     InitRepLine = False
                      
                     ResultValue = resRS!Result
                     InitRepLine = True
                     NTEcnt = 0
                     OBXcnt = OBXcnt + 1
                     SegCnt = SegCnt + 1
                     ReDim Preserve outStr(0 To SegCnt)
                     outStr(SegCnt) = "OBX" & fs _
                                 & Trim(Str(OBXcnt)) & fs _
                                 & istr & fs _
                                 & jstr & cs & resRS!Result_Rubric & cs & "LC"
                        
                     If resRS!Status = "" Then
                        jstr = "F"
                     Else
                        jstr = resRS!Status
                     End If
                        
                     If istr = "NM" Then
                        If Not IsNull(resRS!Read_V2RC) Or rcIndex > 0 Then
                           outStr(SegCnt) = outStr(SegCnt) & cs & resRS!Read_V2RC & cs & resRS!Read_V2Rubric & cs & "RC"
                        End If
                        
                        If resRS!Lower_Range & "" = "" Then
                           rangeStr = "<" & resRS!Upper_Range
                        ElseIf resRS!Upper_Range & "" = "" Then
                           rangeStr = ">" & resRS!Lower_Range
                        Else
                           rangeStr = resRS!Lower_Range & "-" & resRS!Upper_Range
                        End If
                        
                        outStr(SegCnt) = outStr(SegCnt) & fs _
                                         & fs _
                                         & ResultValue & fs _
                                         & resRS!UOM_Text & fs _
                                         & rangeStr & fs _
                                         & Astr & fs _
                                         & fs _
                                         & fs _
                                         & jstr & fs _
                                         & fs _
                                         & fs
                     Else
                        outStr(SegCnt) = outStr(SegCnt) & fs _
                                         & fs _
                                         & ResultValue & fs _
                                         & fs _
                                         & fs _
                                         & Astr & fs _
                                         & fs _
                                         & fs _
                                         & jstr & fs _
                                         & fs _
                                         & fs
                     End If
                        
                     outStr(SegCnt) = outStr(SegCnt) _
                                      & Format(repListRS!Report_Date, "yyyymmddhhmm") & fs ' _
                                      & ServiceId
                      
                     With iceCmd
                        .CommandText = "ICELABCOMM_Report_ResultComment"
                        .Parameters(0).Value = resRS!Result_Index
                        Set cRS = .Execute
                     End With
                      
                     Do Until cRS.EOF
                        If ResultValue = "" Or cRS!Comment_Type <> "R" Then
                           NTEcnt = NTEcnt + 1
                           SegCnt = SegCnt + 1
                           ReDim Preserve outStr(0 To SegCnt)
                           outStr(SegCnt) = "NTE" & fs & Trim(Str(NTEcnt)) & fs & "L" & fs & cRS!Service_Result_Comment
                        Else
'                           Select Case cRS!Comment_Type
'                              Case "R"
                                 InitRepLine = True
                                 ResultValue = cRS!Service_Result_Comment
                                 NTEcnt = 0
                                 OBXcnt = OBXcnt + 1
                                 SegCnt = SegCnt + 1
                                 ReDim Preserve outStr(0 To SegCnt)
                                 outStr(SegCnt) = "OBX" & fs _
                                                  & Trim(Str(OBXcnt)) & fs _
                                                  & istr & fs _
                                                  & jstr & cs & resRS!Result_Rubric & cs & "LC"
                                 
                                 If Not IsNull(resRS!Read_V2RC) Then
                                    outStr(SegCnt) = outStr(SegCnt) & cs & resRS!Read_V2RC & cs & resRS!Read_V2Rubric & cs & "RC"
                                 End If
                                 
                                 outStr(SegCnt) = outStr(SegCnt) & fs _
                                                  & fs _
                                                  & ResultValue & fs _
                                                  & resRS!UOM_Text & fs _
                                                  & rangeStr & fs _
                                                  & Astr & fs _
                                                  & fs _
                                                  & fs _
                                                  & resRS!Status & fs _
                                                  & fs _
                                                  & fs
                                 
                                 outStr(SegCnt) = outStr(SegCnt) _
                                                  & Format(repListRS!Report_Date, "yyyymmddhhmm") & fs _
                                                  & ServiceId
 '                             Case Else
 '                                NTEcnt = NTEcnt + 1
 '                                SegCnt = SegCnt + 1
 '                                ReDim Preserve outStr(0 To SegCnt)
 '                                outStr(SegCnt) = "NTE" & fs & Trim(Str(NTEcnt)) & fs & "L" & fs & cRS!Service_Result_Comment
 '
 '                          End Select
                        End If
                        cRS.MoveNext
                     Loop
                     cRS.Close
                  End If
                  
                  resRS.MoveNext
                  rPosn = resRS.AbsolutePosition - 1
               Loop
            End If
            resRS.Close
         End If
         
         invRS.MoveNext
         iPosn = invRS.AbsolutePosition - 1
      Loop
      invRS.Close
      sampRS.MoveNext
   Loop
   
   If segCheck = SegCnt Then
      If sampRS.RecordCount = 0 Then
         msgControl.ReportStatus = RS_DATA_INTEGRITY
         msgControl.LogReportMessage RS_DATA_INTEGRITY, "No samples recorded against this report"
      Else
         msgControl.ReportStatus = RS_SUPPRESSION
         msgControl.LogReportMessage RS_SUPPRESSION, "No tests/results to ouput." & _
                                     " Check for any suppression messages"
      End If
   End If
      
   sampRS.Close
   
   Set cRS = Nothing
   Set resRS = Nothing
   Set invRS = Nothing
   Set sampRS = Nothing
   Set iceCmd = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HL7_XML.Process_Report"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub IncXMLCnt()
    XMLCnt = XMLCnt + 1
    ReDim Preserve istr(0 To XMLCnt)
End Sub

Public Sub Complete_Last_Segment(I As Integer)
   IncXMLCnt
   Select Case I
    'Case 1
    '    iStr(XMLCnt) = "</ORU_R01.PATIENT>"
      Case 2
         istr(XMLCnt) = "</ORU_R01.PATIENT_VISIT>"
         IncXMLCnt
         istr(XMLCnt) = "</ORU_R01.PATIENT>"
      Case 3
         istr(XMLCnt) = "</ORU_R01.ORDER_OBSERVATION>"
      Case 4
         istr(XMLCnt) = "</ORU_R01.ORDER_OBSERVATION>"
      Case 5
         istr(XMLCnt) = "</ORU_R01.OBSERVATION>"
    
   End Select
End Sub

Public Sub xComplete_Last_Segment(I As Integer)
IncXMLCnt
Select Case I
    Case 1
        istr(XMLCnt) = "</ORU_R01.PATIENT>"
    Case 2
        istr(XMLCnt) = "</ORU_R01.VISIT>"
    Case 3
         istr(XMLCnt) = "</ORU_R01.ORDER_OBSERVATION>"
    Case 4
         istr(XMLCnt) = "</ORU_R01.ORDER_OBSERVATION>"
    Case 5
         istr(XMLCnt) = "</ORU_R01.OBSERVATION_RESULT>"
    
End Select

End Sub

Public Property Let ProduceXML(blnNewValue As Boolean)
   blnXMLReqd = blnNewValue
End Property

Public Sub Create_HL7XML()
   Dim I As Integer
   Dim arrayStart As Integer
   Dim j As Integer
   Dim n As Integer
   Dim Q As Integer
   Dim NTEType As Integer
   Dim tStr
   Dim uStr
   Dim yStr

   XMLCnt = 0
   NTEType = 0
   ReDim Preserve istr(0 To XMLCnt)
'   IncXMLCnt
'   istr(XMLCnt) = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
'   IncXMLCnt
'   istr(XMLCnt) = "<ORU_R01>"
   
'   For I = Abs(blnNoMSH) To SegCnt
   For I = 1 To SegCnt
      outStr(I) = Replace(outStr(I), "&", "&amp;")
      outStr(I) = Replace(outStr(I), "<", "&lt;")
      outStr(I) = Replace(outStr(I), ">", "&gt;")
      outStr(I) = Replace(outStr(I), Chr(34), "&quot;")
      outStr(I) = Replace(outStr(I), "'", "&apos;")
      tStr = Split(outStr(I), fs, -1)
      
      Select Case tStr(0)
         Case "MSH"
            NTEType = 0
            IncXMLCnt
            istr(XMLCnt) = "<MSH>"
            IncXMLCnt
            istr(XMLCnt) = "<MSH.1>" & fs & "</MSH.1>"
            
            For j = 1 To UBound(tStr)
               n = j + 1
               tStr(j) = Replace(tStr(j), "<", "&lt;")
               If Trim(tStr(j)) <> "" Then
                  IncXMLCnt
                  istr(XMLCnt) = "<MSH." & Trim(Str(n)) & ">"
                  
                  Select Case n
                     Case 2, 10
                        istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                            
                     Case 3, 5, 6
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<HD.1>" & tStr(j) & "</HD.1>" & vbCrLf
                            
                     Case 4
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<HD.1>" & tStr(j) & "</HD.1>" & vbCrLf
                     
                     Case 7
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                     
                     Case 9
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<MSG.1>" & uStr(0) & "</MSG.1>"
                        If UBound(uStr) > 0 Then
                           If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<MSG.2>" & uStr(1) & "</MSG.2>" & vbCrLf
                        End If
                     
                     Case 11
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<PT.1>" & tStr(j) & "</PT.1>" & vbCrLf
                            
                     Case 12
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<VID.1>2.4</VID.1>" & vbCrLf
                        
                  End Select
                  
                  istr(XMLCnt) = istr(XMLCnt) & "</MSH." & Trim(Str(n)) & ">"
               End If
            Next j
            
            IncXMLCnt
            istr(XMLCnt) = "</MSH>"
        
         Case "PID"
            NTEType = 1
            IncXMLCnt
            istr(XMLCnt) = "<ORU_R01.PATIENT_RESULT>"
            IncXMLCnt
            istr(XMLCnt) = "<ORU_R01.PATIENT>" & vbCrLf & "<PID>"
            
            For j = 1 To UBound(tStr)
               n = j
               If Trim(tStr(j)) <> "" Then
                  IncXMLCnt
                  istr(XMLCnt) = "<PID." & Trim(Str(n)) & ">"
                  
                  Select Case n
                     Case 1
                        istr(XMLCnt) = istr(XMLCnt) & "1"
                     
                     Case 2
                     'not used by ICE
                            
                     Case 3
                        yStr = Split(tStr(j), RS, -1)
                        For Q = 0 To UBound(yStr)
                           If Q = 1 Then
                              istr(XMLCnt) = istr(XMLCnt) & "</PID." & Trim(Str(n)) & ">"
                              IncXMLCnt
                              istr(XMLCnt) = "<PID." & Trim(Str(n)) & ">"
                           End If
                           uStr = Split(yStr(Q), cs, -1)
                           If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CX.1>" & uStr(0) & "</CX.1>"
'                                    If UBound(uStr) > 2 Then
'                                        If Trim(uStr(3)) <> "" Then iStr(XMLCnt) = iStr(XMLCnt) & vbCrLf & "<CX.5>" & vbCrLf & "<HD.1>" & uStr(3) & "</HD.1>" & vbCrLf & "</CX.5>" & vbCrLf
'                                    End If
                           If UBound(uStr) > 3 Then
                              If Trim(uStr(4)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CX.5>" & uStr(4) & "</CX.5>" & vbCrLf
                           End If
                        Next Q
                            
                     Case 5
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XPN.1>" & vbCrLf & "<FN.1>" & uStr(0) & "</FN.1>" & vbCrLf & "</XPN.1>"
                        If UBound(uStr) > 0 Then
                           If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XPN.2>" & uStr(1) & vbCrLf & "</XPN.2>" & vbCrLf
                        End If
                        
                        If UBound(uStr) > 5 Then
                           If Trim(uStr(6)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XPN.7>" & uStr(6) & vbCrLf & "</XPN.7>" & vbCrLf
                        End If
                     
                     Case 7
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                     
                     Case 8, 12, 19
                        istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                            
                     Case 11
                        uStr = Split(tStr(j), cs, -1)
                        For Q = 0 To UBound(uStr)
                           If Trim(uStr(Q)) <> "" Then
                              If Q = 0 Then
                                 istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XAD." & Trim(Str(Q + 1)) & ">" & vbCrLf & "<SAD.1>" & uStr(0) & vbCrLf & "</SAD.1>" & vbCrLf & "</XAD." & Trim(Str(Q + 1)) & ">"
                              Else
                                 istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XAD." & Trim(Str(Q + 1)) & ">" & uStr(Q) & "</XAD." & Trim(Str(Q + 1)) & ">"
                              End If
                           End If
                        Next Q
                        
                  End Select
                  
                  istr(XMLCnt) = istr(XMLCnt) & "</PID." & Trim(Str(n)) & ">"
               End If
            Next j
            IncXMLCnt
            istr(XMLCnt) = "</PID>"
        
         Case "PV1"
'                If NTEType = 1 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</PID>"
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORU_R01.PATIENT>"
'                End If
                'Call Complete_Last_Segment(NTEType)
            NTEType = 2
            IncXMLCnt
            istr(XMLCnt) = "<ORU_R01.PATIENT_VISIT>" & vbCrLf & "<PV1>"
            
            For j = 1 To UBound(tStr)
               n = j
               If Trim(tStr(j)) <> "" Then
                  IncXMLCnt
                  istr(XMLCnt) = "<PV1." & Trim(Str(n)) & ">"
                  
                  Select Case n
                     Case 2
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & tStr(j) & vbCrLf
                  
                     Case 3
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<PL.4><HD.1>" & tStr(j) & "</HD.1></PL.4>" & vbCrLf
                     
                     Case 7, 8
                        If blnClinInPV1 Then
                           uStr = Split(tStr(j), cs, -1)
                           
                           If Trim(uStr(0)) <> "" Then
                              istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & "</XCN.1>" & vbCrLf
                           End If
                           
                           If UBound(uStr) > 0 Then
                              If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                           End If
                        End If
                        
                     Case 19
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CX.1>" & tStr(j) & "</CX.1>" & vbCrLf
                        
'                     Case 51
'                        istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                        
                  End Select
                  
                  istr(XMLCnt) = istr(XMLCnt) & "</PV1." & Trim(Str(n)) & ">"
               End If
            Next j
            
            IncXMLCnt
            istr(XMLCnt) = "</PV1>"
        
         Case "ORC"
'                If NTEType = 2 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</PV1>"
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORU_R01.VISIT>"
'                End If
            Call Complete_Last_Segment(NTEType)
            NTEType = 3
            IncXMLCnt
            istr(XMLCnt) = "<ORU_R01.ORDER_OBSERVATION>" & vbCrLf & "<ORC>"
            For j = 1 To UBound(tStr)
               n = j
               If Trim(tStr(j)) <> "" Then
                  IncXMLCnt
                  istr(XMLCnt) = "<ORC." & Trim(Str(n)) & ">"
                  
                  Select Case n
                     Case 1
                        istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                     
                     Case 3
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<EI.1>" & tStr(j) & "</EI.1>" & vbCrLf
                     
                     Case 12
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & vbCrLf & "</XCN.1>" & vbCrLf
                        If UBound(uStr) > 0 Then
                           If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                        End If
                     
                     Case 13
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<PL.1>" & tStr(j) & "</PL.1>" & vbCrLf
                  
                  End Select
                  
                  istr(XMLCnt) = istr(XMLCnt) & "</ORC." & Trim(Str(n)) & ">"
               End If
            Next j
            
            IncXMLCnt
            istr(XMLCnt) = "</ORC>"
        
         Case "OBR"
'                If NTEType = 3 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORC>"
'                End If
            If NTEType > 3 Or NTEType = 2 Then
               Call Complete_Last_Segment(NTEType)
               IncXMLCnt
'              Added by me
               If NTEType > 3 Then
                  Call Complete_Last_Segment(4)
                  IncXMLCnt
               End If
               istr(XMLCnt) = "<ORU_R01.ORDER_OBSERVATION>"
            End If
            NTEType = 4
            IncXMLCnt
            istr(XMLCnt) = "<OBR>"
            
            For j = 1 To UBound(tStr)
               n = j
               If Trim(tStr(j)) <> "" Then
                  IncXMLCnt
                  istr(XMLCnt) = "<OBR." & Trim(Str(n)) & ">"
                  
                  Select Case n
                     Case 1, 2, 13, 20, 24, 25
                        istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                        
                     Case 3
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<EI.1>" & tStr(j) & "</EI.1>" & vbCrLf
                        istr(XMLCnt) = istr(XMLCnt) & "<EI.2>" & GetSampleLab(CStr(tStr(j))) & "</EI.2>" & vbCrLf
                     
                     Case 4
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf
                        If UBound(uStr) > 0 Then
                           If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.2>" & uStr(1) & "</CE.2>" & vbCrLf
                        End If
                        
                        If UBound(uStr) > 1 Then
                           If Trim(uStr(2)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.3>" & uStr(2) & "</CE.3>" & vbCrLf
                        End If
                                
                        If UBound(uStr) > 2 Then
                           If Trim(uStr(3)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.4>" & uStr(3) & "</CE.4>" & vbCrLf
                        End If
                        
                        If UBound(uStr) > 3 Then
                           If Trim(uStr(4)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.5>" & uStr(4) & "</CE.5>" & vbCrLf
                        End If
                                
                        If UBound(uStr) > 4 Then
                           If Trim(uStr(5)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.6>" & uStr(5) & "</CE.6>" & vbCrLf
                        End If
                     
                     Case 7, 14, 22
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                     
                     Case 15
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<SPS.1>" & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf & "</SPS.1>" & vbCrLf
                        If UBound(uStr) > 0 Then
                              If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<SPS.2>" & uStr(1) & "</SPS.2>" & vbCrLf
                        End If
                     
                     Case 16
                        uStr = Split(tStr(j), cs, -1)
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & vbCrLf & "</XCN.1>" & vbCrLf
                        If UBound(uStr) > 0 Then
                           If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                        End If
                     
                  End Select
                        
                  istr(XMLCnt) = istr(XMLCnt) & "</OBR." & Trim(Str(n)) & ">"
               End If
            Next j
            
            IncXMLCnt
            istr(XMLCnt) = "</OBR>"
        
         Case "OBX"
'                If NTEType = 4 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORU_R01.ORDER_OBSERVATION>"
'                End If
            If NTEType > 4 Then Call Complete_Last_Segment(NTEType)
            NTEType = 5
            IncXMLCnt
            istr(XMLCnt) = "<ORU_R01.OBSERVATION>" & vbCrLf & "<OBX>"
            For j = 1 To UBound(tStr)
               n = j
               If Trim(tStr(j)) <> "" Then
                  IncXMLCnt
                  istr(XMLCnt) = "<OBX." & Trim(Str(n)) & ">"
                  
                  Select Case n
                     Case 1, 2, 5, 7, 8, 11
                        istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                  
                     Case 3
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf
                        If UBound(uStr) > 0 Then
                           If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.2>" & uStr(1) & "</CE.2>" & vbCrLf
                        End If
                        
                        If UBound(uStr) > 1 Then
                           If Trim(uStr(2)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.3>" & uStr(2) & "</CE.3>" & vbCrLf
                        End If
                        
                        If UBound(uStr) > 2 Then
                           If Trim(uStr(3)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.4>" & uStr(3) & "</CE.4>" & vbCrLf
                        End If
                             
                        If UBound(uStr) > 3 Then
                           If Trim(uStr(4)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.5>" & uStr(4) & "</CE.5>" & vbCrLf
                        End If
                             
                        If UBound(uStr) > 4 Then
                           If Trim(uStr(5)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.6>" & uStr(5) & "</CE.6>" & vbCrLf
                        End If
                         
                     Case 6, 15
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf
                         
                     Case 11
                     
                     
                     Case 12
                        uStr = Split(tStr(j), cs, -1)
                        If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & vbCrLf & "</XCN.1>" & vbCrLf
                        If UBound(uStr) > 0 Then
                           If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                        End If
                         
                     Case 14
                        istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                     
                  End Select
               
                  istr(XMLCnt) = istr(XMLCnt) & "</OBX." & Trim(Str(n)) & ">"
               End If
            Next j
            
            IncXMLCnt
            istr(XMLCnt) = "</OBX>"
        
         Case "NTE"
            IncXMLCnt
            istr(XMLCnt) = "<NTE>"
            For j = 1 To UBound(tStr)
               n = j
               If Trim(tStr(j)) <> "" Then
                  IncXMLCnt
                  istr(XMLCnt) = "<NTE." & Trim(Str(n)) & ">"
                  istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                  istr(XMLCnt) = istr(XMLCnt) & "</NTE." & Trim(Str(n)) & ">"
               End If
            Next j
            
            IncXMLCnt
            istr(XMLCnt) = "</NTE>"
      
      End Select
   Next I
   
   Call Complete_Last_Segment(NTEType)

'  The 'If' clause Added by me
   If NTEType <> 4 Then
      Call Complete_Last_Segment(4)
   End If
   IncXMLCnt
   istr(XMLCnt) = "</ORU_R01.PATIENT_RESULT>"
'   IncXMLCnt
'   istr(XMLCnt) = "</ORU_R01>"
End Sub

Public Sub xCreate_HL7XML()
   On Error GoTo procEH
   Dim I As Integer
   Dim j As Integer
   Dim n As Integer
   Dim Q As Integer
   Dim NTEType As Integer
   Dim tStr
   Dim uStr
   Dim yStr

XMLCnt = 0
NTEType = 0
ReDim Preserve istr(0 To XMLCnt)
IncXMLCnt
istr(XMLCnt) = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
IncXMLCnt
istr(XMLCnt) = "<ORU_R01>"
For I = 0 To SegCnt
    outStr(I) = Replace(outStr(I), "&", "&amp;")
    outStr(I) = Replace(outStr(I), "<", "&lt;")
    outStr(I) = Replace(outStr(I), ">", "&gt;")
    outStr(I) = Replace(outStr(I), Chr(34), "&quot;")
    outStr(I) = Replace(outStr(I), "'", "&apos;")
    tStr = Split(outStr(I), fs, -1)
    Select Case tStr(0)
        Case "MSH"
                NTEType = 0
                IncXMLCnt
                istr(XMLCnt) = "<MSH>"
                IncXMLCnt
                istr(XMLCnt) = "<MSH.1>" & fs & "</MSH.1>"
                For j = 1 To UBound(tStr)
                    n = j + 1
                    tStr(j) = Replace(tStr(j), "<", "&lt;")
                    If Trim(tStr(j)) <> "" Then
                        IncXMLCnt
                        istr(XMLCnt) = "<MSH." & Trim(Str(n)) & ">"
                        Select Case n
                            Case 2, 10
                                istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                            Case 3, 4, 5, 6
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<HD.1>" & tStr(j) & "</HD.1>" & vbCrLf
                            Case 7
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                            Case 9
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CM_MSG.1>" & uStr(0) & "</CM_MSG.1>"
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CM_MSG.2>" & uStr(1) & "</CM_MSG.2>" & vbCrLf
                                End If
                            Case 11
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<PT.1>" & tStr(j) & "</PT.1>" & vbCrLf
                            Case 12
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<VID.1>2.4</VID.1>" & vbCrLf
                        End Select
                        istr(XMLCnt) = istr(XMLCnt) & "</MSH." & Trim(Str(n)) & ">"
                    End If
                Next j
                IncXMLCnt
                istr(XMLCnt) = "</MSH>"
        Case "PID"
                NTEType = 1
                IncXMLCnt
                istr(XMLCnt) = "<ORU_R01.PATIENT>" & vbCrLf & "<PID>"
                For j = 1 To UBound(tStr)
                    n = j
                    If Trim(tStr(j)) <> "" Then
                        IncXMLCnt
                        istr(XMLCnt) = "<PID." & Trim(Str(n)) & ">"
                        Select Case n
                            Case 1, 2
                                'not used by ICE
                            Case 3
                                yStr = Split(tStr(j), RS, -1)
                                For Q = 0 To UBound(yStr)
                                    If Q = 1 Then
                                        istr(XMLCnt) = istr(XMLCnt) & "</PID." & Trim(Str(n)) & ">"
                                        IncXMLCnt
                                        istr(XMLCnt) = "<PID." & Trim(Str(n)) & ">"
                                    End If
                                    uStr = Split(yStr(Q), cs, -1)
                                    If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CX.1>" & uStr(0) & "</CX.1>"
'                                    If UBound(uStr) > 2 Then
'                                        If Trim(uStr(3)) <> "" Then iStr(XMLCnt) = iStr(XMLCnt) & vbCrLf & "<CX.5>" & vbCrLf & "<HD.1>" & uStr(3) & "</HD.1>" & vbCrLf & "</CX.5>" & vbCrLf
'                                    End If
                                    If UBound(uStr) > 3 Then
                                        If Trim(uStr(4)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CX.5>" & uStr(4) & "</CX.5>" & vbCrLf
                                    End If
                                Next Q
                            Case 5
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XPN.1>" & vbCrLf & "<FN.1>" & uStr(0) & "</FN.1>" & vbCrLf & "</XPN.1>"
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XPN.2>" & uStr(1) & vbCrLf & "</XPN.2>" & vbCrLf
                                End If
                                If UBound(uStr) > 5 Then
                                    If Trim(uStr(6)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XPN.7>" & uStr(6) & vbCrLf & "</XPN.7>" & vbCrLf
                                End If
                            Case 7
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                            Case 8, 12, 19
                                istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                            Case 11
                                uStr = Split(tStr(j), cs, -1)
                                For Q = 0 To UBound(uStr)
                                    If Trim(uStr(Q)) <> "" Then
                                        If Q = 0 Then
                                            istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XAD." & Trim(Str(Q + 1)) & ">" & vbCrLf & "<SAD.1>" & uStr(0) & vbCrLf & "</SAD.1>" & vbCrLf & "</XAD." & Trim(Str(Q + 1)) & ">"
                                        Else
                                            istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XAD." & Trim(Str(Q + 1)) & ">" & uStr(Q) & "</XAD." & Trim(Str(Q + 1)) & ">"
                                        End If
                                    End If
                                Next Q
                        End Select
                        istr(XMLCnt) = istr(XMLCnt) & "</PID." & Trim(Str(n)) & ">"
                    End If
                Next j
                IncXMLCnt
                istr(XMLCnt) = "</PID>"
        Case "PV1"
'                If NTEType = 1 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</PID>"
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORU_R01.PATIENT>"
'                End If
                Call Complete_Last_Segment(NTEType)
                NTEType = 2
                IncXMLCnt
                istr(XMLCnt) = "<ORU_R01.VISIT>" & vbCrLf & "<PV1>"
                For j = 1 To UBound(tStr)
                    n = j
                    If Trim(tStr(j)) <> "" Then
                        IncXMLCnt
                        istr(XMLCnt) = "<PV1." & Trim(Str(n)) & ">"
                        Select Case n
                            Case 3
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<PL.1>" & tStr(j) & "</PL.1>" & vbCrLf
                            Case 7, 8
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & "</XCN.1>" & vbCrLf
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                                End If
                        End Select
                        istr(XMLCnt) = istr(XMLCnt) & "</PV1." & Trim(Str(n)) & ">"
                    End If
                Next j
                IncXMLCnt
                istr(XMLCnt) = "</PV1>"
        Case "ORC"
'                If NTEType = 2 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</PV1>"
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORU_R01.VISIT>"
'                End If
                Call Complete_Last_Segment(NTEType)
                NTEType = 3
                IncXMLCnt
                istr(XMLCnt) = "<ORU_R01.ORDER_OBSERVATION>" & vbCrLf & "<ORC>"
                For j = 1 To UBound(tStr)
                    n = j
                    If Trim(tStr(j)) <> "" Then
                        IncXMLCnt
                        istr(XMLCnt) = "<ORC." & Trim(Str(n)) & ">"
                        Select Case n
                            Case 1
                                istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                            Case 3
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<EI.1>" & tStr(j) & "</EI.1>" & vbCrLf
                            Case 12
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & vbCrLf & "</XCN.1>" & vbCrLf
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                                End If
                            Case 13
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<PL.1>" & tStr(j) & "</PL.1>" & vbCrLf
                        End Select
                        istr(XMLCnt) = istr(XMLCnt) & "</ORC." & Trim(Str(n)) & ">"
                    End If
                Next j
                IncXMLCnt
                istr(XMLCnt) = "</ORC>"
        Case "OBR"
'                If NTEType = 3 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORC>"
'                End If
                If NTEType > 3 Then
                    Call Complete_Last_Segment(NTEType)
                    IncXMLCnt
                    istr(XMLCnt) = "<ORU_R01.ORDER_OBSERVATION>"
                End If
                NTEType = 4
                IncXMLCnt
                istr(XMLCnt) = "<OBR>"
                For j = 1 To UBound(tStr)
                    n = j
                    If Trim(tStr(j)) <> "" Then
                        IncXMLCnt
                        istr(XMLCnt) = "<OBR." & Trim(Str(n)) & ">"
                        Select Case n
                            Case 1, 13, 20, 24, 25
                                istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                            Case 3
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<EI.1>" & tStr(j) & "</EI.1>" & vbCrLf
                            Case 4
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.2>" & uStr(1) & "</CE.2>" & vbCrLf
                                End If
                                If UBound(uStr) > 1 Then
                                    If Trim(uStr(2)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.3>" & uStr(2) & "</CE.3>" & vbCrLf
                                End If
                                If UBound(uStr) > 2 Then
                                    If Trim(uStr(3)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.4>" & uStr(3) & "</CE.4>" & vbCrLf
                                End If
                                If UBound(uStr) > 3 Then
                                    If Trim(uStr(4)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.5>" & uStr(4) & "</CE.5>" & vbCrLf
                                End If
                                If UBound(uStr) > 4 Then
                                    If Trim(uStr(5)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.6>" & uStr(5) & "</CE.6>" & vbCrLf
                                End If
                            Case 7, 14, 22
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                            Case 15
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CM_SPS.1>" & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf & "</CM_SPS.1>" & vbCrLf
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CM_SPS.2>" & uStr(1) & "</CM_SPS.2>" & vbCrLf
                                End If
                            Case 16
                                uStr = Split(tStr(j), cs, -1)
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & vbCrLf & "</XCN.1>" & vbCrLf
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                                End If
                        End Select
                        istr(XMLCnt) = istr(XMLCnt) & "</OBR." & Trim(Str(n)) & ">"
                    End If
                Next j
                IncXMLCnt
                istr(XMLCnt) = "</OBR>"
        Case "OBX"
'                If NTEType = 4 Then
'                    IncXMLCnt
'                    iStr(XMLCnt) = "</ORU_R01.ORDER_OBSERVATION>"
'                End If
                If NTEType > 3 Then Call Complete_Last_Segment(NTEType)
                NTEType = 5
                IncXMLCnt
                istr(XMLCnt) = "<ORU_R01.OBSERVATION_RESULT>" & vbCrLf & "<OBX>"
                For j = 1 To UBound(tStr)
                    n = j
                    If Trim(tStr(j)) <> "" Then
                        IncXMLCnt
                        istr(XMLCnt) = "<OBX." & Trim(Str(n)) & ">"
                        Select Case n
                            Case 1, 2, 5, 7, 8
                                istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                            Case 3
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.2>" & uStr(1) & "</CE.2>" & vbCrLf
                                End If
                                If UBound(uStr) > 1 Then
                                    If Trim(uStr(2)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.3>" & uStr(2) & "</CE.3>" & vbCrLf
                                End If
                                If UBound(uStr) > 2 Then
                                    If Trim(uStr(3)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.4>" & uStr(3) & "</CE.4>" & vbCrLf
                                End If
                                If UBound(uStr) > 3 Then
                                    If Trim(uStr(4)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.5>" & uStr(4) & "</CE.5>" & vbCrLf
                                End If
                                If UBound(uStr) > 4 Then
                                    If Trim(uStr(5)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<CE.6>" & uStr(5) & "</CE.6>" & vbCrLf
                                End If
                            Case 6, 15
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<CE.1>" & uStr(0) & "</CE.1>" & vbCrLf
                            Case 12
                                uStr = Split(tStr(j), cs, -1)
                                If Trim(uStr(0)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<XCN.1>" & uStr(0) & vbCrLf & "</XCN.1>" & vbCrLf
                                If UBound(uStr) > 0 Then
                                    If Trim(uStr(1)) <> "" Then istr(XMLCnt) = istr(XMLCnt) & "<XCN.2>" & vbCrLf & "<FN.1>" & uStr(1) & "</FN.1>" & vbCrLf & "</XCN.2>" & vbCrLf
                                End If
                            Case 14
                                istr(XMLCnt) = istr(XMLCnt) & vbCrLf & "<TS.1>" & tStr(j) & "</TS.1>" & vbCrLf
                        End Select
                        istr(XMLCnt) = istr(XMLCnt) & "</OBX." & Trim(Str(n)) & ">"
                    End If
                Next j
                IncXMLCnt
                istr(XMLCnt) = "</OBX>"
        Case "NTE"
                IncXMLCnt
                istr(XMLCnt) = "<NTE>"
                For j = 1 To UBound(tStr)
                    n = j
                    If Trim(tStr(j)) <> "" Then
                        IncXMLCnt
                        istr(XMLCnt) = "<NTE." & Trim(Str(n)) & ">"
                        istr(XMLCnt) = istr(XMLCnt) & tStr(j)
                        istr(XMLCnt) = istr(XMLCnt) & "</NTE." & Trim(Str(n)) & ">"
                    End If
                Next j
                IncXMLCnt
                istr(XMLCnt) = "</NTE>"
    
    
    End Select
Next I
Call Complete_Last_Segment(NTEType)
IncXMLCnt
istr(XMLCnt) = "</ORU_R01>"
Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HL7_XML.Create_HL7XML"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
'
'Public Function Data_To_Process() As Boolean
'Dim strSQL As String
'  Data_To_Process = False
'  strSQL = "select * from HL7_REP_LIST"
'  Set IceRst = New ADODB.Recordset
'  IceRst.Open strSQL, ICEcon, adOpenKeyset, adLockPessimistic
'  If (Not IceRst.BOF) And (Not IceRst.EOF) Then
'    Data_To_Process = True
'  End If
'End Function

Public Function CreateMessage(RepId As Long) As String
   On Error GoTo procEH
   Dim sbuf As New StringBuffer
   Dim I As Integer

'   RepId = repListRS!EDI_Report_Index
   ReDim outStr(0)
   SegCnt = 0
   ProcessReport (RepId)
   
   DoEvents
      
   If blnXMLReqd Then
      Create_HL7XML
      For I = 0 To XMLCnt
         sbuf.Append istr(I)
      Next I
'      Open "c:\ice\labcomm\xml1.xml" For Output As #3
'      Print #3, sbuf.Value
'      Close #3
   Else
'     Call to ensure HL7 file are created 1 per batch
      frmMain.NewFileRequired
      
      sbuf.Clear
      For I = 0 To SegCnt
         sbuf.Append outStr(I) & st
      Next I
'      Open "c:\ice\labcomm\hl71.hl7" For Output As #3
'      Print #3, sbuf.Value
'      Close #3
   End If

   CreateMessage = sbuf.Value
   
'   If SegCnt > 0 Then
'      Output_HL7_Array
'      Create_HL7XML
'      Output_XML_Array
'   End If
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HL7_XML.CreateMessage"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function
