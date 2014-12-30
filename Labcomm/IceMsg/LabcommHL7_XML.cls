VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabcommHL7_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xmlVersion As String
Private som As String
Private eom As String
Private cs As String
Private rc As String
Private ec As String
Private ss As String
Private fs As String
Private st As String
Private procID As String
Private Vers As String
Private Country As String
Private hl7Sender As String
Private xmlSendFacility As String
Private blnUseOrgName As Boolean

Private SampId As String

Private OBRNo As Integer
Private CDasNTE As Boolean
Private RITasNTE As Boolean
Private SCasNTE As Boolean
Private ORCOUT As Boolean

Private xmlDoc As New MSXML2.DOMDocument
Private xsltDoc As New MSXML2.DOMDocument
Private root As MSXML2.IXMLDOMElement
Private mainEl As MSXML2.IXMLDOMElement
Private lev1El As MSXML2.IXMLDOMElement
Private lev2El As MSXML2.IXMLDOMElement
Private lev3El As MSXML2.IXMLDOMElement
Private lev4El As MSXML2.IXMLDOMElement
Private lev5El As MSXML2.IXMLDOMElement
Private lev6El As MSXML2.IXMLDOMElement
Private nTxt As MSXML2.IXMLDOMText

Private Status As Long
Private blnHL7 As Boolean
Private Recipient As String
Private Requestor As String
Private repClin As String
Private reqClin As String
Private hl7Src As String
Private strArray() As String
Private SampleCmts As String
Private ClinDets As String
Private srcCode As String
Private srcExpand As String
Private srcDisc As String
Private srcRadSName As String
Private srcRadFName As String

Private iceCmd As New ADODB.Command

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim vRS As New ADODB.Recordset
   Dim strSQL As String
   Dim delim As String

   If Read_Ini_Var("Data", "ClinDets", INIFile) = "OBR" Then CDasNTE = False
   If Read_Ini_Var("Data", "TextResults", INIFile) = "OBX" Then RITasNTE = False
   If Read_Ini_Var("Data", "SampleComments", INIFile) = "OBR" Then SCasNTE = False
   If Read_Ini_Var("Data", "OutputORC", INIFile) = "1" Then ORCOUT = True
   
   '  Set defaults
   
   som = Chr(11)
   eom = Chr(28)
   ec = Chr(92)
   fs = Chr(124)
   cs = Chr(94)
   ss = Chr(38)
   rc = Chr(126)
   blnUseOrgName = False
   
   '  Overwrite defaults if required
   
   strSQL = "SELECT * FROM EDI_HL7_Params"
   vRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   Do Until vRS.EOF
      Select Case UCase(vRS!ParameterId)
         Case "START"
            som = Chr(CLng(vRS!ParameterValue))
            
         Case "END"
            eom = Chr(CLng(vRS!ParameterValue))
         
         Case "ESCAPE"
            ec = Chr(CLng(vRS!ParameterValue))
         
         Case "FIELD"
            fs = Chr(CLng(vRS!ParameterValue))
         
         Case "COMPONENT"
            cs = Chr(CLng(vRS!ParameterValue))
         
         Case "SUBCOMP"
            ss = Chr(CLng(vRS!ParameterValue))
         
         Case "REPETITION"
            rc = Chr(CLng(vRS!ParameterValue))
            
         Case "PROCID"
            procID = vRS!ParameterValue
            
         Case "VERSION"
            Vers = vRS!ParameterValue
            
         Case "COUNTRY"
            Country = vRS!ParameterValue
         
         Case "XSLT"
            xsltDoc.Load App.Path & "/" & vRS!ParameterValue
            If xsltDoc.childNodes.Length = 0 Then
               Err.Raise 100000, "LabcommHL7_XML", "Transformation File (" & vRS!ParameterValue & ") failed to load"
            End If
            
         Case "USEORGNAME"
            blnUseOrgName = (UCase(vRS!ParameterValue) = "TRUE")
            
      End Select
      
      vRS.MoveNext
   Loop
   
   vRS.Close
   
   strSQL = "SELECT name " & _
            "FROM master.dbo.sysdatabases " & _
            "WHERE name = 'ICEHL7History'"
   vRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   blnHL7 = Not vRS.EOF
   vRS.Close
   
   Set vRS = Nothing
   xmlVersion = "<ORU_R01 xmlns=" & Chr(34) & "urn:hl7-org:v2xml" & Chr(34) & ">"
   
   st = vbCr & vbLf
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.CreateMessage"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function CreateMessage(RepId As Long) As String
   On Error GoTo procEH
   Dim sbuf As String
   Dim escBR As String
   Dim srcRad As String
   Dim I As Integer
   Dim RS As New ADODB.Recordset
   
   Set xmlDoc = Nothing
   Set xmlDoc = New MSXML2.DOMDocument
   
'   If blnHL7 Then
'      strSQL = "SELECT Source_App, Source_Expansion, Source_Discipline, " & _
'                  "substring(message,patindex('%" & vbCrLf & "PV1|%',message)+2,40) " & _
'               "FROM ICEHL7History.dbo.HL7_History hl7 " & _
'                  "INNER JOIN ICEHL7History.dbo.HL7_SourceMap hSrc " & _
'                  "ON hl7.Source_App = hSrc.Source_Code " & _
'                  "INNER JOIN Service_ImpExp_Headers sh " & _
'                     "INNER JOIN Service_ImpExp_Messages sm " & _
'                     "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
'                  "ON ImpExp_File = hl7.HL7_Index " & _
'               "WHERE IsNumeric(ImpExp_File) = 1 " & _
'                  "AND Service_Report_Index = " & RepId
'      RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
'
'      If RS.EOF Then
'         srcCode = ""
'      Else
'         srcCode = RS(0)
'         srcExpand = "X-RAY " & RS(1)
'         srcDisc = RS(2)
'
'         srcRad = Split(RS(3), "|")(4)
'
'         srcRadSName = Split(srcRad, "^")(0)
'         srcRadFName = Split(srcRad, "^")(1)
'
'      End If
'      RS.Close
'   End If
   
   Set RS = Nothing
   
   'xmlDoc.appendChild MSH_Element
   ReadReport RepId
   
   DoEvents
   
   CreateMessage = xmlDoc.transformNode(xsltDoc)
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.CreateMessage"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function FileHeader() As String
   On Error GoTo procEH
   Dim xsltM As New MSXML2.DOMDocument
   Dim MSHEl As MSXML2.IXMLDOMElement
   
   Set xmlDoc = Nothing
   Set xmlDoc = New MSXML2.DOMDocument
   
   xsltM.Load App.Path & "\HL7_MSH.xslt"
   If xsltM.childNodes.Length = 0 Then
      Err.Raise 100000, "LabcommHL7_XML", "Transformation file (HL7_MSH.xslt) failed to load"
   End If
     
   Set MSHEl = MSH_Element
   FileHeader = som & MSHEl.transformNode(xsltM)
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.FileHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function FileTrailer() As String
   'FileTrailer = "</ICE_HL7>"
   FileTrailer = eom
End Function

Private Function Get_DiscCode(txtDiscipline As String) As String
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

Private Function GetSampleLab(SampleText As String, _
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

Private Function PID_Element(RepId As Long) As MSXML2.IXMLDOMElement
   On Error GoTo procEH
   Dim PID As MSXML2.IXMLDOMElement
   Dim PID_1 As MSXML2.IXMLDOMElement
   Dim PID_2 As MSXML2.IXMLDOMElement
   Dim PID_3 As MSXML2.IXMLDOMElement
   
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim idFlag As Integer
   Dim strPID As String
   Dim strIdType As String
   Dim strTemp As String
   Dim CaseNote As String
   Dim P_Surname As String
   Dim P_Midname As String
   Dim P_Forename As String
   Dim P_Title As String
   Dim P_Pat_Addr_Line1 As String
   Dim P_Pat_Addr_Line2 As String
   Dim P_Pat_Addr_Line3 As String
   Dim P_Pat_Addr_Line4 As String
   Dim P_Pat_Addr_Line5 As String
   Dim P_Pat_PostCode As String
   Dim P_Pat_Id As String
   
   Dim blnNoPatId As Boolean
   
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Patient"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RepId)
      Set RS = .Execute
   End With
   
   strSQL = "SELECT HospitalNumber " & _
            "FROM Patient_Local_Ids " & _
            "WHERE Patient_Id_Key in ( " & _
               "SELECT Patient_Id_Key FROM Patient_Local_Ids WHERE Id = " & RS!Patient_Id & ") " & _
               "AND HospitalNumber like 'P%'"
   Set RS2 = iceCon.Execute(strSQL)
   
   If RS2.EOF Then
      CaseNote = ""
   Else
      CaseNote = RS2(0)
   End If
   
   RS2.Close
   
   If msgControl.AnonymizeData Then
      P_Pat_Id = "999999999"
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
      P_Surname = RS!Surname & ""
      P_Midname = RS!Midname & ""
      P_Forename = RS!Forename & ""
      P_Title = RS!Title & ""
      P_Pat_Addr_Line1 = RS!Pat_Addr_Line1 & ""
      P_Pat_Addr_Line2 = RS!Pat_Addr_Line2 & ""
      P_Pat_Addr_Line3 = RS!Pat_Addr_Line3 & ""
      P_Pat_Addr_Line4 = RS!Pat_Addr_Line4 & ""
      P_Pat_Addr_Line5 = RS!Pat_Addr_Line5 & ""
      P_Pat_PostCode = RS!Pat_PostCode & ""
   End If
   
   With xmlDoc
      Set PID = .createElement("PID")
      Set PID_1 = .createElement("PID.1")
      Set nTxt = .createTextNode("1")
      PID_1.appendChild nTxt
      PID.appendChild PID_1
      
      Set PID_1 = .createElement("PID.3")
      Set PID_2 = .createElement("PD.1")
      
      If RS!Hospital_Number & "" <> "" Then
         P_Pat_Id = RS!Hospital_Number
         idFlag = 1
      End If
      
      If RS!New_NHS_No & "" <> "" Then
         idFlag = (idFlag Or 2)
      End If
      
      If idFlag = 0 Then
         P_Pat_Id = RS!Service_Id
         idFlag = 1
      End If
      
      If (idFlag And 1) = 1 Then
         Set PID_3 = .createElement("CX.1")
         Set nTxt = .createTextNode(P_Pat_Id)
         PID_3.appendChild nTxt
         PID_2.appendChild PID_3
               
         Set PID_3 = .createElement("CX.4")
         Set nTxt = .createTextNode(repListRS!EDI_Provider_Org)
         PID_3.appendChild nTxt
         PID_2.appendChild PID_3
               
         Set PID_3 = .createElement("CX.5")
         Set nTxt = .createTextNode("MRN")
         PID_3.appendChild nTxt
         PID_2.appendChild PID_3
         
         PID_1.appendChild PID_2
      End If
      
      If (idFlag And 2) = 2 Then
         Set PID_2 = .createElement("PD.2")
         
         Set PID_3 = .createElement("CX.1")
         Set nTxt = .createTextNode(RS!New_NHS_No)
         PID_3.appendChild nTxt
         PID_2.appendChild PID_3
               
         Set PID_3 = .createElement("CX.4")
         Set nTxt = .createTextNode(repListRS!EDI_Provider_Org)
         PID_3.appendChild nTxt
         PID_2.appendChild PID_3
               
         Set PID_3 = .createElement("CX.5")
         Set nTxt = .createTextNode("NHS")
         PID_3.appendChild nTxt
         PID_2.appendChild PID_3
         PID_1.appendChild PID_2
      End If
         
      Set PID_2 = .createElement("PD.3")
      
      Set PID_3 = .createElement("CX.1")
      Set nTxt = .createTextNode(CaseNote)
      PID_3.appendChild nTxt
      PID_2.appendChild PID_3
      PID_1.appendChild PID_2
         
      PID.appendChild PID_1
'      If blnNoPatId Then
'         If XMLDefPrefix = "" Then
'            strIdType = "LAB"
'         Else
'            strIdType = XMLDefPrefix
'         End If
'      Else
'         If XMLMRNPrefix = "" Then
'            If XMLDefPrefix = "" Then
'               If UCase(Left(P_Pat_Id, 3)) = "APX" Then
'                  strIdType = "LAB"
'               Else
'                  strIdType = "MRN"
'               End If
'            Else
'               strIdType = XMLDefPrefix
'            End If
'         Else
'            If InStr(1, XMLMRNPrefix, Left(P_Pat_Id, 1), vbTextCompare) > 0 Then
'               strIdType = "MRN"
'            ElseIf XMLDefPrefix = "" Then
'               strIdType = "LAB"
'            Else
'               strIdType = XMLDefPrefix
''            ElseIf InStr(1, XMLLABPrefix, Left(P_Pat_Id, 1), vbTextCompare) > 0 Then
''               strIdType = "LAB"
'            End If
'         End If
'      End If
'
'      Set PID_1 = .createElement("PID.3")
'      Set PID_2 = .createElement("CX.1")
'      Set nTxt = .createTextNode(P_Pat_Id)
'      PID_2.appendChild nTxt
'      PID_1.appendChild PID_2
'
'      If srcCode <> "" Then
'         Set PID_2 = .createElement("CX.4")
'         Set PID_3 = .createElement("HD.1")
'         Set nTxt = .createTextNode(srcExpand)
'         PID_3.appendChild nTxt
'         PID_2.appendChild PID_3
'         PID_1.appendChild PID_2
'      End If
      
      Set PID_2 = .createElement("CX.5")
      Set nTxt = .createTextNode(strIdType)
      PID_2.appendChild nTxt
      PID_1.appendChild PID_2
      
      PID.appendChild PID_1
      
      Set PID_1 = .createElement("PID.5")
      Set PID_2 = .createElement("XPN.1")
      Set PID_3 = .createElement("FN.1")
      
      Set nTxt = .createTextNode(P_Surname)
      PID_3.appendChild nTxt
      PID_2.appendChild PID_3
      PID_1.appendChild PID_2
      
      Set PID_2 = .createElement("XPN.2")
      Set nTxt = .createTextNode(P_Forename)
      PID_2.appendChild nTxt
      PID_1.appendChild PID_2
      
      If Trim(RS!Midname & "") <> "" Then
         Set PID_2 = .createElement("XPN.3")
         Set nTxt = .createTextNode(P_Midname)
         PID_2.appendChild nTxt
         PID_1.appendChild PID_2
      End If
      
      If Trim(RS!Title & "") <> "" Then
         Set PID_2 = .createElement("XPN.5")
         Set nTxt = .createTextNode(P_Title)
         PID_2.appendChild nTxt
         PID_1.appendChild PID_2
      End If
      
      Set PID_2 = .createElement("XPN.7")
      Set nTxt = .createTextNode("L")
      PID_2.appendChild nTxt
      PID_1.appendChild PID_2
      PID.appendChild PID_1
      
      If RS!Date_Of_Birth <> "" Then
         Set PID_1 = .createElement("PID.7")
         Set PID_2 = .createElement("TS.1")
         Set nTxt = .createTextNode(Format(RS!Date_Of_Birth, "yyyymmdd"))
         PID_2.appendChild nTxt
         PID_1.appendChild PID_2
         PID.appendChild PID_1
      End If
      
      Select Case RS!Sex
         Case 0
            strTemp = "U"
         
         Case 1
            strTemp = "M"
            
         Case 2
            strTemp = "F"
            
         Case 9
            strTemp = "S"
      End Select
            
      Set PID_1 = .createElement("PID.8")
      Set nTxt = .createTextNode(strTemp)
      PID_1.appendChild nTxt
      PID.appendChild PID_1
      
      Set PID_1 = .createElement("PID.11")
      
      If P_Pat_Addr_Line1 <> "" Then
         Set PID_2 = .createElement("XAD.1")
         Set PID_3 = .createElement("SAD.1")
         Set nTxt = .createTextNode(P_Pat_Addr_Line1)
         PID_3.appendChild nTxt
         PID_2.appendChild PID_3
         PID_1.appendChild PID_2
      End If
      
      If P_Pat_Addr_Line2 <> "" Then
         Set PID_2 = .createElement("XAD.2")
         Set nTxt = .createTextNode(P_Pat_Addr_Line2)
         PID_2.appendChild nTxt
         PID_1.appendChild PID_2
      End If
      
      If P_Pat_Addr_Line3 <> "" Then
         Set PID_2 = .createElement("XAD.3")
         Set nTxt = .createTextNode(P_Pat_Addr_Line3)
         PID_2.appendChild nTxt
         PID_1.appendChild PID_2
      End If
      
      If P_Pat_Addr_Line4 <> "" Then
         Set PID_2 = .createElement("XAD.4")
         Set nTxt = .createTextNode(P_Pat_Addr_Line4)
         PID_2.appendChild nTxt
         PID_1.appendChild PID_2
      End If
      
      If P_Pat_Addr_Line5 <> "" Then
         Set PID_2 = .createElement("XAD.5")
         Set nTxt = .createTextNode(P_Pat_Addr_Line5)
         PID_2.appendChild nTxt
         PID_1.appendChild PID_2
      End If
      
      PID.appendChild PID_1
      
      If P_Pat_PostCode <> "" Then
         Set PID_1 = .createElement("PID.12")
         Set nTxt = .createTextNode(P_Pat_PostCode)
         PID_1.appendChild nTxt
      End If
      
      PID.appendChild PID_1
   End With
   
   RS.Close
   Set RS = Nothing
      
   Set PID_Element = PID
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.PID_Element"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function PV1_Element(RepId As Long) As MSXML2.IXMLDOMElement
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strArray() As String
   'Dim blnHL7 As Boolean
   Dim PV1 As MSXML2.IXMLDOMElement
   Dim PV1_1 As MSXML2.IXMLDOMElement
   Dim PV1_2 As MSXML2.IXMLDOMElement
   Dim PV1_3 As MSXML2.IXMLDOMElement
   
   Dim pLoc As String
   Dim RecipOrg As String
   Dim ReqOrg As String
   Dim DiscCode As String
   Dim hpData As New clsHealthParties
     
'   With iceCmd
'      .CommandText = "ICELABCOMM_Report_HealthParties"
'      .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , ltIndex)
'      Set RS = .Execute
'   End With
   
   Recipient = ""
   Requestor = ""
   repClin = ""
   reqClin = ""
   
   hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
   hpData.LTSIndex = repListRS!EDI_LTS_Index
   hpData.Read RepId, True, False
   
   RecipOrg = repListRS!EDI_Loc_Nat_Code_To
   
   Recipient = hpData.HP902Code
   repClin = hpData.HP902Name
   
   Requestor = hpData.HP906Code
   reqClin = hpData.HP906Name
   
   orgCode = msgControl.OrganisationCode
   
   If ReqOrg = "" Then
      If Requestor = Recipient Then
         ReqOrg = RecipOrg
      Else
         ReqOrg = orgCode
      End If
   End If
   
   strSQL = "SELECT Clinic_Name " & _
            "From Location " & _
               "INNER JOIN Service_Reports " & _
               "ON Location_Index = Location_IDX " & _
            "WHERE Service_Report_Index=2429483"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   If RS.EOF Then
      pLoc = "Unknown"
   Else
      pLoc = RS(0)
   End If
   
   RS.Close
   
'   blnHL7 = RS.EOF
'   RS.Close
'
'   If blnHL7 Then
'      strSQL = "SELECT Source_App, Source_Expansion " & _
'               "FROM ICEHL7History.dbo.HL7_History hl7 " & _
'                  "INNER JOIN ICEHL7History.dbo.HL7_SourceMap hSrc " & _
'                  "ON hl7.Source_App = hSrc.Source_Code " & _
'                  "INNER JOIN Service_ImpExp_Headers sh " & _
'                     "INNER JOIN Service_ImpExp_Messages sm " & _
'                     "ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id " & _
'                  "ON ImpExp_File = hl7.HL7_Index " & _
'               "WHERE IsNumeric(ImpExp_File) = 1 " & _
'                  "AND Service_Report_Index = " & RepId
'      RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
'      If RS.EOF Then
'         srcCode = ""
'      Else
'         srcCode = RS(0)
'      End If
'      RS.Close
'   End If
   
   With xmlDoc
      Set PV1 = .createElement("PV1")
      
      Set PV1_1 = .createElement("PV1.2")
      Set nTxt = .createTextNode("O")
      PV1_1.appendChild nTxt
      PV1.appendChild PV1_1
      
      Set PV1_1 = .createElement("PV1.3")
      
      If srcCode <> "" Then
         Set PV1_2 = .createElement("PL.1")
         Set nTxt = .createTextNode(srcCode)
         PV1_2.appendChild nTxt
         PV1_1.appendChild PV1_2
      End If
      
      Set PV1_2 = .createElement("PL.4")
      Set PV1_3 = .createElement("HD.1")
      Set nTxt = .createTextNode(ReqOrg)
      PV1_3.appendChild nTxt
      PV1_2.appendChild PV1_3
      
      Set PV1_3 = .createElement("HD.2")
      Set nTxt = .createTextNode(pLoc)
      PV1_3.appendChild nTxt
      PV1_2.appendChild PV1_3
      
      PV1_1.appendChild PV1_2
      
      If srcCode <> "" Then
         Set PV1_2 = .createElement("PL.6")
         Set nTxt = .createTextNode("D")
         PV1_2.appendChild nTxt
         PV1_1.appendChild PV1_2
      End If
      
      PV1.appendChild PV1_1
      
      '****************
      '  Recipient
      '****************
      Set PV1_1 = .createElement("PV1.7")
      
      Set PV1_2 = .createElement("XCN.1")
      Set nTxt = .createTextNode(Recipient)
      PV1_2.appendChild nTxt
      PV1_1.appendChild PV1_2
      
      Set PV1_2 = .createElement("XCN.2")
      Set nTxt = .createTextNode(repClin)
      PV1_2.appendChild nTxt
      PV1_1.appendChild PV1_2
      
      PV1.appendChild PV1_1
         
      '************************
      '  Requestor
      '************************
      Set PV1_1 = .createElement("PV1.8")
      
      Set PV1_2 = .createElement("XCN.1")
      Set nTxt = .createTextNode(Requestor)
      PV1_2.appendChild nTxt
      PV1_1.appendChild PV1_2
      
      Set PV1_2 = .createElement("XCN.2")
      Set nTxt = .createTextNode(reqClin)
      PV1_2.appendChild nTxt
      PV1_1.appendChild PV1_2
      PV1.appendChild PV1_1
      
'      If blnIncludeRepId Then
'         Set PV1_1 = .createElement("PV1.19")
'         Set PV1_2 = .createElement("CX.1")
'         Set nTxt = .createTextNode(RepId)
'         PV1_2.appendChild nTxt
'         PV1_1.appendChild PV1_2
'         PV1.appendChild PV1_1
'      End If
   End With
   
   Set RS = Nothing
   Set PV1_Element = PV1
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.PV1_Element"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function MSH_Element() As MSXML2.IXMLDOMElement
   Dim rootEl As MSXML2.IXMLDOMElement
   Dim MSH As MSXML2.IXMLDOMElement
   Dim MSH_1 As MSXML2.IXMLDOMElement
   Dim MSH_2 As MSXML2.IXMLDOMElement
   Dim uStr() As String
   Dim nTxt As MSXML2.IXMLDOMText
   Dim nAtt As MSXML2.IXMLDOMAttribute
   Dim I As Integer
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "SELECT * " & _
            "FROM EDI_Local_Trader_Settings " & _
            "WHERE EDI_LTS_Index = " & repListRS!EDI_LTS_Index
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      
   hl7Sender = RS!EDI_Name
   
   RS.Close
   Set RS = Nothing
   
   With xmlDoc
      Set MSH = .createElement("MSH")
      
      Set MSH_1 = .createElement("MSH.1")
      Set nTxt = .createTextNode(fs)
      MSH_1.appendChild nTxt
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.2")
      Set nTxt = .createTextNode(cs & rc & ec & ss)
      MSH_1.appendChild nTxt
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.3")
      Set MSH_2 = .createElement("HD.1")
      Set nTxt = .createTextNode("ICE")
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      
      Set MSH_2 = .createElement("HD.2")
      Set nTxt = .createTextNode(repListRS!service_Report_Type)
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.4")
      Set MSH_2 = .createElement("HD.1")
      If blnUseOrgName Then
         Set nTxt = .createTextNode(msgControl.Organisation)
      Else
         Set nTxt = .createTextNode(repListRS!EDI_Provider_Org)
      End If
      
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      
      Set MSH_2 = .createElement("HD.2")
      Set nTxt = .createTextNode(hl7Sender)
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.6")
      Set MSH_2 = .createElement("HD.1")
      Set nTxt = .createTextNode(msgControl.NationalCode)
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.7")
      Set MSH_2 = .createElement("TS.1")
      Set nTxt = .createTextNode(Format(Now(), "yyyymmddhhnn"))
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.9")
      Set MSH_2 = .createElement("MSG.1")
      Set nTxt = .createTextNode("ORU")
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      
      Set MSH_2 = .createElement("MSG.2")
      Set nTxt = .createTextNode("R01")
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.10")
      Set nTxt = .createTextNode(Trim(Str(msgControl.CurrentInterchange)))
      MSH_1.appendChild nTxt
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.11")
      Set MSH_2 = .createElement("PT.1")
      Set nTxt = .createTextNode(Trim(procID))
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      MSH.appendChild MSH_1
      
      Set MSH_1 = .createElement("MSH.12")
      Set MSH_2 = .createElement("VID.1")
      Set nTxt = .createTextNode(Vers)
      MSH_2.appendChild nTxt
      MSH_1.appendChild MSH_2
      MSH.appendChild MSH_1
      
      '.appendChild MSH
   End With
   
   Set MSH_Element = MSH
End Function

Public Property Get Notification() As Boolean
   Notification = False
End Property

Private Function NTENode(NTENo As Integer, _
                         NTEType As String, _
                         strComment As String) As MSXML2.IXMLDOMElement
   Dim NTE As MSXML2.IXMLDOMElement
   Dim NTE_1 As MSXML2.IXMLDOMElement
   
   With xmlDoc
      Set NTE = .createElement("NTE")
      
      Set NTE_1 = .createElement("NTE.1")
      Set nTxt = .createTextNode(NTENo)
      NTE_1.appendChild nTxt
      NTE.appendChild NTE_1
      
      Set NTE_1 = .createElement("NTE.2")
      Set nTxt = .createTextNode(NTEType)
      NTE_1.appendChild nTxt
      NTE.appendChild NTE_1
      
      'If strComment <> "" Then
         Set NTE_1 = .createElement("NTE.3")
         Set nTxt = .createTextNode(strComment)
         NTE_1.appendChild nTxt
         NTE.appendChild NTE_1
      'End If
   End With
   
   Set NTENode = NTE
End Function

Private Function OBR_Element(ByRef sampRS As ADODB.Recordset, _
                             ByRef invRS As ADODB.Recordset) As MSXML2.IXMLDOMElement
   On Error GoTo procEH
   Dim RS As ADODB.Recordset
   Dim OBR As MSXML2.IXMLDOMElement
   Dim OBR_1 As MSXML2.IXMLDOMElement
   Dim OBR_2 As MSXML2.IXMLDOMElement
   Dim OBR_3 As MSXML2.IXMLDOMElement
   
   Dim Status As Long
   Dim strTemp As String
   Dim fStr As String
   Dim DiscCode As String
   Dim sampRange As String
   Dim testDesc As String
   Dim testCount As Integer
   Dim strArray() As String
   Dim I As Integer
   Dim blnAbnormal As Boolean
      
   If blnUseSampleForDisc Then
      DiscCode = ""
   Else
      DiscCode = Get_DiscCode(repListRS!service_Report_Type)
   End If
   
   If DiscCode = "" Then
      DiscCode = GetSampleLab(sampRS!Sample_Id, blnUseSampleForDisc)
   End If
   
   SampId = sampRS!Sample_Id
   
   Set OBR = Nothing
         
   'OBRNo = 1
      
   eClass.FurtherInfo = "Sample Index = " & sampRS!Sample_Index & " Investigation Index = " & invRS!Investigation_Index
   testDesc = Replace(invRS!EDI_Local_Test_Code & "", Chr(0), "") & " (" & invRS!Investigation_Requested & ") "
         
   If invRS!EDI_OP_Suppress Then
      Status = IS_INV_SUPPRESSED
      msgControl.LogReportMessage IS_INV_SUPPRESSED, testDesc & "Battery Header suppressed - no output for this investigation"
   Else
      
      strTemp = Replace(invRS!Investigation_Code, Chr(0), "")
      
      If sampRS!Collection_Text = "Fasting" Then
         fStr = "F"
      End If
            
      With xmlDoc
         Set OBR = .createElement("OBR")
         
         Set OBR_1 = .createElement("OBR.1")
         Set nTxt = .createTextNode(OBRNo)
         OBR_1.appendChild nTxt
         OBR.appendChild OBR_1
         OBRNo = OBRNo + 1
         
         Set OBR_1 = .createElement("OBR.3")
         Set OBR_2 = .createElement("EI.1")
         Set nTxt = .createTextNode(sampRS!Sample_Id)
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
   
         Set OBR_2 = .createElement("EI.2")
         If srcCode = "" Then
            Set nTxt = .createTextNode(GetSampleLab(sampRS!Sample_Id))
         Else
            Set nTxt = .createTextNode(srcExpand)
         End If
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         
'         If blnHL7 Then
'            Set OBR_2 = .createElement("EI.3")
'            Set nTxt = .createTextNode("X-RAY")
'            OBR_2.appendChild nTxt
'            OBR_1.appendChild OBR_2
'         End If
         
         OBR.appendChild OBR_1
   
         Set OBR_1 = .createElement("OBR.4")
         Set OBR_2 = .createElement("CE.1")
         Set nTxt = .createTextNode(Replace(invRS!Investigation_Code, Chr(0), ""))
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
   
         Set OBR_2 = .createElement("CE.2")
         Set nTxt = .createTextNode(invRS!Investigation_Requested)
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         
         Set OBR_2 = .createElement("CE.3")
         Set nTxt = .createTextNode("LC")
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         
         If invRS!EDI_OP_Active And (Not IsNull(invRS!Read_V2RC)) Then
            Set OBR_2 = .createElement("CE.4")
            Set nTxt = .createTextNode(invRS!Read_V2RC)
            OBR_2.appendChild nTxt
            OBR_1.appendChild OBR_2
            
            Set OBR_2 = .createElement("CE.5")
            Set nTxt = .createTextNode(invRS!Read_V2Rubric)
            OBR_2.appendChild nTxt
            OBR_1.appendChild OBR_2
            
            Set OBR_2 = .createElement("CE.6")
            Set nTxt = .createTextNode("RC")
            OBR_2.appendChild nTxt
            OBR_1.appendChild OBR_2
         End If
         
         OBR.appendChild OBR_1
         
         Set OBR_1 = .createElement("OBR.7")
         Set OBR_2 = .createElement("TS.1")
         Set nTxt = .createTextNode(Format(sampRS!Collection_DateTime, "yyyymmddhhmm"))
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         OBR.appendChild OBR_1
         
         If ClinDets <> "" Then
            Set OBR_1 = .createElement("OBR.13")
            Set nTxt = .createTextNode(ClinDets)
            OBR_1.appendChild nTxt
            OBR.appendChild OBR_1
            ClinDets = ""
            SampleCmts = ""
         End If
         
         Set OBR_1 = .createElement("OBR.14")
         Set OBR_2 = .createElement("TS.1")
         Set nTxt = .createTextNode(Format(sampRS!Collection_DateTimeReceived, "yyyymmddhhmm"))
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         OBR.appendChild OBR_1
         
         Set OBR_1 = .createElement("OBR.15")
         Set OBR_2 = .createElement("SPS.1")
         Set OBR_3 = .createElement("CE.1")
         Set nTxt = .createTextNode(Trim(sampRS!Sample_Code))
         OBR_3.appendChild nTxt
         OBR_2.appendChild OBR_3
         OBR_1.appendChild OBR_2
         
         Set OBR_2 = .createElement("SPS.2")
         Set nTxt = .createTextNode(sampRS!Sample_Text)
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         OBR.appendChild OBR_1
         
         Set OBR_1 = .createElement("OBR.16")
         Set OBR_2 = .createElement("XCN.1")
         Set nTxt = .createTextNode(Recipient)
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         
         Set OBR_2 = .createElement("XCN.2")
         Set OBR_3 = .createElement("FN.1")
         
         strArray = Split(repClin, ":")
         
         If UBound(strArray) > 0 Then
            Set nTxt = .createTextNode(strArray(0))
            OBR_3.appendChild nTxt
            OBR_2.appendChild OBR_3
            OBR_1.appendChild OBR_2
            
            For I = 1 To UBound(strArray)
               Set OBR_2 = .createElement("XCN." & (I + 2))
               Set nTxt = .createTextNode(strArray(I))
               OBR_2.appendChild nTxt
               OBR_1.appendChild OBR_2
            Next I
                        
         Else
            Set nTxt = .createTextNode(repClin)
            OBR_3.appendChild nTxt
            OBR_2.appendChild OBR_3
            OBR_1.appendChild OBR_2
         End If
         
         OBR.appendChild OBR_1
                  
         Set OBR_1 = .createElement("OBR.22")
         Set OBR_2 = .createElement("TS.1")
         Set nTxt = .createTextNode(Format(repListRS!Report_Date, "yyyymmddhhmm"))
         OBR_2.appendChild nTxt
         OBR_1.appendChild OBR_2
         OBR.appendChild OBR_1
         
         Set OBR_1 = .createElement("OBR.24")
         Set nTxt = .createTextNode(repListRS!service_Report_Type)
         
         OBR_1.appendChild nTxt
         OBR.appendChild OBR_1
      
         strTemp = IIf((Mid(sampRS!Status, 1, 1) = "" Or Mid(sampRS!Status, 1, 1) = "U"), "", Mid(sampRS!Status, 1, 1))
         
         If strTemp <> "" Then
            Set OBR_1 = .createElement("OBR.25")
            Set nTxt = .createTextNode(strTemp)
            OBR_1.appendChild nTxt
            OBR.appendChild OBR_1
         End If
         
         If Requestor <> Recipient Then
            Set OBR_1 = .createElement("OBR.28")
            Set OBR_2 = .createElement("XCN.1")
            Set nTxt = .createTextNode(Requestor)
            OBR_2.appendChild nTxt
            OBR_1.appendChild OBR_2
            
            Set OBR_2 = .createElement("XCN.2")
            Set OBR_3 = .createElement("FN.1")
         
            strArray = Split(reqClin, ":")
            If UBound(strArray) > 0 Then
               Set nTxt = .createTextNode(strArray(0))
               OBR_3.appendChild nTxt
               OBR_2.appendChild OBR_3
               OBR_1.appendChild OBR_2
               
               For I = 1 To UBound(strArray)
                  Set OBR_2 = .createElement("XCN." & (I + 1))
                  Set nTxt = .createTextNode(strArray(I))
                  OBR_2.appendChild nTxt
                  OBR_1.appendChild OBR_2
               Next I
               
            Else
               Set nTxt = .createTextNode(reqClin)
               OBR_3.appendChild nTxt
               OBR_2.appendChild OBR_3
               OBR_1.appendChild OBR_2
            End If
         
            OBR.appendChild OBR_1
         
         End If
      End With
   End If
   Set OBR_Element = OBR
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.OBR_Element"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function OBX_Element(ByRef resRS As ADODB.Recordset, _
                             SampleIndex As Long, _
                             OBXNo As Integer) As MSXML2.IXMLDOMElement
   On Error GoTo procEH
   Dim OBX As MSXML2.IXMLDOMElement
   Dim OBX_1 As MSXML2.IXMLDOMElement
   Dim OBX_2 As MSXML2.IXMLDOMElement
   Dim OBX_3 As MSXML2.IXMLDOMElement
   Dim testDesc As String
   Dim testCount As Integer
   Dim strTemp As String
   Dim strStatus As String
   Dim rangeStr As String
   Dim rPosn As Integer
   Dim resOOR() As String
   Dim strHiLo As String
   Dim resValue As String
   
   testDesc = Replace(resRS!Result_Code, Chr(0), "") & " (" & resRS!Result_Rubric & ") "
   
   Set OBX = Nothing
   
   If resRS!EDI_OP_Suppress Then
      Status = IS_TEST_SUPPRESSED
      msgControl.LogReportMessage Status, testDesc & "Test suppressed"
   Else
            
      With xmlDoc
         Set OBX = .createElement("OBX")
         
         Set OBX_1 = .createElement("OBX.1")
         Set nTxt = .createTextNode(OBXNo)
         OBX_1.appendChild nTxt
         OBX.appendChild OBX_1
         
         If IsNumeric(resRS!Result) Then
            strStatus = "NM"
         Else
            strStatus = "ST"
         End If
            
         rangeStr = ""
         
         If (resRS!Lower_Range & resRS!Upper_Range & "") <> "" Then
            If resRS!Lower_Range & "" = "" Then
               rangeStr = "<" & resRS!Upper_Range
            ElseIf resRS!Upper_Range & "" = "" Then
               rangeStr = ">" & resRS!Lower_Range
            Else
               rangeStr = resRS!Lower_Range & "-" & resRS!Upper_Range
            End If
         End If
         
         Set OBX_1 = .createElement("OBX.2")
         Set nTxt = .createTextNode(strStatus)
         OBX_1.appendChild nTxt
         OBX.appendChild OBX_1

         Set OBX_1 = .createElement("OBX.3")
         Set OBX_2 = .createElement("CE.1")
         Set nTxt = .createTextNode(Replace(resRS!Result_Code, Chr(0), ""))
         OBX_2.appendChild nTxt
         OBX_1.appendChild OBX_2
         

         Set OBX_2 = .createElement("CE.2")
         Set nTxt = .createTextNode(resRS!Result_Rubric)
         OBX_2.appendChild nTxt
         OBX_1.appendChild OBX_2
         
         Set OBX_2 = .createElement("CE.3")
         Set nTxt = .createTextNode("LC")
         OBX_2.appendChild nTxt
         OBX_1.appendChild OBX_2
         OBX.appendChild OBX_1
         
         If resRS!EDI_OP_Active And (Not IsNull(resRS!Read_V2RC)) Then
            Set OBX_2 = .createElement("CE.4")
            Set nTxt = .createTextNode(resRS!Read_V2RC)
            OBX_2.appendChild nTxt
            OBX_1.appendChild OBX_2
            OBX.appendChild OBX_1
            
            Set OBX_2 = .createElement("CE.5")
            Set nTxt = .createTextNode(resRS!Read_V2Rubric)
            OBX_2.appendChild nTxt
            OBX_1.appendChild OBX_2
            OBX.appendChild OBX_1
            
            Set OBX_2 = .createElement("CE.6")
            Set nTxt = .createTextNode("RC")
            OBX_2.appendChild nTxt
            OBX_1.appendChild OBX_2
            OBX.appendChild OBX_1
         End If
         
         If resRS!Result & "" <> "" Then
            Set OBX_1 = .createElement("OBX.5")
            
            If blnAmendUOM Then
               Set nTxt = .createTextNode(Replace(resRS!Result, "^", "E"))
            Else
               Set nTxt = .createTextNode(resRS!Result)
            End If
            
            OBX_1.appendChild nTxt
            OBX.appendChild OBX_1
         End If
         
         'If resRS!UOM_Text & "" <> "" Then
            Set OBX_1 = .createElement("OBX.6")
            Set OBX_2 = .createElement("CE.1")
            
            If resRS!EDI_Nat_UOM & "" = "" Then
               Set nTxt = .createTextNode(resRS!UOM_Code)
            Else
               Set nTxt = .createTextNode(resRS!EDI_Nat_UOM)
            End If
            
            OBX_2.appendChild nTxt
            OBX_1.appendChild OBX_2
            OBX.appendChild OBX_1
         'End If
         
         If rangeStr <> "" Then
            Set OBX_1 = .createElement("OBX.7")
            Set nTxt = .createTextNode(rangeStr)
            OBX_1.appendChild nTxt
            OBX.appendChild OBX_1
         End If
         
         msgData.AdjustRangeData Trim(resRS!Lower_Range & ""), Trim(resRS!Upper_Range & "")
         
         strHiLo = ""
         If resRS!Abnormal_Flag Then
            resOOR = Split(msgData.AnalyseResultString(resRS!Result), "|")
            resValue = resOOR(0)
            'resValue = Split(msgData.AnalyseResultString(resRS!Result), "|")(0)
            
            If resOOR(1) <> "" Then
               If resOOR(1) = "6" Then
                  strHiLo = "H"
               ElseIf resOOR(1) = "7" Then
                  strHiLo = "L"
               End If
            Else
               If IsNumeric(resValue) Then
                  If Val(resValue) < Val(msgData.LowRange) Then
                     strHiLo = "L"
                  ElseIf Val(resValue) > Val(msgData.HighRange) Then
                     strHiLo = "H"
                  End If
               Else
                  strHiLo = "A"
               End If
            End If
         End If

         If strHiLo <> "" Then
            Set OBX_1 = .createElement("OBX.8")
            Set nTxt = .createTextNode(strHiLo)
            OBX_1.appendChild nTxt
            OBX.appendChild OBX_1
         End If
   
         strTemp = resRS!Status
         If resRS!Status = "" Then
            strTemp = "F"
         ElseIf XMLPrelimId <> "" Then
            If strTemp = "PR" Then
               strTemp = XMLPrelimId
            End If
         End If
         
         Set OBX_1 = .createElement("OBX.11")
         Set nTxt = .createTextNode(strTemp)
         OBX_1.appendChild nTxt
         OBX.appendChild OBX_1
         
         Set OBX_1 = .createElement("OBX.14")
         Set OBX_2 = .createElement("TS.1")
         Set nTxt = .createTextNode(Format(repListRS!Report_Date, "yyyymmddhhmm"))
         OBX_2.appendChild nTxt
         OBX_1.appendChild OBX_2
         OBX.appendChild OBX_1
   
'*****************************************************
'        Virus Lab Test
'*****************************************************
'         If XMLVirusPrefix <> "" Then
'            If UCase(Left(SampId, 2)) = UCase(XMLVirusPrefix) Then
'
'               If IsNumeric(XMLVirusRangeStart) And IsNumeric(XMLVirusRangeEnd) Then
'                  If Mid(SampId, 3, Len(SampId) - 3) >= XMLVirusRangeStart Then
'                     If Mid(SampId, 3, Len(SampId) - 3) <= XMLVirusRangeEnd Then
                     
'         If blnExternalLab Then
'            Set OBX_1 = .createElement("OBX.15")
'
'            Set OBX_2 = .createElement("CE.1")
'            If XMLVirusCode = "" Then
'               Set nTxt = .createTextNode(SampId)
'            Else
'               Set nTxt = .createTextNode(XMLVirusCode)
'            End If
'
'            OBX_2.appendChild nTxt
'            OBX_1.appendChild OBX_2
'
'            Set OBX_2 = .createElement("CE.2")
'            Set nTxt = .createTextNode(XMLVirusLab)
'            OBX_2.appendChild nTxt
'            OBX_1.appendChild OBX_2
'
'            Set OBX_2 = .createElement("CE.3")
'            Set nTxt = .createTextNode("L")
'            OBX_2.appendChild nTxt
'            OBX_1.appendChild OBX_2
'
'            OBX.appendChild OBX_1
'         End If
''                  End If
'               End If
'            End If
'         End If
   
      End With
   End If
   
   Set OBX_Element = OBX
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.OBX_Element"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Sub ReadReport(RepId As Long)
   On Error GoTo procEH
   Dim totNodes As Long
   Dim msgRoot As MSXML2.IXMLDOMElement
   Dim ORU As MSXML2.IXMLDOMElement
   Dim ORU_1 As MSXML2.IXMLDOMElement
   Dim ORU_2 As MSXML2.IXMLDOMElement
   Dim ORU_3 As MSXML2.IXMLDOMElement
   
   Dim sampRS As New ADODB.Recordset
   Dim invRS As New ADODB.Recordset
   Dim resRS As New ADODB.Recordset
   Dim RS As New ADODB.Recordset
   Dim blnAbnormal As Boolean
   Dim NTENo As Integer
   Dim OBXNo As Integer
   
   Set iceCmd = Nothing
   Set iceCmd = New ADODB.Command
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
   End With
   
   With xmlDoc
      Set msgRoot = .createElement("ORU_R01")
'      msgRoot.appendChild MSH_Element
      
      Set ORU = .createElement("ORU_R01.PATIENT_RESULT")
      Set ORU_1 = .createElement("ORU_R01.PATIENT")
      
'**********************************************
'     PID
'**********************************************
      ORU_1.appendChild PID_Element(RepId)
      
'**********************************************
'     Report Level Comments
'**********************************************
      With iceCmd
         .CommandText = "ICELABCOMM_Report_Comments"
         Set RS = .Execute
      End With
      
      Set ORU_2 = xmlDoc.createElement("NTE")
      ClinDets = ""
      SampleCmts = ""
      
      NTENo = 1
      Do Until RS.EOF
         Select Case RS!Comment_Type
            Case "R"
               ORU_1.appendChild NTENode(NTENo, "C", RS!Service_Report_Comment)
               NTENo = NTENo + 1
            
            Case "D"
               If CDasNTE Then
                  ORU_1.appendChild NTENode(NTENo, "C", RS!Service_Impexp_Comment)
                  NTENo = NTENo + 1
               Else
                  If ClinDets = "" Then
                     ClinDets = RS!Service_Report_Comment
                  Else
                      ClinDets = ClinDets & RS!Service_Report_Comment
                  End If
               End If
               
            Case "S"
               If SCasNTE Then
                  ORU_1.appendChild NTENode(NTENo, "C", RS!Service_Impexp_Comment)
                  NTENo = NTENo + 1
               Else
                  If SampleCmts = "" Then
                     SampleCmts = RS!Service_Report_Comment
                  Else
                      'SampleCmts = SampleCmts & ".br" & RS!Service_Report_Comment
                      SampleCmts = SampleCmts & vbCrLf & RS!Service_Report_Comment
                  End If
               End If
         
         End Select
         
         RS.MoveNext
      Loop
         
      If SampleCmts <> "" Then
          If ClinDets <> "" Then
              'ClinDets = ClinDets & ".br" & SampleCmts
              ClinDets = ClinDets & vbCrLf & SampleCmts
          Else
              ClinDets = SampleCmts
          End If
      End If
      
      RS.Close
      
'**********************************************
'     PV1
'**********************************************
      Set ORU_2 = .createElement("ORU_R01.PATIENT_VISIT")
      ORU_2.appendChild PV1_Element(RepId)
      
      ORU_1.appendChild ORU_2
      ORU.appendChild ORU_1
      
'**************************
'     OBR
'**************************
      
      With iceCmd
         .CommandText = "ICELABCOMM_Report_Sample"
         Set sampRS = .Execute
      End With
      
      totNodes = ORU.childNodes.Length
      OBRNo = 1
      Do Until sampRS.EOF
         With iceCmd
            .CommandText = IIf(blnUseRCIndex, "ICELABCOMM_Report_Invest_By_Index", "ICELABCOMM_Report_Invest_By_Code")
            .Parameters(0).Value = RepId
            .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
            .Parameters.Append .CreateParameter("Samp", adInteger, adParamInput, , sampRS!Sample_Index)
            .Parameters.Append .CreateParameter("Ab", adBoolean, adParamOutput, , blnAbnormal)
            Set invRS = .Execute
            .Parameters.Delete 3
            .Parameters.Delete 2
            .Parameters.Delete 1
         End With
         
         invRS.Filter = "EDI_OP_Suppress = 0"
         If invRS.RecordCount = 0 Then
            Status = IS_INV_SUPPRESSED
            msgControl.LogReportMessage Status, invRS!Investigation_Code & " (" & invRS!Investigation_Requested & _
                                                 ") battery header output suppressed"
         Else
            Do Until invRS.EOF
               Set ORU_1 = .createElement("ORU_R01.ORDER_OBSERVATION")
               Set ORU_2 = OBR_Element(sampRS, invRS)
               
               If Not ORU_2 Is Nothing Then
                  ORU_1.appendChild ORU_2
                  
                  With iceCmd
                     .CommandText = "ICELABCOMM_Report_InvestComments"
                     .Parameters(0) = invRS!Investigation_Index
                     Set RS = .Execute
                  End With
                  
                  NTENo = 1
                  Do Until RS.EOF
                     If Trim(RS!Service_Investigation_Comment) <> "" Then
                        ORU_1.appendChild NTENode(NTENo, "I", RS!Service_Investigation_Comment)
                        NTENo = NTENo + 1
                     End If
                     RS.MoveNext
                  Loop
                  
                  RS.Close
               
'**************************************
'                 OBX
'**************************************
                  With iceCmd
                     .CommandText = IIf(blnUseRCIndex, "ICELABCOMM_Report_Results_By_Index", "ICELABCOMM_Report_Results_By_Code")
                     .Parameters(0).Value = invRS!Investigation_Index
                     .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
                     Set resRS = .Execute
                     .Parameters.Delete 1
                  End With
                  
                  resRS.Filter = "EDI_OP_Suppress = 0"
               
                  If resRS.RecordCount = 0 Then
                     Status = IS_INV_SUPPRESSED
                     msgControl.LogReportMessage Status, invRS!Investigation_Code & " (" & invRS!Investigation_Requested & _
                                                 ") ALL the tests in this battery have their output suppressed"
                  Else
                     
                     If resRS.EOF Then
                        Set ORU_2 = .createElement("ORU_R01.OBSERVATION")
                        ORU_1.appendChild ORU_2
                     End If
                     
                     OBXNo = 0
                     Do Until resRS.EOF
                        Set ORU_2 = .createElement("ORU_R01.OBSERVATION")
                        OBXNo = OBXNo + 1
                        Set ORU_3 = OBX_Element(resRS, sampRS!Sample_Index, OBXNo)
                        If Not ORU_3 Is Nothing Then
                           ORU_2.appendChild ORU_3
                           
'********************************************
'                          Report Comments
'********************************************
                           With iceCmd
                              .CommandText = "ICELABCOMM_Report_ResultComment"
                              .Parameters(0).Value = resRS!Result_Index
                              Set RS = .Execute
                           End With
                           
                           NTENo = 1
                           Do Until RS.EOF
                              If resRS!Result = "" Or RS!Comment_Type <> "R" Then
                                 ORU_2.appendChild NTENode(NTENo, "L", RS!Service_Result_Comment)
                                 NTENo = NTENo + 1
                              Else
                                 OBXNo = OBXNo + 1
                                 ORU_2.appendChild OBX_Element(resRS, sampRS!Sample_Index, OBXNo)
                              End If
                              
                              RS.MoveNext
                           Loop
                           RS.Close
                           
                           ORU_1.appendChild ORU_2
                        End If
                        
                        resRS.MoveNext
                     Loop
                     
                     resRS.Close
                  End If
                  ORU.appendChild ORU_1
               End If
               
               invRS.MoveNext
            Loop
         End If
         
         invRS.Close
         sampRS.MoveNext
      Loop
      
      sampRS.Close
      
      If ORU.childNodes.Length > totNodes Then
         msgRoot.appendChild ORU
         .appendChild msgRoot
      Else
         msgControl.LogReportMessage MS_NO_OUTPUT, "No data to output - check for any suppression messages"
         msgControl.MessageStatus = MS_NO_OUTPUT
      End If
   End With
   
   
   Set sampRS = Nothing
   Set invRS = Nothing
   Set resRS = Nothing
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Labcomm_HL7_XML.ReadReport"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get ReleaseChar() As String
   ReleaseChar = ""
End Property

Public Property Get RetainGrave() As Boolean
   RetainGrave = False
End Property





