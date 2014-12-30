VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ReportId As Long
Private blnTemplateRead As Boolean
Private LabNo As String
Private Specialty As String
Private HospNo As String
Private pForename As String
Private pSurname As String
Private Clinician As String
Private Destination As String
Private Collected As String
Private Received As String
Private Reported As String
Private SampText As String
Private htmlHeader As String
Private htmlTrailer As String

Private hDoc As MSHTML.HTMLDocument
Private hBody As MSHTML.IHTMLDOMNode
Private hTab As MSHTML.HTMLTable
Private hAtt As MSHTML.IHTMLDOMAttribute
Private hAttGrp As MSHTML.IHTMLAttributeCollection
Private hEL As MSHTML.IHTMLDOMNode
   
Private tRow As MSHTML.IHTMLDOMNode
Private tCell As MSHTML.IHTMLDOMNode
Private td As MSHTML.HTMLTableCell
Private hTxt As MSHTML.IHTMLDOMNode
Private hNode As MSHTML.IHTMLDOMNode

Private Property Get Comments() As String
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim objBuff As StringBuffer
    
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Comments"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   RS.Filter = "Comment_Type = 'D'"
   If RS.RecordCount > 0 Then
      Set objBuff = New StringBuffer
      objBuff.Append "<BR><STRONG>Report Comments:</STRONG><BR>"
      Do Until RS.EOF
          objBuff.Append RS!Service_Report_Comment
          objBuff.Append "<BR>"
          RS.MoveNext
      Loop
      Comments = objBuff.Value
      Set objBuff = Nothing
    End If
    
    RS.Close
    Set objBuff = Nothing
    Set RS = Nothing
    Set iceCmd = Nothing
End Property

Public Function CreateMessage(RepId As Long) As String
   On Error GoTo procEH
   ReportId = RepId
   CreateMessage = ReportBody(RepId)
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.CreateMessage"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Property Get FileHeader() As String
   FileHeader = "<HTML><HEAD><STYLE>" & _
               "TABLE {FONT-SIZE: 10px; FONT-FAMILY: Tahoma}" & _
               ".label {font-Weight: Bold}" & _
               ".oor {COLOR:    red}" & _
               "TR.result {background -COLOR: gainsboro}" & _
               "DIV.samp {LEFT: 20px; POSITION: relative}" & _
               "DIV.inv {LEFT: 40px; POSITION: relative}" & _
               ".comment {BACKGROUND-COLOR: #fffacd;text-indent : 10pt;}" & _
               ".suppress {color: blue;font-weight: bold;}" & _
               "</STYLE></HEAD>" & _
               "<BODY>"

End Property

Public Property Get FileTrailer() As String
   FileTrailer = "</BODY></HTML>"
End Property

Public Property Get InvalidChars() As Variant
   ReDim strArray(2)
   strArray(0) = ""
   strArray(1) = ""
   InvalidChars = strArray
End Property

Public Property Get ReleaseChar() As String
   ReleaseChar = ""
End Property

Public Property Get RetainGrave() As Boolean
   RetainGrave = False
End Property

Private Sub HealthParties()
   Dim hpData As New clsHealthParties
   
   hpData.IndividualIndex = repListRS!EDI_Individual_Index_To
   hpData.LTSIndex = repListRS!EDI_LTS_Index
   hpData.Read repListRS!EDI_Report_Index, True
   
   Clinician = hpData.HP902Name
   
   Destination = repListRS!EDI_Loc_Nat_Code_To
   Set hpData = Nothing
End Sub

Public Property Get Notification() As Boolean
   Notification = False
End Property

Public Function ReportBody(RepId As Long) As String
   On Error GoTo procEH
   
   Set hDoc = New MSHTML.HTMLDocument
   Set hBody = hDoc.Body
   hBody.removeChild hBody.firstChild
   
   hBody.appendChild ReportHeader
   
   Set hEL = hDoc.createElement("hr")
   Set hAttGrp = hEL.Attributes
   Set hAtt = hDoc.createAttribute("width")
   hAtt.nodeValue = "80%"
   hAttGrp.setNamedItem hAtt
   
   hBody.appendChild hEL
   hBody.appendChild ReportSamples(RepId)
   
   Set hEL = hDoc.createElement("br")
   hBody.appendChild hEL
   
   Set hEL = hDoc.createElement("hr")
   Set hAttGrp = hEL.Attributes
   Set hAtt = hDoc.createAttribute("width")
   hAtt.nodeValue = "100%"
   hAttGrp.setNamedItem hAtt
   
   Set hAtt = hDoc.createAttribute("size")
   hAtt.nodeValue = "4"
   hAttGrp.setNamedItem hAtt
   
   Set hAtt = hDoc.createAttribute("noshade")
   hAtt.nodeValue = "yes"
   hAttGrp.setNamedItem hAtt
   
   hBody.appendChild hEL
   
   Set hEL = hDoc.createElement("br")
   hBody.appendChild hEL
   
   
'   Open "C:\Documents and Settings\Bernie_Perrett\My Documents\Projects\HTML Tests\DOMrep.html" For Output As #3
'   Print #3, hDoc.documentElement.outerHTML
'   Close #3
   
   ReportBody = hDoc.Body.innerHTML
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "clsHTML.ReportBody"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function ReportHeader() As MSHTML.IHTMLDOMNode
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim pAdr As String
   Dim hDiv As MSHTML.IHTMLDOMNode
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Patient"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
'  Set up variables used by StandaloneReport as email header
   LabNo = RS!Service_Report_Id
   Specialty = RS!Service_Report_Type
   HospNo = RS!Hospital_Number
   pForename = RS!Forename
   pSurname = RS!Surname
   
   Set hDiv = hDoc.createElement("div")
   
   Set hTab = hDoc.createElement("table")
   Set hAttGrp = hTab.Attributes
   Set hAtt = hDoc.createAttribute("width")
   hAtt.nodeValue = "100%"
   hAttGrp.setNamedItem hAtt
   
   Set tRow = hDoc.createElement("tr")
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Patient Name:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(RS!Forename & " " & RS!Surname)
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("NHS Number:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(IIf(Trim(RS!New_NHS_No & "") = "", "Not Known", RS!New_NHS_No))
   tCell.appendChild hTxt
   tRow.appendChild tCell
   hTab.appendChild tRow
   
   Set tRow = hDoc.createElement("tr")
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Date of Birth:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(RS!Date_Of_Birth & "")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Hospital Number:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(IIf(Trim(RS!Hospital_Number & "") = "", "Not Known", RS!Hospital_Number))
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   hTab.appendChild tRow
   
   Set tRow = hDoc.createElement("tr")
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Address:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   pAdr = RS!Pat_Addr_Line1
   If RS!Pat_Addr_Line2 <> "" Then
       pAdr = pAdr & ", " & RS!Pat_Addr_Line2
   End If
   If RS!Pat_Addr_Line3 <> "" Then
       pAdr = pAdr & ", " & RS!Pat_Addr_Line3
   End If
   If RS!Pat_Addr_Line4 <> "" Then
       pAdr = pAdr & ", " & RS!Pat_Addr_Line4
   End If
   If RS!Pat_Addr_Line5 <> "" Then
       pAdr = pAdr & ", " & RS!Pat_Addr_Line5
   End If
   If RS!Pat_PostCode <> "" Then
       pAdr = pAdr & " " & RS!Pat_PostCode
   End If
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(pAdr)
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   hTab.appendChild tRow

   hDiv.appendChild hTab
   
   Set hEL = hDoc.createElement("hr")
   Set hAttGrp = hEL.Attributes
   Set hAtt = hDoc.createAttribute("width")
   hAtt.nodeValue = "85%"
   hAttGrp.setNamedItem hAtt
   
   hDiv.appendChild hEL
   
   Set hTab = hDoc.createElement("table")
   Set hAttGrp = hTab.Attributes
   Set hAtt = hDoc.createAttribute("width")
   hAtt.nodeValue = "100%"
   hAttGrp.setNamedItem hAtt
   
   Set tRow = hDoc.createElement("tr")
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Lab No:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(RS!Service_Report_Id)
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Specialty:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(RS!Service_Report_Type)
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Destination:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(repListRS!EDI_Loc_Nat_Code_To)
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   hTab.appendChild tRow
   
   Set tRow = hDoc.createElement("tr")
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   Set hTxt = hDoc.createTextNode("Report Date:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   Set tCell = hDoc.createElement("td")
   Set hTxt = hDoc.createTextNode(RS!DateTime_Of_Report)
   tCell.appendChild hTxt
   tRow.appendChild tCell
      
   RS.Close
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "label"
   hAttGrp.setNamedItem hAtt
   
   Set hTxt = hDoc.createTextNode("Clinician:")
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   HealthParties
   
   Set tCell = hDoc.createElement("td")
   Set hAttGrp = tCell.Attributes
   Set hAtt = hDoc.createAttribute("colspan")
   hAtt.nodeValue = "3"
   hAttGrp.setNamedItem hAtt
   
   Set hTxt = hDoc.createTextNode(Clinician)
   tCell.appendChild hTxt
   tRow.appendChild tCell
   
   hTab.appendChild tRow
   
   iceCmd.CommandText = "ICELABCOMM_Report_Comments"
   Set RS = iceCmd.Execute
   
   RS.Filter = "Comment_Type='D'"
   Do Until RS.EOF
      Set tRow = hDoc.createElement("tr")
      Set tCell = hDoc.createElement("td")
      Set hAttGrp = tCell.Attributes
      
      Set hAtt = hDoc.createAttribute("colspan")
      hAtt.nodeValue = "6"
      hAttGrp.setNamedItem hAtt
      
      Set hAtt = hDoc.createAttribute("class")
      hAtt.nodeValue = "comment"
      hAttGrp.setNamedItem hAtt
      
      Set hTxt = hDoc.createTextNode(RS!Service_Report_Comment)
      tCell.appendChild hTxt
      
      tRow.appendChild tCell
      hTab.appendChild tRow
      
      RS.MoveNext
   Loop
   RS.Close
   
   hDiv.appendChild hTab
   
   Set RS = Nothing
   Set iceCmd = Nothing
   
   Set ReportHeader = hDiv
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "clsHTML.ReportHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function ReportInvest(sampID As Long) As MSHTML.IHTMLDOMNode
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim cText As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   Dim hDiv As MSHTML.IHTMLDOMNode
   
   Set iceCmd = New ADODB.Command
   
   If blnUseRCIndex Then
      cText = "ICELABCOMM_Report_Invest_By_Index"
   Else
      cText = "ICELABCOMM_Report_Invest_By_Code"
   End If
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = cText
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
      .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , sampID)
      .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
      Set RS = .Execute
      .Parameters.Delete "Abnormal"
      .Parameters.Delete "SampId"
      .Parameters.Delete "LTSIndex"
   End With
   
   Set hDiv = hDoc.createElement("div")
   Set hAttGrp = hDiv.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "inv"
   hAttGrp.setNamedItem hAtt
   
   Set hTab = hDoc.createElement("table")
   Set hAttGrp = hTab.Attributes
   Set hAtt = hDoc.createAttribute("width")
   hAtt.nodeValue = "95%"
   hAttGrp.setNamedItem hAtt
   
   Do Until RS.EOF
      Set tRow = hDoc.createElement("tr")
      Set tCell = hDoc.createElement("td")
      Set hAttGrp = tCell.Attributes
      Set hAtt = hDoc.createAttribute("colspan")
      hAtt.nodeValue = "5"
      hAttGrp.setNamedItem hAtt
      
      If RS!EDI_OP_Suppress Then
         Set hAtt = hDoc.createAttribute("class")
         hAtt.nodeValue = "suppress"
         hAttGrp.setNamedItem hAtt
      
         Set hTxt = hDoc.createTextNode("Output suppressed for investigation " & RS!Investigation_Requested)
         tCell.appendChild hTxt
         tRow.appendChild tCell
         
         hTab.appendChild tRow
      
      Else
         
         Set hAtt = hDoc.createAttribute("class")
         hAtt.nodeValue = "label"
         hAttGrp.setNamedItem hAtt
      
         If IsNull(RS!Read_V2RC) Then
            Set hTxt = hDoc.createTextNode(RS!Investigation_Requested)
         Else
            Set hTxt = hDoc.createTextNode(RS!Investigation_Requested & " [" & RS!Read_V2RC & "]")
         End If
         
         tCell.appendChild hTxt
         tRow.appendChild tCell
         
         hTab.appendChild tRow
      
         If RS!Comment_Marker Then
            With iceCmd
               .CommandText = "ICELABCOMM_Report_InvestComments"
               .Parameters(0).Value = RS!Investigation_Index
               Set RS2 = .Execute
            End With
            
            Do Until RS2.EOF
               Set tRow = hDoc.createElement("tr")
               Set tCell = hDoc.createElement("td")
               Set hAttGrp = tCell.Attributes
               Set hAtt = hDoc.createAttribute("colspan")
               hAtt.nodeValue = "4"
               hAttGrp.setNamedItem hAtt
               
               Set hAtt = hDoc.createAttribute("class")
               hAtt.nodeValue = "comment"
               hAttGrp.setNamedItem hAtt
               
               Set hTxt = hDoc.createTextNode(RS2!Service_Investigation_Comment)
               tCell.appendChild hTxt
               tRow.appendChild tCell
               hTab.appendChild tRow
               RS2.MoveNext
            Loop
            
            RS2.Close
            Set RS2 = Nothing
         End If
         
         With iceCmd
            .CommandText = "ICELABCOMM_Report_Results"
            .Parameters(0).Value = RS!Investigation_Index
            .Parameters.Append .CreateParameter("LTSIndex", adInteger, adParamInput, , repListRS!EDI_LTS_Index)
            Set RS2 = .Execute
            .Parameters.Delete ("LTSIndex")
         End With
         
         Do Until RS2.EOF
            If RS2!EDI_OP_Suppress Then
               Set tRow = hDoc.createElement("tr")
               Set hAttGrp = tRow.Attributes
               Set hAtt = hDoc.createAttribute("class")
               hAtt.nodeValue = "suppress"
               hAttGrp.setNamedItem hAtt
                  
               Set tCell = hDoc.createElement("td")
               Set hAttGrp = tCell.Attributes
               Set hAtt = hDoc.createAttribute("colspan")
               hAtt.nodeValue = "4"
               hAttGrp.setNamedItem hAtt
               
               Set hTxt = hDoc.createTextNode("Output suppressed for test " & RS2!Result_Rubric)
               tCell.appendChild hTxt
               tRow.appendChild tCell
               hTab.appendChild tRow
               
            Else
               Set tRow = hDoc.createElement("tr")
               Set hAttGrp = tRow.Attributes
               Set hAtt = hDoc.createAttribute("class")
               hAtt.nodeValue = "result"
               hAttGrp.setNamedItem hAtt
                  
               Set tCell = hDoc.createElement("td")
               Set hAttGrp = tCell.Attributes
               Set hAtt = hDoc.createAttribute("width")
               hAtt.nodeValue = "55%"
               hAttGrp.setNamedItem hAtt
               
               If RS2!Abnormal_Flag Then
                  Set hAtt = hDoc.createAttribute("class")
                  hAtt.nodeValue = "oor"
                  hAttGrp.setNamedItem hAtt
               End If
            
               If IsNull(RS2!Read_V2RC) Then
                  Set hTxt = hDoc.createTextNode(RS2!Result_Rubric)
               Else
                  Set hTxt = hDoc.createTextNode(RS2!Result_Rubric & " [" & RS2!Read_V2RC & "]")
               End If
               
               tCell.appendChild hTxt
               tRow.appendChild tCell
               
               Set tCell = hDoc.createElement("td")
               Set hAttGrp = tCell.Attributes
               Set hAtt = hDoc.createAttribute("width")
               hAtt.nodeValue = "15%"
               hAttGrp.setNamedItem hAtt
               
               
               If RS2!Abnormal_Flag Then
                  Set hAtt = hDoc.createAttribute("class")
                  hAtt.nodeValue = "oor"
                  hAttGrp.setNamedItem hAtt
               End If
            
               Set hAtt = hDoc.createAttribute("align")
               hAtt.nodeValue = "right"
               hAttGrp.setNamedItem hAtt
                  
               Set hTxt = hDoc.createTextNode(RS2!Result)
               tCell.appendChild hTxt
               tRow.appendChild tCell
               
               Set tCell = hDoc.createElement("td")
               Set hAttGrp = tCell.Attributes
               Set hAtt = hDoc.createAttribute("width")
               hAtt.nodeValue = "15%"
               hAttGrp.setNamedItem hAtt
               
               Set hTxt = hDoc.createTextNode(RS2!UOM_Code)
               tCell.appendChild hTxt
               tRow.appendChild tCell
               
               Set tCell = hDoc.createElement("td")
               Set hAttGrp = tCell.Attributes
               Set hAtt = hDoc.createAttribute("width")
               hAtt.nodeValue = "15%"
               hAttGrp.setNamedItem hAtt
               
               If (RS2!Lower_Range & "") = "" Then
                  If (RS2!Upper_Range & "") = "" Then
                     Set hTxt = hDoc.createTextNode("")
                  Else
                     Set hTxt = hDoc.createTextNode("< " & RS2!Upper_Range)
                  End If
               ElseIf (RS2!Upper_Range & "") = "" Then
                  If (RS2!Lower_Range & "") <> "" Then
                     Set hTxt = hDoc.createTextNode("> " & RS2!Lower_Range)
                  End If
               Else
                  Set hTxt = hDoc.createTextNode(RS2!Lower_Range & " to " & RS2!Upper_Range)
               End If
               
               tCell.appendChild hTxt
               tRow.appendChild tCell
               
               hTab.appendChild tRow
               
               If RS2!Comment_Marker Then
                  With iceCmd
                     .CommandText = "ICELABCOMM_Report_ResultComment"
                     .Parameters(0).Value = RS2!Result_Index
                     Set RS3 = .Execute
                  End With
                  
                  Do Until RS3.EOF
                     Set tRow = hDoc.createElement("tr")
                     
                     Set tCell = hDoc.createElement("td")
                     Set hAttGrp = tCell.Attributes
                     Set hAtt = hDoc.createAttribute("colspan")
                     hAtt.nodeValue = "4"
                     hAttGrp.setNamedItem hAtt
                  
                     Set hAtt = hDoc.createAttribute("class")
                     hAtt.nodeValue = "comment"
                     hAttGrp.setNamedItem hAtt
                     
                     Set hTxt = hDoc.createTextNode(RS3!Service_Result_Comment)
                     tCell.appendChild hTxt
                     tRow.appendChild tCell
            
                     hTab.appendChild tRow
                     
                     RS3.MoveNext
                  Loop
                  
                  RS3.Close
               End If
            End If
            
            RS2.MoveNext
         Loop
         
         RS2.Close
      End If
      
      RS.MoveNext
   Loop
   
   hDiv.appendChild hTab
   
   Set ReportInvest = hDiv
   
   RS.Close
   
   Set RS3 = Nothing
   Set RS2 = Nothing
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "HTMLClass.ReportInvest"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function ReportSamples(RepId As Long) As MSHTML.IHTMLDOMNode
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   Dim hDiv As MSHTML.IHTMLDOMNode
   Dim oDiv As MSHTML.IHTMLDOMNode
   
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Sample"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   Set oDiv = hDoc.createElement("div")
   Set hAttGrp = oDiv.Attributes
   Set hAtt = hDoc.createAttribute("class")
   hAtt.nodeValue = "report"
   hAttGrp.setNamedItem hAtt
   
   Do Until RS.EOF
      Set hDiv = hDoc.createElement("div")
      Set hAttGrp = hDiv.Attributes
      Set hAtt = hDoc.createAttribute("class")
      hAtt.nodeValue = "samp"
      hAttGrp.setNamedItem hAtt
   
      Set hTab = hDoc.createElement("table")
      Set hAttGrp = hTab.Attributes
      Set hAtt = hDoc.createAttribute("width")
      hAtt.nodeValue = "100%"
      hAttGrp.setNamedItem hAtt
   
      Set tRow = hDoc.createElement("tr")
      
      Set tCell = hDoc.createElement("td")
      Set hAttGrp = tCell.Attributes
      Set hAtt = hDoc.createAttribute("class")
      hAtt.nodeValue = "label"
      hAttGrp.setNamedItem hAtt
      Set hTxt = hDoc.createTextNode("Sample:")
      tCell.appendChild hTxt
      tRow.appendChild tCell
      
      Set tCell = hDoc.createElement("td")
      Set hTxt = hDoc.createTextNode(RS!Sample_Id & " (" & RS!Sample_Text & ")")
      tCell.appendChild hTxt
      tRow.appendChild tCell
      
      Set tCell = hDoc.createElement("td")
      Set hAttGrp = tCell.Attributes
      Set hAtt = hDoc.createAttribute("class")
      hAtt.nodeValue = "label"
      hAttGrp.setNamedItem hAtt
      Set hTxt = hDoc.createTextNode("Collected:")
      tCell.appendChild hTxt
      tRow.appendChild tCell
      
      Set tCell = hDoc.createElement("td")
      Set hTxt = hDoc.createTextNode(RS!Collection_DateTime)
      tCell.appendChild hTxt
      tRow.appendChild tCell
      
      Set tCell = hDoc.createElement("td")
      Set hAttGrp = tCell.Attributes
      Set hAtt = hDoc.createAttribute("class")
      hAtt.nodeValue = "label"
      hAttGrp.setNamedItem hAtt
      Set hTxt = hDoc.createTextNode("Received:")
      tCell.appendChild hTxt
      tRow.appendChild tCell
      
      Set tCell = hDoc.createElement("td")
      Set hTxt = hDoc.createTextNode(RS!Collection_DateTimeReceived)
      tCell.appendChild hTxt
      tRow.appendChild tCell
      hTab.appendChild tRow
      
'        Report level comments
      With iceCmd
         .CommandText = "ICELABCOMM_Report_Comments"
         Set RS2 = .Execute
      End With
      
      RS2.Filter = "Comment_Type<>'D'"
      Do Until RS2.EOF
         Set tRow = hDoc.createElement("tr")
         Set tCell = hDoc.createElement("td")
         Set hAttGrp = tCell.Attributes
         
         Set hAtt = hDoc.createAttribute("colspan")
         hAtt.nodeValue = "6"
         hAttGrp.setNamedItem hAtt
         
         Set hAtt = hDoc.createAttribute("class")
         hAtt.nodeValue = "comment"
         hAttGrp.setNamedItem hAtt
         
         Set hTxt = hDoc.createTextNode(RS2!Service_Report_Comment)
         tCell.appendChild hTxt
         
         tRow.appendChild tCell
         hTab.appendChild tRow
         
         RS2.MoveNext
      Loop
      RS2.Close
      
      hTab.appendChild tRow
      hDiv.appendChild hTab
      oDiv.appendChild hDiv
      
      oDiv.appendChild ReportInvest(RS!Sample_Index)
      
      RS.MoveNext
   Loop
   
   Set ReportSamples = oDiv
   RS.Close
   
   Set RS3 = Nothing
   Set RS2 = Nothing
   Set RS = Nothing
   Set iceCmd = Nothing
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "clsHTML.ReportSamples"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Sub SendHTMLCopy(ReportIndex As Long, _
                        Mailaddress As String)
   On Error GoTo procEH
   Dim strHTML As String
   Dim objApp As New Outlook.Application
   Dim objNameSpace As Outlook.NameSpace
   Dim objMailItem As Outlook.MailItem
   
   ReportId = ReportIndex
   StandaloneReport ReportIndex
'  Send using Outlook 2000
   eClass.FurtherInfo = "Preparing to send via Outlook"
   
   Set objNameSpace = objApp.GetNamespace("MAPI")
   Set objMailItem = objApp.CreateItem(olMailItem)
   objMailItem.Recipients.Add Mailaddress
   
   objMailItem.Subject = "Copy Report - Requesting Clinician: " & Clinician & " (" & _
                         msgControl.CurrentPractice & "/" & repListRS!EDI_Loc_Nat_Code_To & _
                         " Trust: " & msgControl.Organisation & "/" & msgControl.Trust & ")"
   
   msgControl.LogReportMessage -1, "HTML copy of report " & LabNo & " for " & pForename & " " & pSurname & _
                                   " sent to: " & Mailaddress
   objMailItem.htmlBody = hDoc.documentElement.outerHTML
   objMailItem.Send
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "clsHTML.SendHTMLCopy"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub StandaloneReport(RepIndex As Long)
   On Error GoTo procEH
   Dim ss As MSHTML.IHTMLStyleSheet
   
   Set hDoc = New MSHTML.HTMLDocument
   Set ss = hDoc.createStyleSheet '("table{FONT-FAMILY:Tahoma;FONT-SIZE:10px;}")
   ss.cssText = "table{FONT-FAMILY:Tahoma;FONT-SIZE:10px;}" & _
                ".label{FONT-WEIGHT:Bold;}" & _
                ".oor{color:red;}" & _
                "tr.result{background-color:gainsboro;}" & _
                "div.samp{left : 20px;position : relative;}" & _
                "div.inv{left : 40px;position : relative;}" & _
                ".comment{background-color: #FFFACD;text-indent: 10pt;}" & _
                ".suppress {color: blue;font-weight: bold;}"

   Set hBody = hDoc.Body
   hBody.removeChild hBody.firstChild
   
   hBody.appendChild ReportHeader
   Set hEL = hDoc.createElement("hr")
   Set hAttGrp = hEL.Attributes
   Set hAtt = hDoc.createAttribute("width")
   hAtt.nodeValue = "80%"
   hAttGrp.setNamedItem hAtt
   
   hBody.appendChild hEL
   hBody.appendChild ReportSamples(RepIndex)
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "clsHTML.StandaloneReport"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

