VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngReportID As Long
Private iceCmd As New ADODB.Command
Private patRS As New ADODB.Recordset
Private clinRs As New ADODB.Recordset
Private dtReported As String
Private dtCollected As String
Private dtReceived As String

Public Property Get Patient() As String
   Dim objBuff As New StringBuffer
   
   With objBuff
      .Append "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=2>"
      .Append "<TR><TD width=""18%""><NOBR><STRONG>Patient Name:</STRONG></NOBR></TD><TD width=""33%""><NOBR>"
      .Append patRS!Forename
      .Append " "
      .Append patRS!Surname
      .Append "</NOBR></TD><TD width=""22%""><STRONG><NOBR>NHS Number:</NOBR></STRONG></TD><TD width=""27%""><NOBR>"
      .Append patRS!NEW_NHS_No
      .Append "</NOBR></TD></TR><TR><TD><STRONG><NOBR>Date of Birth:</NOBR></STRONG></TD><TD><NOBR>"
      .Append patRS!Date_Of_Birth
      .Append "</NOBR></TD><TD><STRONG><NOBR>Hospital Number:</NOBR></STRONG></TD><TD><NOBR>"
      .Append patRS!Hospital_Number
      .Append "</NOBR></TD></TR><TR><TD><STRONG>Address:</STRONG></TD><TD colspan=3><NOBR>"
      .Append patRS!Pat_Addr_Line1
      If patRS!Pat_Addr_Line2 <> "" Then
          .Append ", "
          .Append patRS!Pat_Addr_Line2
      End If
      If patRS!Pat_Addr_Line3 <> "" Then
          .Append ", "
          .Append patRS!Pat_Addr_Line3
      End If
      If patRS!Pat_Addr_Line4 <> "" Then
          .Append ", "
          .Append patRS!Pat_Addr_Line4
      End If
      If patRS!Pat_Addr_Line5 <> "" Then
          .Append ", "
          .Append patRS!Pat_Addr_Line5
      End If
      If patRS!Pat_PostCode <> "" Then
          .Append ", "
          .Append patRS!Pat_PostCode
      End If
      .Append "</NOBR></TD></TR></TABLE>"
      
      Patient = .Value
   End With
End Property

Public Property Get ReportId() As String
    ReportId = CStr(mlngReportID)
End Property

Public Property Get LabNo() As String
   LabNo = patRS!Service_Report_Id
End Property

Public Property Get Specialty() As String
   Specialty = patRS!Service_Report_Type
End Property

Public Property Get Clinician() As String
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = ICEcon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_HealthParties"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set clinRs = .Execute
   End With
   clinRs.Filter = "EDI_HP_Type = '902'"
   Clinician = clinRs!Clinician_Surname
   clinRs.Close
   Set clinRs = Nothing
   Set iceCmd = Nothing
End Property

Public Property Get Destination() As String
    Destination = mobjReport.Destination
End Property
Public Property Get Comments() As String
   Dim RS As New ADODB.Recordset
   Dim objBuff As StringBuffer
    
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = ICEcon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Comments"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      .Parameters.Append .CreateParameter("Type", adVarChar, adParamInput, 1, "D")
      Set RS = .Execute
   End With
   
   If RS.RecordCount > 0 Then
        Set objBuff = New StringBuffer
        objBuff.Append "<BR><STRONG>Report Comments:</STRONG><BR>"
        Do Until RS.EOF
            objBuff.Append RS!Service_Report_Comment
            objBuff.Append "<BR>"
        Next
        Comments = objBuff.Value
        Set objBuff = Nothing
    End If
    RS.Close
    Set RS = Nothing
    Set iceCmd = Nothing
End Property

Public Property Get Investigations() As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim RS3 As New ADODB.Recordset
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim m As Integer
   Dim intResultCount As Integer
   Dim objBuff As StringBuffer
   Dim blnFirst As Boolean
   
   Set iceCmd = New ADODB.Command
   With iceCmd
      .ActiveConnection = ICEcon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Sample"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , ReportId)
      Set RS = .Execute
   End With
   
   Collected = RS!Collection_DateTime
   Reported = RS!DateTime_Of_Report
   
   With iceCmd
      .CommandText = "ICELABCOMM_Report_Investigation"
      .Parameters.Append .CreateParameter("SampId", adInteger, adParamInput, , RS!Sample_Index)
      .Parameters.Append .CreateParameter("Abnormal", adBoolean, adParamOutput)
      RS.Close
      Set RS = .Execute
      .Parameters.Delete (2)
      .Parameters.Delete (1)
   End With
   blnFirst = True
    
   Set objBuff = New StringBuffer
   objBuff.Append "<TABLE border=0 style=""WIDTH: 90%"" align=""center"" id=""tblResults"">"
   objBuff.Append "<TR>"
   objBuff.Append "<TD width=""51%""></TD>"
   objBuff.Append "<TD width=""4%""></TD>"
   objBuff.Append "<TD width=""15%"">"
   
   If blnFirst Then objBuff.Append "<STRONG>Result</STRONG>"
   objBuff.Append "</TD>"
   objBuff.Append "<TD width=""15%"">"
   If blnFirst Then objBuff.Append "<STRONG>Units</STRONG>"
   objBuff.Append "</TD>"
   objBuff.Append "<TD width=""15%"">"
   If blnFirst Then objBuff.Append "<STRONG>Range</STRONG>"
   objBuff.Append "</TD></TR>"
   blnFirst = False
                
   Do Until RS.EOF
      objBuff.Append "<TR><TD><STRONG>"
      objBuff.Append Trim(RS!Investigation_Requested)
      objBuff.Append "</STRONG></TD></TR>"
      Set iceCmd = Nothing
      
      If RS!Comment_Marker Then
         Set iceCmd = New ADODB.Command
         With iceCmd
            .ActiveConnection = ICEcon
            .CommandType = adCmdStoredProc
            .CommandText = "ICELABCOMM_Report_InvestComments"
            .Parameters.Append .CreateParameter("InvId", adInteger, adParamInput, , invRS!Investigation_Index)
            Set RS2 = .Execute
         End With
         Do Until RS2.EOF
            objBuff.Append "<TR><TD Colspan=5 STYLE=""padding-left:10px"">"
            objBuff.Append RTrim(RS2!Service_Investigation_Comment)
            objBuff.Append "</TD></TR>"
         Loop
         RS2.Close
         Set RS2 = Nothing
         Set iceCmd = Nothing
      End If
      
      With iceCmd
         .ActiveConnection = ICEcon
         .CommandType = adCmdStoredProc
         .CommandText = "ICELABCOMM_Report_Results"
         .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , RS!Investigation_Index)
         Set RS2 = .Execute
      End With
      
      Do Until RS2.EOF
         objBuff.Append "<TR style=""background-color:gainsboro""><TD STYLE=""padding-left:10px"">"
         objBuff.Append RS2!Result_Rubric
         objBuff.Append "</TD><TD align=""center"">"
         If RS2!Abnormal_Flag Then
             objBuff.Append "<FONT COLOR='RED'>*</FONT>"
         End If
         
         objBuff.Append "</TD><TD>"
         
         If RS2!Abnormal_Flag Then
             objBuff.Append "<FONT COLOR='RED'>"
             objBuff.Append RS2!Result
             objBuff.Append "</FONT>"
         Else
             objBuff.Append RS2!Result
         End If
         
         objBuff.Append "</TD><TD>"
         objBuff.Append RS2!UOM_Code
         objBuff.Append "</TD><TD>"
         objBuff.Append RS2!Lower_Range & " to " & RS2!Upper_Range
         objBuff.Append "</TD></TR>"
          ' Get comments
         Set iceCmd = Nothing
         If RS2!Comment_Marker Then
            With iceCmd
               .CommandText = "ICELABCOMM_Report_ResultComment"
               .Parameters(0).Value = RS2!Result_Index
               Set RS3 = .Execute
            End With
            
            Do Until RS3.EOF
               objBuff.Append "<TR><TD Colspan=5 STYLE=""Font-Family:Courier"">"
               objBuff.Append RS3!Service_Result_Comment
               objBuff.Append "</TD></TR>"
               RS3.MoveNext
            Loop
            RS3.Close
            Set RS3 = Nothing
         End If
         RS2.MoveNext
      Loop
      
      RS2.Close
      Set RS2 = Nothing
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
   Set iceCmd = Nothing
End Property

Public Property Get Collected() As String
   Collected = dtCollected
End Property

Public Property Get Received() As String
    If mobjReport.Samples(0).CollectionDateTime_Received > CDate("01/01/1900") Then
        Received = Format$(mobjReport.Samples(0).CollectionDateTime_Received, "DD/MM/YYYY HH:MM")
    Else
        Received = ""
    End If
End Property

Public Property Get Reported() As String
   Reported = dtReported
End Property

Public Function GetByPatient(ByVal ReportId As Long) As String
   Dim iceCmd As New ADODB.Command
   
   mlngReportID = ReportId
   
   With iceCmd
      .ActiveConnection = ICEcon
      .CommandType = adCmdStoredProc
      .CommandText = "ICELABCOMM_Report_Patient"
      .Parameters.Append .CreateParameter("RepId", adInteger, adParamInput, , mlngReportID)
      Set patRS = .Execute
   End With
End Function

Public Function GetPreviousReport(ByVal ReportId As Long) As Boolean
    Dim lngPrevRepID As Long
    lngPrevRepID = mobjReports.GetPrevReportID(ReportId)
    If lngPrevRepID = 0 Then
        'assume for time being lngPrevRepID = 0 when there is no prev record
        GetPreviousReport = False
    Else
       GetByPatient lngPrevRepID
       GetPreviousReport = True
    End If
End Function

Public Function GetNextReport(ByVal ReportId As Long) As Boolean
    Dim lngNextRepID As Long
    lngNextRepID = mobjReports.GetNextReportID(ReportId)
        If lngNextRepID = 0 Then
        'assume for time being lngPrevRepID = 0 when there is no next record
        GetNextReport = False
    Else
       GetByPatient lngNextRepID
       GetNextReport = True
    End If
End Function


Private Sub Class_Initialize()
'    Set mobjReports = New AHSLReporting.Reports
   Set mobjReports = ahreps
End Sub
