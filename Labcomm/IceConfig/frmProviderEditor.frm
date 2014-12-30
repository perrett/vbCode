VERSION 5.00
Object = "{3F118CA4-97A8-4926-AA0A-6FD80DFB3DCE}#2.6#0"; "PrpList2.ocx"
Begin VB.Form frmProviderEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provider Editor"
   ClientHeight    =   5685
   ClientLeft      =   1620
   ClientTop       =   3870
   ClientWidth     =   4725
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   2175
   End
   Begin PropertiesListCtl.PropertiesList PL1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9128
      LicenceData     =   "003E5E20294324255423005823381113235F243E5E3C6C6131395D701F41313E5A353F"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxDropItems    =   0
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmProviderEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blnReload As Boolean

Private Sub Command1_Click()
    On Error GoTo procEH
    Dim SPID As Integer
    If Label1.Caption <> "" Then
        SPID = Val(Label1.Caption)
        TempStr = "Update Service_Providers Set "
        TempStr = TempStr + "Provider_Code='" & PL1.PropertyItems("PROV_CODE").value & "',"
        TempStr = TempStr + "Provider_Nat_Code='" & PL1.PropertyItems("NAT_CODE").value & "',"
        TempStr = TempStr + "Provider_Name='" & PL1.PropertyItems("NAME").value & "',"
        TempStr = TempStr + "Discipline_Index =  " & PL1("DISC").value & ", "
'        TempStr = TempStr + "Specialty='" & PL1("SPEC").value & "',"
        TempStr = TempStr + "Request_Header_1='" & PL1.PropertyItems("HEADER_1").value & "',"
        TempStr = TempStr + "Request_Header_2='" & PL1.PropertyItems("HEADER_2").value & "',"
        TempStr = TempStr + "Request_Copies=" & Format(PL1.PropertyItems("COPIES").value) & ","
        TempStr = TempStr + "Requests_To_Barcode=" & Abs(CInt(PL1.PropertyItems("BARCODE").value)) & ","
        TempStr = TempStr + "Requests_To_Outtray=" & Abs(CInt(PL1.PropertyItems("OUTTRAY").value)) & ","
        TempStr = TempStr + "Sample_Panel_Id='" & PL1("SPAN").value & "',"
        TempStr = TempStr + "HDR_Request_Logo='" & PL1.PropertyItems("HEADER").value & "',"
        TempStr = TempStr + "HDR_Request_Logo_Text='" & PL1.PropertyItems("HEADER_CAPTION").value & "',"
        TempStr = TempStr + "FTR_Request_Logo='" & PL1.PropertyItems("FOOTER").value & "',"
        TempStr = TempStr + "FTR_Request_Logo_Text='" & PL1.PropertyItems("FOOTER_CAPTION").value & "',"
        TempStr = TempStr + "Orientation='" & PL1.PropertyItems("ORIENTATION").ListItems(PL1.PropertyItems("ORIENTATION").value).Name & "',"
        If PL1.PropertyItems("LBL_FORMAT").value > 0 Then
            TempStr = TempStr + "Label_Format='" & PL1.PropertyItems("LBL_FORMAT").ListItems(PL1.PropertyItems("LBL_FORMAT").value).Name & "',"
        Else
            TempStr = TempStr + "Label_Format='', "
        End If
        If PL1.PropertyItems("LBL_LABNOS").value > 0 Then
           TempStr = TempStr + "Label_LabNos='" & PL1.PropertyItems("LBL_LABNOS").ListItems(PL1.PropertyItems("LBL_LABNOS").value).Name & "' "
        Else
           TempStr = TempStr + "Label_LabNos='' "
        End If
        TempStr = TempStr + "Where Provider_ID=" & Label1.Caption
        iceCon.Execute TempStr
    Else
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open "Select Max(Provider_ID) 'Prov_ID' From Service_Providers", iceCon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount = 1 Then
            SPID = RS!Prov_ID + 1
        Else
            SPID = 1
        End If
        RS.Close
        TempStr = "Insert Into Service_Providers (Provider_ID,Provider_Code,Provider_Nat_Code,Provider_Name,Discipline_Index,Request_Header_1,Request_Header_2,Request_Copies,"
        TempStr = TempStr + "Requests_To_Barcode,Requests_To_Outtray,HDR_Request_Logo,HDR_Request_Logo_Text,FTR_Request_Logo,FTR_Request_Logo_Text,Orientation,Label_Format,Label_LabNos,Sample_Panel_ID) Values ("
        TempStr = TempStr + Format(SPID) + ","
        TempStr = TempStr + "'" & PL1.PropertyItems("PROV_CODE").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("NAT_CODE").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("NAME").value & "',"
        TempStr = TempStr & PL1("DISC").value & ", "
'        TempStr = TempStr + "'" & PL1("SPEC").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("HEADER_1").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("HEADER_2").value & "',"
        TempStr = TempStr + Format(PL1.PropertyItems("COPIES").value) & ","
        TempStr = TempStr + Format(Abs(CInt(PL1.PropertyItems("BARCODE").value))) & ","
        TempStr = TempStr + Format(Abs(CInt(PL1.PropertyItems("OUTTRAY").value))) & ","
        TempStr = TempStr + "'" & PL1.PropertyItems("HEADER").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("HEADER_CAPTION").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("FOOTER").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("FOOTER_CAPTION").value & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("ORIENTATION").ListItems(PL1.PropertyItems("ORIENTATION").value).Name & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("LBL_FORMAT").ListItems(PL1.PropertyItems("LBL_FORMAT").value).Name & "',"
        TempStr = TempStr + "'" & PL1.PropertyItems("LBL_LABNOS").ListItems(PL1.PropertyItems("LBL_LABNOS").value).Name & "',"
        TempStr = TempStr + CStr(PL1("SPAN").value) & ")"
        Debug.Print TempStr
        iceCon.Execute TempStr
    End If
    If blnReload Then
      If Label1.Caption <> "" Then
         TempStr = "UPDATE Request_Tests SET " & _
                     "Discipline_Index = " & IIf(Val(PL1("DISC").value) = 0, "Null", PL1("DISC").value) & _
                  " WHERE Provider_ID = " & Label1.Caption
         iceCon.Execute TempStr
      End If
      frmMain.edipr("TESTSPEC").value = IIf(Val(PL1("DISC").value) = 0, "", PL1("DISC").value)
      frmMain.SSListBar1_ListItemClick frmMain.SSListBar1.ListItems("Request Details")
    End If
    Unload Me
    Exit Sub
    
procEH:
    eClass.CurrentProcedure = "frmProviderEditor.Command1_Click"
    eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
 
Private Sub Form_Activate()
    If Label1.Caption <> "" Then
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open "Select * From Service_Providers Where Provider_ID=" & Label1.Caption, iceCon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount = 1 Then
            PL1.PropertyItems("PROV_CODE").value = Trim(RS!Provider_Code) & ""
            PL1.PropertyItems("NAT_CODE").value = Trim(RS!Provider_Nat_Code) & ""
            PL1.PropertyItems("NAME").value = Trim(RS!Provider_Name) & ""
            PL1("SPAN").value = Val(Trim(RS!Sample_Panel_Id & ""))
'            PL1("SPEC").value = RS!Specialty
            PL1("DISC").value = Val(Trim(RS!Discipline_Index & ""))
'            PL1("SPEC").value = IIf(IsNull(RS!Specialty) Or RS!Specialty = "", PL1("SPEC").ListItems(1).value, Trim(RS!Specialty))
            PL1.PropertyItems("HEADER_1").value = Trim(RS!Request_Header_1) & ""
            PL1.PropertyItems("HEADER_2").value = Trim(RS!Request_Header_2) & ""
            PL1.PropertyItems("COPIES").value = RS!Request_Copies
            PL1.PropertyItems("BARCODE").value = RS!Requests_To_Barcode
            PL1.PropertyItems("OUTTRAY").value = RS!Requests_To_Outtray
            PL1.PropertyItems("HEADER").value = Trim(RS!HDR_Request_Logo) & ""
            PL1.PropertyItems("HEADER_CAPTION").value = Trim(RS!HDR_Request_Logo_Text) & ""
            PL1.PropertyItems("FOOTER").value = Trim(RS!FTR_Request_Logo) & ""
            PL1.PropertyItems("FOOTER_CAPTION").value = Trim(RS!FTR_Request_Logo_Text) & ""
            PL1.PropertyItems("ORIENTATION").value = PL1.PropertyItems("ORIENTATION").ListItems.NameToIndex(Trim(RS!Orientation) & "")
            PL1.PropertyItems("LBL_FORMAT").value = PL1.PropertyItems("LBL_FORMAT").ListItems.NameToIndex(Trim(RS!Label_Format) & "")
            PL1.PropertyItems("LBL_LABNOS").value = PL1.PropertyItems("LBL_LABNOS").ListItems.NameToIndex(Trim(RS!Label_LabNos) & "")
        End If
        RS.Close
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        MsgBox "The character ' is not permitted, please use the ` character instead"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   PL1.PropertyItems.Add "PROV_CODE", "Provider Code", plpsString, , "Local code for this provider."
   PL1.PropertyItems("PROV_CODE").max = 8
   PL1.PropertyItems("PROV_CODE").value = "LocCode"
   PL1.PropertyItems.Add "NAT_CODE", "National Code", plpsString, frmMain.cboTrust.Text, "NHS National code for this provider"
   PL1.PropertyItems("NAT_CODE").max = 8
   PL1.PropertyItems.Add "NAME", "Name", plpsString, , "Name of the laboratory.  This will be printed at the top of a request form."
   PL1.PropertyItems("NAME").max = 65
   PL1.PropertyItems("NAME").value = "New Provider"
   
   PL1.PropertyItems.Add "DISC", "Discipline", plpsList, , "The discipline associated with this provider"
   strSQL = "SELECT * " & _
            "FROM Service_Discipline_Map " & _
            "WHERE Specialty_Code <> ''" & _
            "ORDER BY Specialty_Code"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockPessimistic
   With PL1("DISC").ListItems
      .Add "None", 0
      Do Until RS.EOF
         .Add RS!Discipline_Expansion, RS!Discipline_Index
         RS.MoveNext
      Loop
   End With
   PL1("DISC").value = 0
   RS.Close
   
   PL1.PropertyItems.Add "SPAN", "Sample Panel", plpsList, , "The sample panel to use with this provider"
   strSQL = "SELECT * " & _
            "FROM Request_Sample_Panels"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   With PL1("SPAN").ListItems
      .Add "Not yet allocated", 0
      Do Until RS.EOF
         .Add RS!Panel_NAme, RS!Sample_Panel_Id
         RS.MoveNext
      Loop
   End With
   PL1("SPAN").value = 0
   RS.Close
   Set RS = Nothing
   
   PL1.PropertyItems.Add "HEADER_1", "Header 1", plpsString, , "Header 1 for request form."
   PL1.PropertyItems("HEADER_1").max = 65
   PL1.PropertyItems.Add "HEADER_2", "Header 2", plpsString, , "Header 2 for request form.  Printed in a larger font beneath header 1."
   PL1.PropertyItems("HEADER_2").max = 65
   PL1.PropertyItems.Add "COPIES", "Copies to Print", plpsNumber, 1, "Copies of request printed on form."
   PL1.PropertyItems("COPIES").Min = 0
   PL1.PropertyItems("COPIES").max = 2
   PL1.PropertyItems("COPIES").Increment = 1
   PL1.PropertyItems.Add "BARCODE", "Send Requests to Barcode", plpsBoolean, False, "Should requests be encoded in PDF barcode for printing."
   PL1.PropertyItems.Add "OUTTRAY", "Send Requests Electronically", plpsBoolean, False, "Should requests be sent electronically"
   PL1.PropertyItems.Add "HEADER", "Path to Header Logo", plpsFile, , "Path to bitmap used as header logo on printed requests."
   PL1.PropertyItems("HEADER").max = 140
   PL1.PropertyItems("HEADER").DialogTitle = "Select image"
   PL1.PropertyItems("HEADER").Filter = "Graphics Files|*.BMP;*.WMF;*.JPG|All Files|*.*"
   PL1.PropertyItems("HEADER").FilterIndex = 0
   PL1.PropertyItems.Add "HEADER_CAPTION", "Header Caption", plpsString, , "Caption to be printed on request form header."
   PL1.PropertyItems("HEADER_CAPTION").max = 45
   PL1.PropertyItems.Add "FOOTER", "Path to Footer Logo", plpsFile, , "Path to bitmap used as footer logo on printed requests."
   PL1.PropertyItems("FOOTER").max = 140
   PL1.PropertyItems("FOOTER").DialogTitle = "Select image"
   PL1.PropertyItems("FOOTER").Filter = "Graphics Files|*.BMP;*.WMF;*.JPG|All Files|*.*"
   PL1.PropertyItems("FOOTER").FilterIndex = 0
   PL1.PropertyItems.Add "FOOTER_CAPTION", "Footer Caption", plpsString, , "Caption to be printed on request form footer."
   PL1.PropertyItems("FOOTER_CAPTION").max = 45
   PL1.PropertyItems.Add "ORIENTATION", "Request Orientation", plpsList, , "Orientation in which requests are printed."
   PL1.PropertyItems("ORIENTATION").ListItems.Add "POR", 1
   PL1.PropertyItems("ORIENTATION").ListItems.Add "LAN", 2
   PL1.PropertyItems("ORIENTATION").value = 1
   PL1.PropertyItems.Add "LBL_FORMAT", "Label Format", plpsList, , "If using on-ward label printers, style of label to print."
   With PL1("LBL_FORMAT").ListItems
      .Add "A4", 1
      .Add "ELTRON_GEN", 2
      .Add "ELTRON_PHL", 3
      .Add "ELTRON_BT", 4
   End With
   PL1.PropertyItems("LBL_FORMAT").value = 1
   PL1.PropertyItems.Add "LBL_LABNOS", "Lab No. Format", plpsList, , "If using on-ward label printers, style of laboratory number barcode."
   With PL1("LBL_LABNOS").ListItems
      .Add "CODABAR", 1
      .Add "CODE 39", 2
      .Add "EAN13", 3
      .Add "EAN-8", 4
      .Add "EAN13+2", 5
      .Add "EAN13+5", 6
      .Add "UPC-A", 7
      .Add "UPC-E", 8
      .Add "ITF-14", 9
      .Add "ITF-6", 10
      .Add "CODE 128", 11
      .Add "EAN-128", 12
      .Add "2 OF 5", 13
      .Add "1-2 OF 5", 14
      .Add "3 0F 9", 15
      .Add "CODE B", 16
      .Add "CODE 11", 17
   End With
   PL1.PropertyItems("LBL_LABNOS").value = 1
   blnReload = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmProvider.CloseMe = blnReload
End Sub

Private Sub PL1_AfterEdit(PropertyItem As PropertiesListCtl.PropertyItem, newValue As Variant, Cancel As Boolean)
   Dim strSQL As String
   Dim testsAffected As String
   Dim newDisc As String
   Dim mbPrompt As String
   Dim RS As New ADODB.Recordset
   
   If PropertyItem.Key = "DISC" Then
      blnReload = True
      strSQL = "SELECT * " & _
               "FROM Service_Discipline_Map " & _
               "WHERE Discipline_Index = " & Val(newValue)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      newDisc = RS!Specialty_Code
      
      RS.Close
      
'      If newValue = 0 Then
'         newDisc = PL1("SPEC").value
'         PL1("SPEC").value = ""
'      Else
'         strSQL = "SELECT * " & _
'                  "FROM Service_Discipline_Map " & _
'                  "WHERE Discipline_Index = " & Val(newValue)
'         RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'         newDisc = RS!Specialty_Code
'   '      PL1("SPEC").value = RS!Specialty_Code
'         RS.Close
'      End If
      
      strSQL = "SELECT * " & _
               "FROM Request_Tests " & _
               "WHERE Provider_Id = '" & Label1.Caption & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      If RS.RecordCount > 0 Then
         Do Until RS.EOF
            testAffected = testAffected & RS!Screen_Caption & " (" & RS!Test_Code & ")" & vbCrLf
            RS.MoveNext
         Loop
         
         mbPrompt = "ALL the following tests will have their discipline " & _
                    IIf(newValue = 0, "REMOVED. ", "Amended to " & newDisc) & _
                    vbCrLf & vbCrLf & testAffected & vbCrLf
         If newValue = 0 Then
            mbPrompt = mbPrompt & "All the tests will need to be manually edited to allocate a discipline - Continue?"
         Else
            mbPrompt = mbPrompt & "Continue?"
         End If
         
         If MsgBox(mbPrompt, vbYesNo, "Are you sure?") = vbYes Then
            Cancel = False
'            If newValue = 0 Then
'               PL1("SPEC").Value = ""
'            Else
'               PL1("SPEC").Value = newDisc
'            End If
         Else
            Cancel = True
         End If
      End If
      
      RS.Close
   End If
   Set RS = Nothing
End Sub

