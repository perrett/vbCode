VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadEDIRecipients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private plist As PropertiesList
Private newNode As Node
Private strSQL As String
Private vData As Variant
Private mCtrl As New ManageControls
Private nd(2) As Node

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.ediPr
   mCtrl.TreeViewUpdate = True
   
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .Caption = "EDI Recipients"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"

'     Add control items as required
      .PropertyItems.Add "NATIONAL", "", plpsString, , "Hidden"
      .PropertyItems.Add "REFID", , plpsNumber, , "Hidden"
      .PropertyItems.Add "SUBHEADER", , plpsString, , "Hidden"
      .PropertyItems.Add "SUBID", , plpsString, , "Hidden"
      .PropertyItems.Add "GPCODE", , plpsString, , "Hidden"
      .PropertyItems.Add "HOLDSTATE", , plpsBoolean, , "Hidden"

'     Add pages as required
      .Pages.Add "EDI", "EDI_Recipients"
      .Pages.Add "PMIP", "EDI_Recipient_Ref"
      .Pages.Add "SP", "EDI_Loc_Specialties"
      .Pages.Add "MS", "EDI_Msg_types"
      .Pages.Add "IN", "EDI_Recipient_Individuals"
      
'     Set up the required property items
      With .PropertyItems
         .Add "NC", "National Code", plpsString, , "Mandatory - The NATIONAL Code for this practice"
         .Add "EMIS", "EMIS System", plpsBoolean, , "Is this practice an EMIS system?"
         .Add "LC", "Local Code", plpsString, , "The LOCAL Code for this practice"
         .Add "TD", "EDI Trader details", plpsCustom, "Click '...' to view", "EDI Trader Account & interchange settings"
         '.Add "ER", "EDI Reference Number", plpsString, , "EDI Trader Account"
         '.Add "FP", "EDI Free Part", plpsString, , "System free part identifier"
         .Add "SYS", "GP System/Version", plpsString, , "The GP system (EMIS TOREX etc) + version"
         .Add "EN", "Encryption Status", plpsList, , "Encrypt and/or sign this message?"
         .Add "OG", "Report Grouping", plpsList, , "'M' for Message Type, 'S' = by Specialty, 'R' = Individual Reports"
         .Add "DR", "Delivery Report", plpsBoolean, , "Is a delivery report required for this message"
         .Add "SA", "Send Attachments", plpsBoolean, , "Send attachments"
         .Add "NA", "Name", plpsString, , "Practice Name"
         .Add "AD", "Address", plpsString, , "Practice Address"
         .Add "CNAME", "Contact name", plpsString, , "The Practice Contact"
         .Add "CTEL", "Contact Telephone", plpsString, , "The Contact's telephone No."
         .Add "CNOTE", "Contact notes", plpsString, , "Any relevant notes"
         .Add "DTS", "DTS Address", plpsString, , "The DTS mail address"
         '.Add "LNK", "Link_Interchange", plpsBoolean, True, "Maintain a single Interchange No. (False = Per message Type"
         
         strSQL = "SELECT EDI_LTS_Index, EDI_Msg_Type, EDI_OrgCode " & _
                  "FROM EDI_Local_Trader_Settings"
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         
         Do Until RS.EOF
            .Add "Spec_" & RS!EDI_LTS_Index, "Specialty (" & RS!EDI_OrgCode & " - " & RS!EDI_Msg_Type & ")", plpsString, RS!EDI_LTS_Index, "Click the '+' on the tree (left) to expand this group"
            With .item("Spec_" & RS!EDI_LTS_Index)
               .ReadOnly = True
               .PageKeys = "EDI"
               .DefaultExtension = "SubNodes"
               .Flags = 1
               .Tag = "MENU"
            End With
            
            RS.MoveNext
         Loop
         RS.Close
         
'         .Add "Spec", "Specialty (" & frmMain.CurrentLTSOrg & " - " & frmMain.CurrentLTSDataStream & ")", plpsString, , "Click the '+' on the tree (left) to expand this group"
         .Add "Msg", "Message Types", plpsString, , "Click the '+' on the tree (left) to expand this group"
         .Add "Ind", "Individuals", plpsString, , "Click the '+' on the tree (left) to expand this group"
         .Add "SM", "SMTP Mail Address", plpsString, , "The internet mail address of this practice"
         .Add "HO", "Hold all output", plpsBoolean, False, "Process any incoming data, but hold the outgoing files."
         .Add "AC", "Output Active", plpsBoolean, , "Is this practice active?"
         
         .Add "SP+MS1", "Korner Code", plpsCustom, , "Click '...' to select"
         .Add "SP+MS2", "Specialty Description", plpsString, , "Dependant on Korner Code"
         .Add "SP+MS3", "Output format", plpsCustom, , "Click on '...' to select"
         .Add "SP+MS4", "Run Frequency", plpsString, "G", "'G' = Global Default, 'Nnn' = Every nn minutes, 'D' = Daily, 'W' = Weekly 'H' = Hold"
         .Add "SP+MS5", "Start Time 1", plpsString, , "The first start time"
         .Add "SP+MS6", "Start Time 2", plpsString, , "The second start time"
         .Add "SP+MS7", "Start Time 3", plpsString, , "The third start time"
         .Add "SP+MS8", "Start Time 4", plpsString, , "The fourth start time"
         .Add "SP+MS9", "Active", plpsBoolean, , "Is this specialty Active?"
      
         .Add "MS+MS1", "EDI Message Format", plpsList, , "In what format is the message to be sent?"
         .Add "MS+MS2", "Delivery Method", plpsList, , "X = X400"
         .Add "MS+MS3", "Encryption Enabled", plpsBoolean, , "Switch Encryption on?"
         .Add "MS+MS9", "Copy To", plpsCustom, , "Send a copy of these messages to this recipient"
         .Add "MS+MS4", "Acknowledgements Active", plpsBoolean, , "Are acknowledgements required?"
         .Add "MS+MS5", "Test Messages Only", plpsBoolean, , "Send only test messages using this format?"
         .Add "MS+MS6", "Anonymize Data", plpsBoolean, , "Send only test messages using this format?"
         .Add "MS+MS7", "Redirect To", plpsCustom, , "Send only test messages using this format?"
         .Add "MS+MS8", "Active", plpsBoolean, , "Is this message type active?"
      
         .Add "IN+IN1", "Name to Output", plpsCustom, , "Mandatory - Name as it will  appear on the report"
         .Add "IN+IN2", "National Code", plpsString, , "Mandatory - National Code for the GP"
         .Add "IN+IN3", "Matching Details", plpsCustom, "Click '...' to view", "The lab system matching details"
         .Add "IN+IN4", "SMTP Mail address", plpsString, , "Internet mail address for this recipient"
         .Add "IN+IN5", "Direct Mail Active", plpsBoolean, , "Send HTML Reports to this individual?"
         .Add "IN+IN6", "Active", plpsBoolean, , "Is this GP Active?"
      End With
   
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 5

'     Set other Control items to the page key
      .PropertyItems("NATIONAL").PageKeys = "Control"
      .PropertyItems("REFID").PageKeys = "Control"
      .PropertyItems("SUBHEADER").PageKeys = "Control"
      .PropertyItems("SUBID").PageKeys = "Control"
      .PropertyItems("GPCODE").PageKeys = "Control"
      .PropertyItems("HOLDSTATE").PageKeys = "Control"
      
'     Set up individual item
'      With .PropertyItems("<Key>")
'         .Tag = ""               '<Database Field> or "MENU"
'         .Visible = True / False 'False means no teree view or properties lst item
'         .DefaultExtension = ""  '<Procedure within class to run>
'         .Flags = 0              '-1/0/1/2 - for popup menu control
'         .PageKeys = ""          'Page this item resides on
'         .DefaultValue = ""
'         .OverlayIcon = 0        'Only used in rare circumstances
'      End With
      With .PropertyItems("NC")
         .Tag = "EDI_NatCode"
         .max = 6
      End With
      With .PropertyItems("EMIS")
         .Tag = "EDI_EMIS_System"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("LC")
         .Tag = "EDI_LocalCode"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("TD")
         .DefaultExtension = "Item"
         .PageKeys = "EDI"
      End With
'      With .PropertyItems("FP")
'         .Tag = "EDI_Free_Part"
'         .DefaultExtension = "Item"
'         .PageKeys = "PMIP"
'      End With
      With .PropertyItems("SYS")
         .Tag = "EDI_GP_System"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("EN")
         .Tag = "EDI_encryption"
         .defaultValue = 0
         With .ListItems
            .Add "Disabled", 0
            .Add "Sign", 1
            .Add "Encrypt", 2
            .Add "Sign & Encrypt", 3
         End With
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("OG")
         .Tag = "EDI_Output_GroupBy"
         With .ListItems
            .Add "Message Type", "M"
            .Add "Specialty", "S"
            .Add "Individual Report", "R"
         End With
         .DefaultExtension = "Item"
      End With
      
      With .PropertyItems("DR")
         .Tag = "EDI_Delivery_Report"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("SA")
         .Tag = "Send_Attachments"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("NA")
         .Tag = "EDI_Name"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("AD")
         .Tag = "EDI_Address"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("DTS")
         .Tag = "DTSAddress"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("CNAME")
         .Tag = "EDI_Contact_Name"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("CTEL")
         .Tag = "EDI_Contact_Tel"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("CNOTE")
         .Tag = "EDI_Contact_Notes"
         .DefaultExtension = "Item"
      End With
'      With .PropertyItems("LNK")
'         .Tag = "Link_Interchange_Nos"
'         .DefaultExtension = "Item"
'         .PageKeys = "PMIP"
'      End With
      
'      With .PropertyItems("Spec")
'         .ReadOnly = True
'         .PageKeys = "EDI"
'         .DefaultExtension = "SubNodes"
'         .Flags = 1
'         .Tag = "MENU"
'      End With
      With .PropertyItems("Msg")
         .ReadOnly = True
         .PageKeys = "EDI"
         .Flags = 1
         .Tag = "MENU"
         .DefaultExtension = "SubNodes"
      End With
      With .PropertyItems("Ind")
         .ReadOnly = True
         .PageKeys = "EDI"
         .Flags = 1
         .Tag = "MENU"
         .DefaultExtension = "SubNodes"
      End With
      With .PropertyItems("SM")
         .Tag = "EDI_SMTP_Mail"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("HO")
         .Tag = "EDI_Hold_Output"
         .DefaultExtension = "item"
      End With
      With .PropertyItems("AC")
         .Tag = "EDI_Active"
         .DefaultExtension = "Item"
      End With
      
      With .PropertyItems("SP+MS1")
         .Tag = "EDI_Korner_Code"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS2")
         .Tag = "EDI_Specialty"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS3")
         .Tag = "EDI_Msg_Format"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS4")
         .Tag = "EDI_Run_Frequency"
'         With .ListItems
'            .Add "Global Default", "G"
'            .Add "Every nn minutes", "N"
'            .Add "Daily", "D"
'            .Add "Weekly", "W"
'         End With
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS5")
         .Tag = "EDI_S_Time1"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS6")
         .Tag = "EDI_S_Time2"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS7")
         .Tag = "EDI_S_Time3"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS8")
         .Tag = "EDI_S_Time4"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("SP+MS9")
         .Tag = "EDI_Msg_Active"
         .DefaultExtension = "SubItem"
      End With
   
      With .PropertyItems("MS+MS1")
         .Tag = "EDI_Msg_Format"
         .DefaultExtension = "SubItem"
         .PageKeys = "MS"
         
         strSQL = "SELECT * " & _
                  "FROM EDI_Msg_Formats " & _
                  "WHERE Active = 1"
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         Do Until RS.EOF
            .ListItems.Add RS!Description, RS!Type & "," & RS!Version
            RS.MoveNext
         Loop
         RS.Close
         Set RS = Nothing
         
'         With .ListItems
'            .Add "Edifact v2", "MEDRPT,NHS002"
'            .Add "Edifact v3", "MEDRPT,NHS003"
'            .Add "Astm", "ASTM1238,ASTM4.1"
'            .Add "RSR", "NHSRSR,NHS002"
'            .Add "FH", "CYTFH,FH"
'            .Add "HTML Report", "HTML,ICE001"
'            .Add "XML File", "XML,ICE002"
'         End With
      End With
      With .PropertyItems("MS+MS2")
         .Tag = "EDI_Delivery_Method"
         .DefaultExtension = "SubItem"
         .defaultValue = "X"
         With .ListItems
            .Add "X400/DTS", "X"
            .Add "SMTP", "S"
         End With
      End With
      With .PropertyItems("MS+MS3")
         .Tag = "EDI_Encrypt_enabled"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("MS+MS9")
         .Tag = "EDI_Copy_To"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("MS+MS4")
         .Tag = "EDI_Acks_Active"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("MS+MS5")
         .Tag = "EDI_Msg_Test"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("MS+MS6")
         .Tag = "EDI_Anonymize_Data"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("MS+MS7")
         .Tag = "EDI_Redirect_To"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("MS+MS8")
         .Tag = "EDI_Msg_Active"
         .DefaultExtension = "SubItem"
         .PageKeys = "MS"
         .defaultValue = True
      End With
      
      With .PropertyItems("IN+IN1")
         .Tag = "EDI_GP_Name"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("IN+IN2")
         .Tag = "EDI_NatCode"
         .DefaultExtension = "SubItem"
         .PageKeys = "IN"
         .ReadOnly = True
      End With
      With .PropertyItems("IN+IN3")
         .DefaultExtension = "SubItem"
         .PageKeys = "IN"
      End With
      With .PropertyItems("IN+IN4")
         .Tag = "EDI_SMTP_Mail"
         .DefaultExtension = "SubItem"
         .PageKeys = "IN"
      End With
      With .PropertyItems("IN+IN5")
         .Tag = "EDI_SMTP_Active"
         .DefaultExtension = "SubItem"
      End With
      With .PropertyItems("IN+IN6")
         .Tag = "EDI_Active"
         .DefaultExtension = "SubItem"
         .defaultValue = True
         .PageKeys = "IN"
      End With
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
   plist.Pages("EDI").Caption = "ICEConfig_Amend_Recipients"
   plist.Pages("SP").Caption = "ICEconfig_Amend_Specialties"
   plist.Pages("MS").Caption = "ICEConfig_Amend_MsgTypes"
   plist.Pages("IN").Caption = "ICEConfig_Amend_Individuals"
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.Class_Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub FirstView(Optional NodeId As Node = Nothing)
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim sampTxt As String
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("New", _
                                              "NewNode", _
                                              "NewMapping"), _
                             "Add new recipient...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT EDI_NatCode, " & _
               "EDI_Name, " & _
               "Ref_Index, " & _
               "Case " & _
                  "When EDI_Hold_Output is Null Then 0 " & _
                  "Else EDI_Hold_Output " & _
               "End as EDI_Hold_Output, " & _
               "EDI_Active " & _
            "FROM EDI_Recipients " & _
            "WHERE Organisation LIKE '" & frmMain.cboTrust.Text & "' " & _
            "ORDER BY EDI_NatCode"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
'   plist("Spec").Caption = "Specialty (" & frmMain.CurrentLTSOrg & " - " & frmMain.CurrentLTSDataStream & ")"
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!EDI_NatCode, _
                                                 RS!Ref_Index & "", _
                                                 "Detail", _
                                                 , _
                                                 ms_DELETE), _
                                RS!EDI_NatCode & " - " + Trim(RS!EDI_Name), _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
      If RS!EDI_Active Then
         nd(0).ForeColor = BPBLUE
      Else
         nd(0).ForeColor = BPRED
      End If
      
      If RS!EDI_Hold_Output Then
         nd(0).Image = 3
         nd(0).SelectedImage = 3
      End If
      
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub DefaultNew(natId As String, LocalId As String)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim strSQL As String
   Dim msgTypes(2) As String
   Dim iceCmd As ADODB.Command
   Dim i As Integer
   Dim intRet As Integer
   Dim failInfo As String
   
   msgTypes(0) = "MEDRPT,NHS002"
   msgTypes(1) = "MEDRPT,NHS003"
   
   strSQL = "SELECT EDI_LTS_Index " & _
            "FROM EDI_Local_Trader_Settings"
   RS2.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Do Until RS2.EOF
      strSQL = "SELECT * " & _
               "FROM CRIR_Specialty " & _
               "WHERE Specialty_code Between '800' And '899' or Specialty_Code = '502'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      Do Until RS.EOF
         Set iceCmd = New ADODB.Command
         With iceCmd
            .ActiveConnection = iceCon
            .CommandText = "ICECONFIG_Amend_Specialties"    '  "ICECONFIG_Test"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
            .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , True)
            .Parameters.Append .CreateParameter("S_OldCode", adVarChar, adParamInput, 3, RS!Specialty_Code)
            .Parameters.Append .CreateParameter("S_Org", adVarChar, adParamInput, 10, Trim(Left(frmMain.OrgList.Text, 6)))
            .Parameters.Append .CreateParameter("S_NatCode", adVarChar, adParamInput, 10, natId)
            .Parameters.Append .CreateParameter("S_Korner", adVarChar, adParamInput, 3, RS!Specialty_Code)
            .Parameters.Append .CreateParameter("S_Specialty", adVarChar, adParamInput, 30, Left(RS!Specialty, 30))
            .Parameters.Append .CreateParameter("S_Format", adVarChar, adParamInput, 16, msgTypes(1))
            .Parameters.Append .CreateParameter("S_RunFreq", adVarChar, adParamInput, 4, "")
            .Parameters.Append .CreateParameter("S_Start1", adVarChar, adParamInput, 5, "")
            .Parameters.Append .CreateParameter("S_Start2", adVarChar, adParamInput, 5, "")
            .Parameters.Append .CreateParameter("S_Start3", adVarChar, adParamInput, 5, "")
            .Parameters.Append .CreateParameter("S_Start4", adVarChar, adParamInput, 5, "")
            .Parameters.Append .CreateParameter("S_Active", adBoolean, adParamInput, , False)
            .Parameters.Append .CreateParameter("S_LTSIndex", adInteger, adParamInput, , RS2!EDI_LTS_Index)
            .Execute
            intRet = .Parameters("Return")
            failInfo = "ICECONFIG_Amend_Specialties"
         End With
         Set iceCmd = Nothing
         RS.MoveNext
      Loop
      
      RS.Close
      RS2.MoveNext
   Loop
   RS2.Close
   
   For i = 0 To 1
      Set iceCmd = New ADODB.Command
      With iceCmd
         .ActiveConnection = iceCon
         .CommandText = "ICECONFIG_Amend_MsgTypes"
         .CommandType = adCmdStoredProc
         .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
         .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , True)
         .Parameters.Append .CreateParameter("OldFmt", adVarChar, adParamInput, 16, "")
         .Parameters.Append .CreateParameter("M_Org", adVarChar, adParamInput, 10, frmMain.cboTrust.Text)
         .Parameters.Append .CreateParameter("M_NatCode", adVarChar, adParamInput, 10, natId)
         .Parameters.Append .CreateParameter("M_MsgFormat", adVarChar, adParamInput, 16, msgTypes(i))
         .Parameters.Append .CreateParameter("M_Delivery", adVarChar, adParamInput, 1, "X")
         .Parameters.Append .CreateParameter("M_Copy", adVarChar, adParamInput, 10, "")
         .Parameters.Append .CreateParameter("M_Encrypt", adBoolean, adParamInput, , IIf(Right(msgTypes(i), 3) = "003", True, False))
         .Parameters.Append .CreateParameter("M_Acks", adBoolean, adParamInput, , False)
         .Parameters.Append .CreateParameter("M_Testing", adBoolean, adParamInput, , False)
         .Parameters.Append .CreateParameter("M_Anon", adBoolean, adParamInput, , False)
         .Parameters.Append .CreateParameter("M_Redirect", adVarChar, adParamInput, 10, "")
         .Parameters.Append .CreateParameter("M_Active", adBoolean, adParamInput, , False)
         .Execute
         intRet = .Parameters("Return")
      End With
      Set iceCmd = Nothing
   Next i
   
   If Trim(natId) <> "" Then
      strSQL = "SELECT * " & _
               "FROM National_GPs " & _
               "WHERE Practice_Code = '" & natId & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      Do Until RS.EOF
         Set iceCmd = New ADODB.Command
         With iceCmd
            .ActiveConnection = iceCon
            .CommandText = "ICECONFIG_Amend_Loc_Recp"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
            .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , True)
            .Parameters.Append .CreateParameter("HPUpdate", adBoolean, adParamInput, , False)
            .Parameters.Append .CreateParameter("I_Index", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("OldNatCode", adVarChar, adParamInput, 10, "")
            .Parameters.Append .CreateParameter("I_Org", adVarChar, adParamInput, 10, frmMain.cboTrust.Text)
            .Parameters.Append .CreateParameter("I_OrgNatCode", adVarChar, adParamInput, 10, natId)
            .Parameters.Append .CreateParameter("I_GpName", adVarChar, adParamInput, 30, RS!Clinician_Name)
            .Parameters.Append .CreateParameter("I_NatCode", adVarChar, adParamInput, 10, RS!Clinician_National_Code)
            .Parameters.Append .CreateParameter("I_SMTP", adVarChar, adParamInput, 250, "")
            .Parameters.Append .CreateParameter("I_HTML", adBoolean, adParamInput, , False)
            .Parameters.Append .CreateParameter("I_Active", adBoolean, adParamInput, , False)
            .Parameters.Append .CreateParameter("Failure", adVarChar, adParamOutput, 40)
            .Execute
            intRet = .Parameters("Return")
            failInfo = "ICECONFIG_Amend_Loc_Recp (" & .Parameters("Failure") & ")"
         End With
         
         Set iceCmd = Nothing
         RS.MoveNext
      Loop
      RS.Close
   End If
   
   Set RS = Nothing
   Set RS2 = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.DefaultNew"""
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
   Dim RS As New ADODB.Recordset
   Dim iceCmd As New ADODB.Command
   Dim natCode As String
   Dim refId As Long
   Dim blnAllowDelete As Boolean
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdText
   End With
   
   Set tNode = objTV.TopLevelNode(NodeId)
   natCode = objTV.NodeLevel(tNode)
'   natCode = objTView.NodeLevel(objTView.TopLevelNode(node).Key)
   
   blnAllowDelete = True
   If NodeId = tNode Then
      eClass.FurtherInfo = "Attempting to delete practice " & natCode
      
      '  Determine the Ref_Index. The Inner Join is used to determine if there are any Recipient_Ref records to delete.
      strSQL = "SELECT er.Ref_Index " & _
               "FROM EDI_Recipients er " & _
                  "INNER JOIN EDI_Recipient_Ref rr " & _
                  "ON er.Ref_Index = rr.Ref_Index " & _
               "WHERE EDI_NatCode = '" & natCode & "'"
               
      RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
      If RS.EOF = False Then
         refId = RS(0)
      Else
         refId = 0
      End If
      
      RS.Close
      
      With iceCmd
         .CommandText = "ICECONFIG_RemoveRecipient"
         .Parameters.Append .CreateParameter("D_NatCode", adVarChar, adParamInput, 10, natCode)
         .CommandType = adCmdStoredProc
      End With
   
      If refId > 0 Then
         strSQL = "SELECT EDI_NatCode " & _
                  "FROM EDI_Recipients " & _
                  "WHERE Ref_Index = " & refId
         RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
         
         If RS.RecordCount = 1 Then
            '  No other EDI_Recipients point to this record, so delete it - remember Phoenix sites!
            strSQL = "DELETE FROM EDI_Recipient_Ref " & _
                     "WHERE Ref_Index = " & refId
            
            iceCon.Execute strSQL
            
            strSQL = "DELETE FROM EDI_Interchange_No WHERE Ref_Index = " & refId
            
            iceCon.Execute strSQL
         End If
      End If
   
   Else
      Select Case Left(NodeId.Parent.Text, 9)
         Case "Specialty"
            eClass.FurtherInfo = "Attempting to delete Specialty " & plist("SP+MS1").value
            strSQL = "EDI_Loc_Specialties WHERE EDI_Nat_Code = '" & natCode & _
                     "' AND EDI_Korner_Code= '" & plist("SP+MS1").value & "' " & _
                     "AND EDI_LTS_Index = " & frmMain.CurrentLTSIndex
            
         Case "Message T"
            eClass.FurtherInfo = "Attempting to delete Message Type " & plist("MS+MS1").value
            
'           Do not allow deletion if interchange information present
            strSQL = "SELECT * " & _
                     "FROM EDI_Msg_Types " & _
                     "WHERE EDI_Org_NatCode = '" & natCode & "' " & _
                        "AND EDI_Msg_Format = '" & plist("MS+MS1").value & "' " & _
                        "AND EDI_Last_Interchange > 0"
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            blnAllowDelete = RS.EOF
            RS.Close
            
'           Do not allow deletion if if the message type is in use by a specialty
            strSQL = "SELECT * " & _
                     "FROM EDI_Loc_Specialties " & _
                     "WHERE EDI_Msg_Format = '" & plist("MS+MS1").value & "' " & _
                        "AND EDI_Nat_Code = '" & natCode & "' " & _
                        "AND Organisation = '" & frmMain.cboTrust.Text & "'"
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            blnAllowDelete = RS.EOF
            RS.Close
            
            strSQL = "EDI_Msg_Types WHERE EDI_Org_NatCode = '" & natCode & _
                     "' AND EDI_Msg_Format = '" & plist("MS+MS1").value & "'"
         
         Case "Individua"
            eClass.FurtherInfo = "Attempting to delete Individual " & frmMain.ediPr("IN+IN3").value
            
            strSQL = "SELECT Count(DISTINCT EDI_Report_Index) As TotRecs " & _
                     "FROM EDI_Health_Parties " & _
                     "WHERE EDI_HP_Nat_Code = '" & Left(frmMain.cboTrust.Text & " ", 6) & plist("IN+IN2").value & "'"
            
            RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
            
            If RS!totRecs > 0 Then
               MsgBox "There are " & RS!totRecs & " reports still allocated to this GP." & vbCrLf & vbCrLf & _
                      "Deletion refused. Inactivate the GP to cease matching.", vbInformation, "Database consistency"
            Else
            
               With iceCmd
                  .CommandType = adCmdStoredProc
                  .CommandText = "ICECONFIG_Delete_EDI_Individual"
                  .Parameters.Append .CreateParameter("I_Index", adVarChar, adParamInput, 6, plist("SUBID").value)
                  .Execute
               End With
            End If
            
      End Select
   End If
   
   If blnAllowDelete Then
      If iceCmd.CommandText = "" Then
         strSQL = "DELETE FROM " & strSQL
         iceCmd.CommandText = strSQL
      End If
      iceCmd.Execute
   
      tv1.Nodes.Remove NodeId.Index
   Else
      MsgBox "Message Type is either in use or has interchange information recorded against it. " & vbCrLf & _
             "The entry may only be disabled", vbExclamation, "Unable to delete Message Type"
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadEDIRecipients.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node, _
                   Refresh As Boolean)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
'   If Refresh Then
'     Set Control values
      vData = objTV.ReadNodeData(NodeId)
      plist("NATIONAL").value = vData(0)
      plist("NC").value = vData(0)
      plist("REFID").value = vData(1)
'   End If
   
'  SQL string to be executed
   objTV.SQL = "SELECT er.*, err.* " & _
               "FROM EDI_Recipients er " & _
                  "LEFT JOIN EDI_Recipient_Ref err " & _
                  "ON er.Ref_Index = err.Ref_Index " & _
               "WHERE Organisation = '" & frmMain.cboTrust.Text & "' " & _
                  "AND EDI_NatCode = '" & plist("NATIONAL").value & "'"
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "EDI", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "EDI", mCtrl, NodeId
   End If
   
'  Specific node actions eg set active colour
      
'   Set nd(0) = objTV.NodeFromKey("AC", NodeId)
'   nd(0).EnsureVisible
   
   plist("HOLDSTATE").value = plist("HO").value
   
   If plist("AC").value = True Then
      NodeId.ForeColor = BPBLUE
   Else
      NodeId.ForeColor = BPRED
   End If
   
   If plist("HO").value Then
      NodeId.Image = 3
      NodeId.SelectedImage = 3
   Else
      NodeId.Image = plist("ICON").Icon
      NodeId.SelectedImage = plist("ICON").Icon
   End If
   
   
'   fView.Show Fra_EDI
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim i As Integer
   
   With plist
      For i = 1 To .PropertyItems.Count
         If .PropertyItems(i).PageKeys <> "Control" Then
            .PropertyItems(i).value = .PropertyItems(i).defaultValue
         End If
      Next i
   End With
   plist("SUBHEADER").value = ""
   plist("SUBID").value = ""
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(CStr(vData(0)), _
                                                   "NewNode", _
                                                   "SubHeader", _
                                                   , _
                                                   , _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
   End If
   
   vData = objTV.ReadNodeData(objTV.ActiveNode)
   
   Select Case Left(vData(1), 3)
      Case "Spe"
         vData = objTV.ReadNodeData(objTV.ActiveNode)
         frmMain.CurrentLTSIndex = Mid(vData(1), 6)
         plist.ActivePage = "SP"
         
      Case "Msg"
         plist.ActivePage = "MS"
         
      Case "Ind"
         plist("GPCODE").value = ""
         plist.ActivePage = "IN"
         
   End Select
   objTV.RefreshNode = objTV.ActiveNode
   objTV.ActiveNode = newNode
   
'  Specific preparations for a new item
'   <Currently none>
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadEDIRecipients.MenuAddEntry"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewMapping()
   On Error GoTo procEH
   Dim i As Integer

   With plist
      .Redraw = False
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
      Next i
      .Redraw = True
   End With
   
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(, _
                                  , _
                                  mCtrl.NewNodeKey("New", _
                                                   "NewNode", _
                                                   "Detail", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
   End If
   
   plist.ActivePage = "EDI"
   objTV.ActiveNode = newNode
   objTV.RefreshNode = newNode
'   objTView.ActiveNode = newNode
   fView.Show Fra_EDI
   
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubNodes(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim PageId As String
   Dim lookupField(2) As String
   Dim textField As String
   Dim lookupValue As String
   Dim actField As String
   Dim LTSIndex As Long
   
   vData = objTV.ReadNodeData(NodeId.Parent)
   plist("NATIONAL").value = vData(0)
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
'  Read this node details and set up controls
   vData = objTV.ReadNodeData(NodeId)
   
   Select Case Left(vData(1), 3)
      Case "Spe"
         frmMain.CurrentLTSIndex = Val(Mid(vData(1), 6))
'         frmMain.OrgList.Visible = True
'         frmMain.labOrgList.Visible = True
         lookupField(0) = "EDI_Korner_Code"
         lookupField(1) = "EDI_Korner_Code"
         textField = "EDI_Specialty"
         actField = "EDI_Msg_Active"
'         strsql = "SELECT DISTINCT EDI_LTS_Index, EDI_Msg_Type, EDI_OrgCode " & _
                  "FROM EDI_Loc_Specialties s " & _
                     "INNER JOIN EDI_Local_Trader_Settings l " & _
                     "ON s.EDI_LTS_Index = l.EDI_LTS_Index " & _
                  "WHERE EDI_Nat_code = '" & plist("NATIONAL").value & "' " & _
                     "AND Organisation='" & frmMain.cboTrust.Text & "' "

         strSQL = "SELECT * " & _
                  "FROM EDI_Loc_Specialties " & _
                  "WHERE EDI_Nat_code = '" & plist("NATIONAL").value & "' " & _
                     "AND Organisation='" & frmMain.cboTrust.Text & "' " & _
                     "AND EDI_LTS_Index = " & frmMain.CurrentLTSIndex & _
                  " ORDER BY EDI_Korner_Code"
      
      Case "Msg"
         lookupField(0) = "EDI_Msg_Format"
         lookupField(1) = "EDI_Msg_Format"
         textField = "EDI_Delivery_Method"
         actField = "EDI_Msg_Active"
         strSQL = "SELECT * " & _
                  "FROM EDI_Msg_Types " & _
                  "WHERE EDI_Org_NatCode = '" & plist("NATIONAL").value & "' " & _
                     "AND Organisation='" & frmMain.cboTrust.Text & "' " & _
                  "ORDER BY EDI_Msg_Format"
      
      Case "Ind"
         lookupField(0) = "Individual_Index"
         lookupField(1) = "EDI_Natcode"
         textField = "EDI_GP_Name"
         actField = "EDI_Active"
         strSQL = "SELECT * " & _
                  "FROM EDI_Recipient_Individuals " & _
                  "WHERE EDI_ORG_NatCode = '" & plist("NATIONAL").value & "' " & _
                     "AND Organisation='" & frmMain.cboTrust.Text & "' " & _
                  "ORDER BY EDI_GP_Name"
      
   End Select
      
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      If lookupField(1) = "" Then
         lookupValue = ""
      Else
         lookupValue = RS(lookupField(1))
      End If
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(CStr(vData(1)), _
                                                 RS(lookupField(0)), _
                                                 "SubHeader", _
                                                 lookupValue, _
                                                 ms_DELETE), _
                                RS(lookupField(1)) & " - " & RS(textField), _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
      If RS(actField) Then
         nd(0).ForeColor = BPBLUE
      Else
         nd(0).ForeColor = BPRED
      End If
      
      RS.MoveNext
   Loop
   
   RS.Close
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.SubNodes"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubHeader(NodeId As Node)
   On Error GoTo procEH
   Dim pos As Integer
   Dim i As Integer
   Dim iconId As Long
   Dim PageId As String
   
'  Read owner data and set up control items
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   plist("NATIONAL").value = vData(0)
'   pList("<Control key>").value = vData(1)
   
'  Read this node details and set up controls
   vData = objTV.ReadNodeData(NodeId)
'   pList("<Control key>").value = vData(0)
   
'  The sql string to be executed
   plist("SUBHEADER").value = vData(1)
   Select Case Left(vData(0), 3)
      Case "Dum"
         PageId = "X400"
         objTV.SQL = "SELECT * " & _
                     "FROM EDI_Recipients " & _
                     "WHERE EDI_NatCode = '" & plist("NATIONAL") & "'"
         
      Case "Spe"
         PageId = "SP"
         frmMain.CurrentLTSIndex = Val(Mid(vData(0), 6))
         objTV.SQL = "SELECT * " & _
                     "FROM EDI_Loc_Specialties " & _
                     "WHERE EDI_Nat_code = '" & plist("NATIONAL").value & "' " & _
                        "AND EDI_LTS_Index='" & frmMain.CurrentLTSIndex & "' " & _
                        "AND EDI_Korner_Code = " & vData(1)
         
      Case "Msg"
         PageId = "MS"
         objTV.SQL = "SELECT * " & _
                     "FROM EDI_Msg_Types " & _
                     "WHERE EDI_Org_Natcode = '" & plist("NATIONAL").value & "' " & _
                        "AND Organisation='" & frmMain.cboTrust.Text & "' " & _
                        "AND EDI_Msg_Format = '" & vData(1) & "'"
         
      Case "Ind"
         plist("SUBID").value = vData(1)
         PageId = "IN"
         objTV.SQL = "SELECT * " & _
                     "FROM EDI_Recipient_Individuals " & _
                     "WHERE EDI_Org_Natcode = '" & plist("NATIONAL").value & "' " & _
                        "AND Organisation='" & frmMain.cboTrust.Text & "' " & _
                        "AND Individual_Index = '" & vData(1) & "'"
      
   End Select
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, PageId, mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, PageId, mCtrl, NodeId
   End If
   
   If vData(0) = "Ind" Then
      plist("GPCODE").value = plist("IN+IN5").value
   End If
'  The SQL condition (if not a stored procedure)
'   pList.Caption = " WHERE  = '" & pList("NATIONAL").value & "' " & _
'                     "AND = '" & pList("Key2>").value & "'"
'
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.SubHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Public Function Refresh() As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = True
   

'  Specific refresh activities
   
   vData = objTV.ReadNodeData(objTV.ActiveNode)
   Select Case vData(2)
      Case "SubItem"
         objTV.RefreshNode = objTV.ActiveNode.Parent.Parent
         
      Case "SubHeader"
         objTV.RefreshNode = objTV.ActiveNode.Parent
         
      Case "SubNodes"
         objTV.ActiveNode = objTV.ActiveNode.Parent
         objTV.RefreshNode = objTV.ActiveNode
         
      Case "Item"
         objTV.ActiveNode = objTV.ActiveNode.Parent
         objTV.RefreshNode = objTV.ActiveNode
         
      Case "Detail"
         objTV.RefreshNode = objTV.ActiveNode
         
   End Select
   
   Select Case plist.Pages(plist.ActivePage).Key
      Case "EDI"
         If objTV.newNode Then
            DefaultNew plist("NC").value, plist("LC").value
         End If
         
         objTV.ActiveNode.Text = plist("NC").value & " - " & plist("NA").value
         objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("NC").value, _
                                                plist("REFID").value, _
                                                "Detail", _
                                                , _
                                                ms_DELETE)
         Set tv1.SelectedItem = objTV.ActiveNode
         
         If plist("HO").value Then
            objTV.ActiveNode.Image = 3
            objTV.ActiveNode.SelectedImage = 3
         Else
            objTV.ActiveNode.Image = plist("ICON").Icon
         End If
         
         objTV.ActiveNode.EnsureVisible
         
      Case "SP"
         plist("SUBHEADER").value = plist("SP+MS1").value
         
      Case "MS"
         plist("SUBHEADER").value = plist("MS+MS1").value

      Case "IN"
         plist("SUBHEADER").value = plist("IN+IN1").value
         plist("SUBID").value = plist("IN+IN3").value
         plist("GPCODE").value = plist("IN+IN2").value

   End Select
   Set newNode = Nothing
   objTV.ActiveNode = objTV.RefreshNode
End Function

Private Sub ReleaseHeldFiles(natCode As String)
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim outDir As String
   Dim curFile As String
   Dim fExt As String
   Dim fl As File
   Dim fc As Integer
   
   If MsgBox("Are you sure you want to release any held files for " & natCode & "?", vbQuestion Or vbOKCancel, _
             "release Held Files") = vbOK Then
      strSQL = "SELECT Connection_TargetDirectory " & _
               "FROM Connections c " & _
                  "INNER JOIN Connect_Modules cm " & _
                  "ON c.Connection_Name = cm.Connection_Name " & _
               "WHERE Module_Name = 'icemsg.exe' " & _
                  "AND Connection_InFlightMapping = 'EDIRECIPLIST'"
      RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
      outDir = fs.BuildPath(RS!Connection_TargetDirectory, "*.h*")
      RS.Close
      
      curFile = Dir(outDir)
      fc = 0
      Do Until curFile = ""
         
         If InStr(1, curFile, natCode) > 0 Then
            Set fl = fs.GetFile(fs.BuildPath(fs.GetParentFolderName(outDir), curFile))
   '         outFile = fs.BuildPath(fs.GetParentFolderName(outDir), fs.GetBaseName(curFile))
            fExt = "." & Mid(fs.GetExtensionName(curFile), 2)
            fl.Name = fs.GetBaseName(curFile) & fExt
            fc = fc + 1
         End If
         
         curFile = Dir
      Loop
      
      If fc > 0 Then
         MsgBox fc & " held file(s) have been released for " & natCode, vbInformation, "Release held files"
      End If
   Else
      plist("HO").value = True
      MsgBox "The hold parameter ha been reset to 'True'", vbInformation, "Files Not released"
   End If
   
   Set RS = Nothing
End Sub

Public Sub RunWhat(NodeId As Node, _
                   Optional Refresh As Boolean = False)
   On Local Error GoTo procEH
   Dim vData As Variant
   Dim lCode As String
   Dim tNode As MSComctlLib.Node
   Dim RefreshId As String
   Dim strArray() As String
   Dim i As Integer
   Dim ic As PropertiesListCtl.PropertyItem
   
'   If TidyUp = False Then
'      tv1.Visible = False
      TidyUp
      vData = objTV.ReadNodeData(NodeId)
      
'      frmMain.OrgList.Visible = False
'      frmMain.labOrgList.Visible = False
      '  Unload frmTraderDets in case another node has been clicked before completion.
      Unload frmTraderDets
      
      Select Case vData(2)
         Case "FirstView"
            FirstView NodeId
         
         Case "Detail"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId
            Details NodeId, Refresh
            fView.Show
            frmMain.wb.Visible = False
            mCtrl.TreeViewUpdate = False
         
         Case "NewMapping"
            NewMapping
         
         Case "SubNodes"
            objTV.RefreshNode = NodeId.Parent
            SubNodes NodeId
            mCtrl.TreeViewUpdate = False
            Details NodeId.Parent, Refresh
            
         Case "Item"
            objTV.RefreshNode = NodeId.Parent
            Details NodeId.Parent, Refresh
            
         Case "SubHeader"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId.Parent
            SubHeader NodeId
            mCtrl.TreeViewUpdate = False
            
         Case "SubItem"
            objTV.RefreshNode = NodeId.Parent.Parent
            SubHeader NodeId.Parent
            
         Case "X400"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId.Parent.Parent
            SubHeader NodeId.Parent
            mCtrl.TreeViewUpdate = False
         
      End Select
      For i = 1 To plist.PropertyItems.Count
         If plist(i).PageKeys = plist.Pages(plist.ActivePage).Key Then
            Set plist.SelectedItem = plist(i)
            Exit For
         End If
      Next i
'      If Not (pList.SelectedItem Is Nothing) Then
''         Set ic = pList.SelectedItem
'         pList.Redraw = False
'         Set pList.SelectedItem = pList(pList.SelectedItem.Index + 1)
'         pList.Refresh True
''         Set pList.SelectedItem = ic
'         pList.Redraw = True
'      End If
'   End If
'   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function TidyUp() As Boolean
   TidyUp = False
   If Not newNode Is Nothing Then
      If objTV.nodeKey(newNode) = "NewNode" Then
         tv1.Nodes.Remove newNode.Index
         tv1.SelectedItem = tv1.Nodes(1)
         TidyUp = True
      End If
      Set newNode = Nothing
   End If
End Function

Public Function Update(PageId As String) As String
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim newValue As String
   Dim strSQL As String
   Dim intRet As Long
   Dim refId As Integer
   Dim i As Integer
   Dim intHP As Integer
   Dim failInfo As String
   Dim blnContinue As Boolean
   Dim blnAmendHP As Boolean
   Dim mbRVal As Integer
   Dim RS As New ADODB.Recordset
   
'  Run a stored procedure or some special SQL
   failInfo = ""
   
   Select Case PageId
      Case "EDI"
         If plist("NC").value <> "" Then
            If plist("HOLDSTATE").value Then
               If plist("HO").value = False Then
                  ReleaseHeldFiles (plist("NATIONAL").value)
               End If
            End If
            
            With iceCmd
               .ActiveConnection = iceCon
               .CommandType = adCmdStoredProc
               .CommandText = "ICECONFIG_Amend_Recipient"  '  storedproc
               .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
               .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , objTV.newNode)
               .Parameters.Append .CreateParameter("R_Org", adVarChar, adParamInput, 6, frmMain.cboTrust.Text)
               .Parameters.Append .CreateParameter("R_Nat", adVarChar, adParamInput, 10, plist("NC").value)
               .Parameters.Append .CreateParameter("R_OldNat", adVarChar, adParamInput, 10, plist("NATIONAL").value)
               .Parameters.Append .CreateParameter("R_EMIS", adBoolean, adParamInput, , plist("EMIS").value)
               .Parameters.Append .CreateParameter("R_Loc", adVarChar, adParamInput, 10, plist("LC").value)
               .Parameters.Append .CreateParameter("R_Sys", adVarChar, adParamInput, 50, plist("SYS").value)
               .Parameters.Append .CreateParameter("R_Encrypt", adInteger, adParamInput, , False)
               .Parameters.Append .CreateParameter("R_GroupBy", adVarChar, adParamInput, 1, plist("OG").value)
               .Parameters.Append .CreateParameter("R_DelRep", adBoolean, adParamInput, , plist("DR").value)
               .Parameters.Append .CreateParameter("R_Attach", adBoolean, adParamInput, , plist("SA").value)
               .Parameters.Append .CreateParameter("R_Name", adVarChar, adParamInput, 50, plist("NA").value)
               .Parameters.Append .CreateParameter("R_Address", adVarChar, adParamInput, 50, plist("AD").value)
               .Parameters.Append .CreateParameter("R_DTS", adVarChar, adParamInput, 15, plist("DTS").value)
               .Parameters.Append .CreateParameter("R_Con_Name", adVarChar, adParamInput, 35, plist("CNAME").value)
               .Parameters.Append .CreateParameter("R_Con_Tel", adVarChar, adParamInput, 20, plist("CTEL").value)
               .Parameters.Append .CreateParameter("R_Con_Note", adVarChar, adParamInput, 100, plist("CNOTE").value)
               .Parameters.Append .CreateParameter("R_SMTP", adVarChar, adParamInput, 250, plist("SM").value)
               .Parameters.Append .CreateParameter("R_Hold", adBoolean, adParamInput, , plist("HO").value)
               .Parameters.Append .CreateParameter("R_Active", adBoolean, adParamInput, , plist("AC").value)
               .Parameters.Append .CreateParameter("R_Index", adInteger, adParamInput, , Val(plist("REFID").value))
               .Parameters.Append .CreateParameter("Failure", adVarChar, adParamOutput, 20)
               
'               For i = 0 To .Parameters.Count - 1
'                  Debug.Print .Parameters(i).Name & " - " & Len(.Parameters(i)) & " Value " & .Parameters(i).value
'               Next i
               
               .Execute
               
               intRet = .Parameters("Return")
               If intRet = 0 Then
                  '  Set the updated flag for the trader details
               
                  strSQL = "SELECT Ref_Index FROM EDI_Recipients WHERE EDI_NatCode='" & plist("NC").value & "'"
                  Set RS = iceCon.Execute(strSQL)
                  plist("REFID") = RS(0)
                  RS.Close
                  
'                  frmTraderDets.ConfirmUpdate
'                  Unload frmTraderDets
               End If
               
               failInfo = "ICECONFIG_Amend_Recipient (" & .Parameters("Failure") & ")"
               eClass.FurtherInfo = "SQL Error when Adding to/Amending table " & .Parameters("Failure")
            End With
         End If
         
      Case "X400"
         With iceCmd
            .ActiveConnection = iceCon
            .CommandText = "ICECONFIG_Amend_X400"  '  storedproc
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
            .Parameters.Append .CreateParameter("R_Nat", adVarChar, adParamInput, 10, plist("NATIONAL").value)
            .Parameters.Append .CreateParameter("R_X400_GName", adVarChar, adParamInput, 16, plist("X4A").value)
            .Parameters.Append .CreateParameter("R_X400_Surname", adVarChar, adParamInput, 40, plist("X4B").value)
            .Parameters.Append .CreateParameter("R_X400_Init", adVarChar, adParamInput, 5, plist("X4C").value)
            .Parameters.Append .CreateParameter("R_X400_Gen", adVarChar, adParamInput, 3, plist("X4D").value)
            .Parameters.Append .CreateParameter("R_X400_Common", adVarChar, adParamInput, 64, plist("X4E").value)
            .Parameters.Append .CreateParameter("R_X400_Org", adVarChar, adParamInput, 64, plist("X41").value)
            .Parameters.Append .CreateParameter("R_X400_OU1", adVarChar, adParamInput, 32, plist("X42").value)
            .Parameters.Append .CreateParameter("R_X400_OU2", adVarChar, adParamInput, 32, plist("X43").value)
            .Parameters.Append .CreateParameter("R_X400_OU3", adVarChar, adParamInput, 32, plist("X44").value)
            .Parameters.Append .CreateParameter("R_X400_OU4", adVarChar, adParamInput, 32, plist("X45").value)
            .Parameters.Append .CreateParameter("R_X400_PRD", adVarChar, adParamInput, 16, plist("X46").value)
            .Parameters.Append .CreateParameter("R_X400_ADM", adVarChar, adParamInput, 16, plist("X47").value)
            .Parameters.Append .CreateParameter("R_X400_C", adVarChar, adParamInput, 2, plist("X48").value)
            .Execute
            intRet = .Parameters("Return")
            failInfo = "ICECONFIG_Amend_X400"
         End With
      
      Case "SP"
         If plist("SP+MS1").value <> "" Then
            With iceCmd
               .ActiveConnection = iceCon
               .CommandText = "ICECONFIG_Amend_Specialties"    '  "ICECONFIG_Test"
               .CommandType = adCmdStoredProc
               .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
               .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , objTV.newNode)
               .Parameters.Append .CreateParameter("S_OldCode", adVarChar, adParamInput, 3, plist("SUBHEADER").value)
               .Parameters.Append .CreateParameter("S_Org", adVarChar, adParamInput, 10, Trim(frmMain.cboTrust.Text))
               .Parameters.Append .CreateParameter("S_NatCode", adVarChar, adParamInput, 10, plist("NATIONAL").value)
               .Parameters.Append .CreateParameter("S_Korner", adVarChar, adParamInput, 3, plist("SP+MS1").value)
               .Parameters.Append .CreateParameter("S_Specialty", adVarChar, adParamInput, 30, plist("SP+MS2").value)
               .Parameters.Append .CreateParameter("S_Format", adVarChar, adParamInput, 20, plist("SP+MS3").value)
               .Parameters.Append .CreateParameter("S_RunFreq", adVarChar, adParamInput, 4, plist("SP+MS4").value)
               .Parameters.Append .CreateParameter("S_Start1", adVarChar, adParamInput, 5, plist("SP+MS5").value)
               .Parameters.Append .CreateParameter("S_Start2", adVarChar, adParamInput, 5, plist("SP+MS6").value)
               .Parameters.Append .CreateParameter("S_Start3", adVarChar, adParamInput, 5, plist("SP+MS7").value)
               .Parameters.Append .CreateParameter("S_Start4", adVarChar, adParamInput, 5, plist("SP+MS8").value)
               .Parameters.Append .CreateParameter("S_Active", adBoolean, adParamInput, , plist("SP+MS9").value)
               .Parameters.Append .CreateParameter("S_LTSIndex", adInteger, adParamInput, , frmMain.CurrentLTSIndex)
               .Execute
               intRet = .Parameters("Return")
               failInfo = "ICECONFIG_Amend_Specialties"
            End With
         End If
         
      Case "MS"
         If plist("MS+MS1").value <> "" Then
            With iceCmd
               .ActiveConnection = iceCon
               .CommandText = "ICECONFIG_Amend_MsgTypes"
               .CommandType = adCmdStoredProc
               .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
               .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , objTV.newNode)
               .Parameters.Append .CreateParameter("OldFmt", adVarChar, adParamInput, 20, plist("SUBHEADER").value)
               .Parameters.Append .CreateParameter("M_Org", adVarChar, adParamInput, 10, frmMain.cboTrust.Text)
               .Parameters.Append .CreateParameter("M_NatCode", adVarChar, adParamInput, 10, plist("NATIONAL").value)
               .Parameters.Append .CreateParameter("M_MsgFormat", adVarChar, adParamInput, 20, plist("MS+MS1").value)
               .Parameters.Append .CreateParameter("M_Delivery", adVarChar, adParamInput, 1, plist("MS+MS2").value)
               .Parameters.Append .CreateParameter("M_Copy", adVarChar, adParamInput, 10, plist("MS+MS9").value)
               .Parameters.Append .CreateParameter("M_Encrypt", adBoolean, adParamInput, , plist("MS+MS3").value)
               .Parameters.Append .CreateParameter("M_Acks", adBoolean, adParamInput, , plist("MS+MS4").value)
               .Parameters.Append .CreateParameter("M_Testing", adBoolean, adParamInput, , plist("MS+MS5").value)
               .Parameters.Append .CreateParameter("M_Anon", adBoolean, adParamInput, , plist("MS+MS6").value)
               .Parameters.Append .CreateParameter("M_Redirect", adVarChar, adParamInput, 10, plist("MS+MS7").value)
               .Parameters.Append .CreateParameter("M_Active", adBoolean, adParamInput, , plist("MS+MS8").value)
               .Execute
               intRet = .Parameters("Return")
               failInfo = "ICECONFIG_Amend_MsgTypes"
            End With
         End If
         
      Case "IN"
         If plist("IN+IN1").value <> "" Or plist("IN+IN3").value <> "" Then
'            intHP = 7
'            If objTV.newNode = False Then 'And plist("IN+IN5").value <> plist("GPCODE").value Then
'               intHP = IndividualKey3Check
'            End If
            
'            strSQL = "SELECT Count(*) as TotRecs " & _
'                     "FROM EDI_Health_Parties " & _
'                     "WHERE EDI_HP_Nat_Code = '" & plist("NATIONAL").value & "'"
'            RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
'
'            If RS!totRecs = 0 Then
'               blnAmendHP = False
'            Else
'               mbRVal = MsgBox(RS!totRecs & " Health Party records have been processed with the old National Code " & _
'                             plist("NATIONAL").value & ". Amend these to " & plist("IN+IN2").value & "?" & vbCrLf & _
'                             vbCrLf & "Please call Anglia on 01603 819666 if you are unsure", vbYesNoCancel, _
'                             "Database Consistency check")
'               blnAmendHP = (mbRVal = vbYes)
'            End If
            
            frmMain.MousePointer = vbHourglass
            frmMain.Refresh
            
            If mbRVal <> vbCancel Then
               With iceCmd
                  .ActiveConnection = iceCon
                  .CommandType = adCmdStoredProc
                  .CommandText = "ICECONFIG_Amend_Loc_Recp"
               
                  .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
                  .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , objTV.newNode)
                  .Parameters.Append .CreateParameter("HPUpdate", adBoolean, adParamInput, , False)
                  .Parameters.Append .CreateParameter("I_Index", adVarChar, adParamInput, 10, Left(plist("SUBID").value, 10))
                  .Parameters.Append .CreateParameter("I_OldNatCode", adVarChar, adParamInput, 10, plist("NATIONAL").value)
                  .Parameters.Append .CreateParameter("I_Org", adVarChar, adParamInput, 10, frmMain.cboTrust.Text)
                  .Parameters.Append .CreateParameter("I_OrgNatCode", adVarChar, adParamInput, 10, plist("NATIONAL"))
                  .Parameters.Append .CreateParameter("I_GPName", adVarChar, adParamInput, 30, Trim(Left(plist("IN+IN1").value, 30)))
                  .Parameters.Append .CreateParameter("I_NatCode", adVarChar, adParamInput, 10, plist("IN+IN2").value)
                  .Parameters.Append .CreateParameter("I_SMTP", adVarChar, adParamInput, 250, plist("IN+IN4").value)
                  .Parameters.Append .CreateParameter("I_HTML", adBoolean, adParamInput, , IIf(plist("IN+IN5").value = "", False, plist("IN+IN5").value))
                  .Parameters.Append .CreateParameter("I_Active", adBoolean, adParamInput, , plist("IN+IN6").value)
                  .Parameters.Append .CreateParameter("Failure", adVarChar, adParamOutput, 40)
                  
                  .Execute
               End With
            End If
            
            frmMain.MousePointer = vbNormal
         End If
   End Select
   If intRet <> 0 Then
      Err.Raise 3157, "LoadEDIRecipients.Update", "Database Update Failed"
   End If
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   If iceCmd.Parameters("Return").value = 0 Then
      Resume Next
   Else
      If iceCmd.Parameters(iceCmd.Parameters.Count - 1).Name = "Failure" Then
         failInfo = "ICECONFIG_Amend_Lco_Recp " & iceCmd.Parameters("Failure").value
      End If
      eClass.CurrentProcedure = "LoadEDIRecipients.Update"
      eClass.FurtherInfo = failInfo & " - " & iceCmd.Parameters("Return").value
      eClass.Add Err.Number, Err.Description, Err.Source
   End If
End Function

Private Function IndividualKey3Check() As Integer
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim strArray() As String
   Dim i As Integer
   Dim actNatCode As String
   Dim existNatCode As String
   Dim newNatCode As String
   Dim actHPRecs As Long
   Dim existHPRecs As Long
   Dim newHPrecs As Long
   Dim OldClinLeft As Integer
   Dim NewClinRecs As Integer
   Dim blnClinExists As Boolean
   Dim blnKey3Exists As Boolean
   Dim Key3Count As Long
   Dim gpNames As String
   Dim HPStatus As Integer
   Dim blnContinue As Boolean
   Dim mbPrompt As String
   Dim intHPC As Integer
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_Check_EDILocKey3"
      .Parameters.Append .CreateParameter("OrgId", adVarChar, adParamInput, 6, frmMain.cboTrust.Text)
      .Parameters.Append .CreateParameter("GPCode", adVarChar, adParamInput, 8, plist("IN+IN5").value)
      .Parameters.Append .CreateParameter("OldKey3", adVarChar, adParamInput, 30, Left(plist("SUBID").value, 30))
      .Parameters.Append .CreateParameter("NewKey3", adVarChar, adParamInput, 30, Left(plist("IN+IN3").value, 30))
      .Parameters.Append .CreateParameter("ActNat", adVarChar, adParamOutput, 8)
      .Parameters.Append .CreateParameter("ExistNat", adVarChar, adParamOutput, 8)
      .Parameters.Append .CreateParameter("ActHP", adInteger, adParamOutput)
      .Parameters.Append .CreateParameter("ExistHP", adInteger, adParamOutput)
      .Parameters.Append .CreateParameter("NewHP", adInteger, adParamOutput)
      .Parameters.Append .CreateParameter("OldClins", adInteger, adParamOutput)
      .Parameters.Append .CreateParameter("NewClins", adInteger, adParamOutput)
      .Parameters.Append .CreateParameter("Surname", adVarChar, adParamOutput, 35)
      .Parameters.Append .CreateParameter("NewClinExists", adBoolean, adParamOutput)
      .Parameters.Append .CreateParameter("Key3Exists", adInteger, adParamOutput)
      .Parameters.Append .CreateParameter("GPNames", adVarChar, adParamOutput, 1500)
      .Execute
      
      actNatCode = Trim(.Parameters("ActNat").value & "")   '  The National code pointed to by old Key3
      existNatCode = Trim(.Parameters("ExistNat").value & "")  '  The National code pointed to by New Key3
      actHPRecs = .Parameters("ActHP").value '  The no of HP Records in use by the actual Nat code
      existHPRecs = .Parameters("ExistHP").value   '  The No. of HP records in use by the Nat Code from Key3
      newHPrecs = .Parameters("NewHP").value '  The No. of HP records in use by the new Nat Code
      newNatCode = plist("IN+IN5").value
      OldClinLeft = .Parameters("OldClins").value  '  The No. of clincians left pointing to the old Nat code
      NewClinRecs = .Parameters("NewClins").value  '  The No. of clinicians pointing to the New loc Key3 Nat Code
      blnClinExists = .Parameters("NewClinExists").value
      Key3Count = .Parameters("Key3Exists").value
      gpNames = .Parameters("GPNames").value
   End With

'  Check validity of current record
   If actNatCode <> "" Then
      If UCase(actNatCode) <> UCase(Trim(plist("GPCODE").value)) Then
'        Current Key3 does not match Clinicians
         If actNatCode = existNatCode Then
'           The Local Key3 has not been changed
            mbPrompt = "The Key3 " & plist("IN+IN3").value & " currently points to National Code " & existNatCode & _
                       vbCrLf & " The following individuals could reference this national code:" & vbCrLf & vbCrLf
            
            mbPrompt = mbPrompt & GPDetails(gpNames, Key3Count)
            
            If actHPRecs > 0 Then
               mbPrompt = mbPrompt & vbCrLf & _
                          "There are " & actHPRecs & " Health Party records with a National Code of " & actNatCode & vbCrLf
            End If
         End If
      End If
   End If
   
   If UCase(actNatCode) <> UCase(Trim(plist("GPCODE").value)) Then
      If newHPrecs > 0 Then
         mbPrompt = mbPrompt & _
                    newNatCode & " already has " & newHPrecs & " Health Party Records associated with it." & vbCrLf
      End If
   End If

   If mbPrompt <> "" Then
      mbPrompt = mbPrompt & vbCrLf & _
                 "Transfer relevant HP and/or Clinician Records to " & newNatCode & " ?" & vbCrLf & _
                 "(Please call Sunquest on 0845 519 4020 if you are unsure)"
   End If
   
   If mbPrompt <> "" Then
      intHPC = MsgBox(mbPrompt, vbYesNoCancel, "Data integrity warning")
   Else
      intHPC = 7
   End If
   
   If intHPC <> 2 Then
      If plist("GPCODE").value <> plist("IN+IN5").value Then
'        The GP National code has changed
         If OldClinLeft = 1 Then
            mbPrompt = "This will transfer " & actHPRecs & " Health Party records to the national code " & newNatCode & _
                       " (which references " & newHPrecs & " HP records) " & vbCrLf
            If newHPrecs > 0 Then
               mbPrompt = mbPrompt & _
                          "Once completed, the process will be irreversible" & vbCrLf
            Else
               mbPrompt = mbPrompt & _
                          "Once data has been processed against either Individual, this process will be irreversible"
            End If
            
            mbPrompt = mbPrompt & _
                       vbCrLf & vbCrLf & "Continue...? (Please call Sunquest on 0845 519 4020 if you are unsure)"
            
            If MsgBox(mbPrompt, vbYesNo, "Data Integrity Warning") = vbNo Then
               intHPC = 2
            Else
               intHPC = 6
            End If
         
         Else
            intHPC = 7
         End If
      End If
      
      If existNatCode <> "" And (plist("IN+IN3").value <> plist("SUBID").value) Then
'        The new Key3 points to an existing Clinician
         If gpNames = "" Then
'         If Left(existNatCode, 3) = "ICE" Then
            intHPC = MsgBox("This will amend the Clinician record national code from " & existNatCode & _
                            " to " & newNatCode & vbCrLf & "Amend the corressponding health party records (recommended)", _
                            vbYesNoCancel, "Local Key3 Amendment warning")
         Else
            If UCase(existNatCode) <> UCase(newNatCode) Then
               mbPrompt = "The Key3 " & plist("IN+IN3").value & " currently points to National Code " & existNatCode & _
                          vbCrLf & " It is referenced from the following individuals:" & vbCrLf & vbCrLf
               strArray = Split(gpNames, "|")
               mbPrompt = mbPrompt & GPDetails(gpNames, Key3Count)
               mbPrompt = mbPrompt & vbCrLf & _
                          " Please amend these entries or select another Local Key3."
               MsgBox mbPrompt, vbCritical + vbOKOnly, "Data Integrity warning"
               intHPC = 2
            End If
         End If
      End If
   End If
   
   Set iceCmd = Nothing
   IndividualKey3Check = intHPC
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipients.IndividualKey3Check"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Private Function GPDetails(data As String, Key3Count As Long) As String
   Dim strArray() As String
   Dim i As Integer
   Dim rStr As String
   
   strArray = Split(data, "|")
   
   For i = 0 To UBound(strArray) Step 6
      rStr = rStr & _
            strArray(i + 2) & vbTab & strArray(i + 1) & vbTab & IIf(Len(strArray(i + 1)) > 7, "", vbTab) & strArray(i) & vbCrLf
      If Len(rStr) > 790 And (i + 6) < UBound(strArray) Then
         rStr = rStr & "...(" & Key3Count - ((i / 6) + 1) & ") further entries..." & vbCrLf
         Exit For
      End If
   Next i
   GPDetails = rStr
End Function




