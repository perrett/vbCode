VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadRepList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private tvNode As MSComctlLib.Node
Private strArray() As String
Private strSQL As String
Private mCtrl As New ManageControls
Private vData As Variant

Private Sub Class_Initialize()
   Set tv1 = frmMain.TreeView1
   fView.FrameToShow = Fra_HELP
'   objTView.Origin = "R"
End Sub

Public Sub Delete(tNode As MSComctlLib.Node)
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim headerId As Long
   Dim strComment As String
   Dim RepId As String
   
   vData = objTV.ReadNodeData(tNode)
   RepId = vData(1)
   iceCon.BeginTrans
   eClass.FurtherInfo = "Report: " & tNode.Text & " - Index: " & RepId
   strSQL = "SELECT Service_ImpExp_Headers.Service_ImpExp_ID " & _
            "FROM Service_ImpExp_Messages " & _
               "INNER JOIN Service_ImpExp_Headers ON " & _
               "Service_ImpExp_Messages.Service_ImpExp_ID = Service_ImpExp_Headers.Service_ImpExp_ID " & _
            "WHERE (Service_ImpExp_Messages.Service_Report_Index = '" & RepId & "' " & _
               "AND Service_ImpExp_Headers.Service_Type = 1)"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   headerId = RS!Service_ImpExp_Id
   RS.Close
   
   strComment = "Report " & RepId & " deleted from Replist by user: " & userId & "."
   strSQL = "SET dateformat dmy; " & _
            "INSERT INTO Service_ImpExp_Comments " & _
               "(Service_ImpExp_Id," & _
               " Service_ImpExp_Comment," & _
               " Service_ImpExp_Process," & _
               " Code," & _
               "Date_Added) " & _
            "VALUES (" & _
               headerId & ", " & _
               "'" & strComment & "', " & _
               "'IceConfig', " & _
               0 & ", " & _
               "'" & Format(Now(), "dd/mm/yyyy hh:nn:ss") & "')"
   iceCon.Execute strSQL
   
   eClass.FurtherInfo = "Delete from replist"
   strSQL = "DELETE FROM EDI_Rep_List " & _
            "WHERE EDI_Report_Index = " & RepId
   iceCon.Execute strSQL

   frmMain.TreeView1.Nodes.Remove tNode.Index
   iceCon.CommitTrans
   eClass.LogMessage "Replist entry " & objTV.nodeKey(tNode) & " - Index " & RepId & " deleted from EDI_Rep_List by " & userId
   If TypeName(frmMain.TreeView1.SelectedItem) <> "Nothing" Then
      frmMain.NodeClick frmMain.TreeView1.SelectedItem
   Else
      fView.Show Fra_HELP
   End If
   fView.FrameToShow = Fra_LOGVIEW
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadRepList.DeleteEntry"
   iceCon.RollbackTrans
   transCount = 0
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node)
   Dim RS As New ADODB.Recordset
   Dim iceCmd As New ADODB.Command
   Dim nd As Node
   Dim i As Integer
   vData = objTV.ReadNodeData(NodeId)
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   'strSQL = "SELECT EDI_Report_Index, " & _
               "Forename, " & _
               "Surname, " & _
               "Left(sr.Service_Report_Id,CharIndex(' ',sr.Service_Report_Id)-1) AS RepId " & _
            "FROM EDI_Rep_List " & _
               "INNER JOIN Service_Reports sr " & _
                  "INNER JOIN Service_Requests sq " & _
                     "INNER JOIN Patient_Local_Ids pl " & _
                        "INNER JOIN Patient p " & _
                        "ON pl.Patient_Id_Key = p.Patient_Id_Key " & _
                     "ON sq.Patient_Id = Id " & _
                  "ON sr.Service_Request_Index = sq.Service_Request_Index " & _
               "ON EDI_Report_Index = Service_Report_Index " & _
            "WHERE EDI_Loc_Nat_Code_To = '" & vData(0) & "' " & _
               "AND EDI_Rep_List.EDI_LTS_Index = " & vData(1)
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_CurrentReplist"
      .Parameters.Append .CreateParameter("natCode", adVarChar, adParamInput, 6, vData(0))
      .Parameters.Append .CreateParameter("LTIndex", adInteger, adParamInput, , vData(1))

      Set RS = .Execute
   End With

   'RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Do Until RS.EOF
      Set nd = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              RS!EDI_Report_Index, _
                                              "Destination", _
                                             , _
                                             ms_DELETE), _
                             RS!RepId & " (" & RS!Surname & " " & RS!Forename & ")", _
                             23, _
                             23)
      tv1.Nodes.Add nd, _
                    tvwChild, _
                    mCtrl.NewNodeKey("Temp", _
                                     "Temp", _
                                     "None"), _
                    "Please wait...", _
                    23, _
                    23
                    
      RS.MoveNext
   Loop
   
   RS.Close
   
   Set RS = Nothing
End Sub

Public Sub FirstView()
   Dim RS As New ADODB.Recordset
   Dim nd As Node
   Dim totRecs As Long
   
   strSQL = "SELECT Count(*) " & _
            "FROM EDI_Rep_List"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   With frmMain.SSListBar1.Groups("EDI Management").ListItems("RepList Entries")
      .Text = RS(0) & " " & .Text
   End With
   RS.Close
   
   strSQL = "SELECT DISTINCT EDI_Loc_Nat_Code_To, " & _
               "Count(EDI_Report_Index) as RepCount, " & _
               "erl.EDI_LTS_Index, " & _
               "EDI_Msg_Type " & _
            "FROM EDI_Rep_List erl " & _
               "INNER JOIN EDI_Local_Trader_Settings el " & _
               "ON erl.EDI_LTS_Index = el.EDI_LTS_Index " & _
            "GROUP BY EDI_Loc_Nat_Code_To, erl.EDI_LTS_Index, EDI_Msg_Type"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Do Until RS.EOF
      Set nd = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey(RS!EDI_Loc_Nat_Code_To, _
                                              RS!EDI_LTS_Index, _
                                             "Details"), _
                             "(" & RS!RepCount & ") " & RS!EDI_Loc_Nat_Code_To & " (" & RS!EDI_Msg_Type & ")", _
                             23, _
                             23)
      tv1.Nodes.Add nd, _
                    tvwChild, _
                    mCtrl.NewNodeKey("Temp", _
                                     "Temp", _
                                     "None"), _
                    "Please wait...", _
                    23, _
                    23
      RS.MoveNext
   Loop
   
   RS.Close
   Set RS = Nothing
End Sub

Public Sub ReadRepList()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim strSQL As String
   Dim hNode As MSComctlLib.Node
   Dim headerId As Long
   Dim strKey As String
   Dim totReps As Integer
   Dim showReps As Integer
   
   strSQL = "SELECT COUNT(EDI_Report_Index) as TotEntries " & _
            "FROM EDI_Rep_List"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   totReps = RS!totEntries
   RS.Close
   showReps = Val(Read_Ini_Var("GENERAL", "MaxRepListEntries", iniFile))
   If showReps = 0 Then
      showReps = 50
   End If
   
   strSQL = "SELECT TOP " & showReps & " erl.*, em"
   If totReps > showReps Then
      strSQL = "SELECT TOP " & showReps & " * " & _
               "FROM EDI_Rep_List " & _
               "ORDER BY Date_Added DESC"
      MsgBox "There are " & totReps & " entries to be processed on the Report List. " & _
             "Only the latest " & showReps & " will be shown", vbExclamation, "Too many Entries"
   Else
      strSQL = "SELECT * " & _
               "FROM EDI_Rep_List"
   End If
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Set tv1 = frmMain.TreeView1
   tv1.Visible = False
   Do Until RS.EOF
      eClass.FurtherInfo = "EDI_Service_Id = " & RS!EDI_Service_Id
'      strSQL = "SELECT Service_ImpExp_Messages.Service_ImpExp_Message_ID " & _
'               "FROM Service_ImpExp_Messages " & _
'                  "INNER JOIN Service_ImpExp_Headers ON " & _
'                  "Service_ImpExp_Messages.Service_ImpExp_ID = Service_ImpExp_Headers.Service_ImpExp_ID " & _
'               "WHERE (Service_ImpExp_Messages.Service_Report_Index = '" & RS!EDI_Report_Index & "' " & _
'                  "AND Service_ImpExp_Headers.Service_Type = 1)"
'      RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
'      headerId = RS2!Service_ImpExp_Message_ID
'      RS2.Close
      strKey = frmMain.cboTrust.Text & "+" & Trim(RS!EDI_Report_Index) & "+" & CStr(headerId)
      Set tvNode = tv1.Nodes.Add(, _
                                 , _
                                 mCtrl.NewNodeKey(RS!EDI_Report_Index, _
                                                     strKey, _
                                                     "Report", _
                                                     , _
                                                     ms_DELETE), _
                                 RS!EDI_Service_Id, _
                                 23, _
                                 23)

      tv1.Nodes.Add tvNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!EDI_Report_Index, RS!EDI_Loc_Nat_Code_To), _
                    "Added on - " & RS!Date_Added, _
                    23, _
                    23
      tv1.Nodes.Add tvNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!EDI_Report_Index, RS!EDI_Provider_Org), _
                    "Organisation - " & RS!EDI_Provider_Org, _
                    23, _
                    23
      tv1.Nodes.Add tvNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!EDI_Report_Index, RS!Date_Added), _
                    "Date on List - " & RS!Date_Added, _
                    23, _
                    23
      Set hNode = tv1.Nodes.Add(tvNode, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!EDI_Report_Index, RS!EDI_LTS_Index, "Trader"), _
                                "Local Trader Details...", _
                                2, _
                                2)
      tv1.Nodes.Add hNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!EDI_Report_Index, RS!EDI_Individual_Index_To), _
                    "Please Wait...", _
                    23, _
                    23
      Set hNode = tv1.Nodes.Add(tvNode, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!EDI_Report_Index, _
                                                    RS!EDI_Individual_Index_To, _
                                                   "SendTo"), _
                                "Destination...", _
                                4, _
                                4)
      tv1.Nodes.Add hNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!EDI_Report_Index, RS!EDI_Individual_Index_To), _
                    "Please Wait...", _
                    23, _
                    23
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
   tv1.Visible = True
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "ICEConfigV3.LoadRepList.ReadRepList"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function Refresh() As String
   objTV.RefreshNode = objTV.ActiveNode
End Function

'Public Sub Report(NodeId As node)
'   Dim OrgStr As String
'   Dim IDStr As String
'   Dim reQueuestr As String
'   Dim RS As New ADODB.Recordset
'
'   strArray = Split(objTView.NodeKey(newNode.Key), "+")
'   OrgStr = strArray(0)
'   IDStr = strArray(1)
'   reQueuestr = strArray(2)
'
'   strSQL = "SELECT * " & _
'            "FROM Service_ImpExp_Messages " & _
'            "WHERE Service_ImpExp_Message_Id = " & reQueuestr & _
'               " AND Service_Report_Index = " & IDStr
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'   If RS!Rejected_Count > 0 Then
'      LogText.Text = "Error: "
'   ElseIf RS!Rejected_Count < 0 Then
'      LogText.Text = "Comment: "
'      MailBtn.Visible = False
'   End If
'   LogText.Text = ""
'   LogText.Text = LogText.Text & RS!Service_Message_Comment & vbCrLf & vbCrLf
'   RS.Close
'   Set RS = Nothing
'
'   Call Load_Report(OrgStr, IDStr)
'   fView.Show Fra_LOGVIEW, fData
'
'End Sub

Public Sub RunWhat(NodeId As MSComctlLib.Node, Optional Refresh As String = "")
   On Local Error GoTo procEH
   Dim vData As Variant
   Dim rType As String
   Dim tNode As MSComctlLib.Node
   Dim RefreshId As String
   Dim strArray() As String
   
   rType = objTV.nodeKey(objTV.TopLevelNode(NodeId))
   vData = objTV.ReadNodeData(NodeId)
   strArray = Split(vData(0), "|")
   NodeId.Selected = True
   Select Case vData(2)
      Case "Trader"
         Trader CStr(vData(1)), NodeId
         
      Case "SendTo"
         SendTo CStr(vData(1)), NodeId
         
      Case "Details"
         Details NodeId
         
      Case "Destination"
         Destination NodeId
         
   End Select
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadEDIRecipents.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Trader(Index As Long, NodeId As MSComctlLib.Node)
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   
   strSQL = "SELECT * " & _
            "FROM EDI_Local_Trader_Settings " & _
            "WHERE EDI_LTS_Index = " & Index
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   If RS.RecordCount > 0 Then
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, RS!Organisation), _
                    "Organisation - " & RS!Organisation, _
                    2, _
                    2
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, RS!EDI_Name), _
                    RS!EDI_Name, _
                    2, _
                    2
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, RS!EDI_Trader_Code), _
                    "Trader Code - " & RS!EDI_Trader_Code, _
                    2, _
                    2
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, RS!EDI_Free_Part), _
                    "Free Part - " & RS!EDI_Free_Part, _
                    2, _
                    2
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, RS!EDI_Msg_Type), _
                    "Source System - " & RS!EDI_Msg_Type, _
                    2, _
                    2
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, RS!EDI_Default_Dept), _
                    "Default Dept. - " & RS!EDI_Default_Dept, _
                    2, _
                    2
   End If
   RS.Close
   Set RS = Nothing
End Sub

Public Sub Destination(NodeId As Node)
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim iceCmd As New ADODB.Command
   Dim nd As Node
   Dim i As Integer
   Dim natCode As String
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   natCode = vData(0)

   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   vData = objTV.ReadNodeData(NodeId)
   
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_ReadReplistEntry"
      .Parameters.Append .CreateParameter("repId", adInteger, adParamInput, , vData(1))
      Set RS = .Execute
   End With
   
'   strSQL = "SELECT EDI_Loc_Nat_Code_To, " & _
               "EDI_Org_natcode, " & _
               "EDI_NatCode, " & _
               "c1.Clinician_National_Code as NatCode902, " & _
               "c2.Clinician_National_Code as NatCode906 " & _
            "FROM EDI_Rep_List erl " & _
               "LEFT JOIN EDI_Matching em " & _
                  "INNER JOIN EDI_Recipient_Individuals ei " & _
                  "ON em.Individual_Index = ei.Individual_Index " & _
               "ON EDI_Individual_Index_To = MatchIndex " & _
               "INNER JOIN EDI_Health_Parties hp1 " & _
                  "LEFT JOIN Clinician c1 " & _
                  "ON c1.Clinician_National_Code = hp1.EDI_HP_Nat_Code " & _
               "ON erl.EDI_Report_Index = hp1.EDI_Report_Index " & _
                  "AND hp1.EDI_HP_Type = '902' " & _
               "INNER JOIN EDI_Health_Parties hp2 " & _
                  "LEFT JOIN Clinician c2 " & _
                  "ON c2.Clinician_National_Code = hp2.EDI_HP_Nat_Code " & _
               "ON erl.EDI_Report_Index = hp2.EDI_Report_Index " & _
                  "AND hp2.EDI_HP_Type = '906' " & _
            "WHERE erl.EDI_Report_Index = " & RepId
   
   strSQL = "SELECT EDI_Msg_Format " & _
            "FROM EDI_Loc_Specialties " & _
            "WHERE EDI_Nat_Code = '" & natCode & "'" & _
               "AND EDI_Korner_Code = '" & RS!EDI_Rep_Specialty & "'"
   RS2.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Set nd = tv1.Nodes.Add(NodeId, _
                          tvwChild, _
                          mCtrl.NewNodeKey(CStr(vData(0)), _
                                           "Temp", _
                                           "None"), _
                          "Specialty " & RS!EDI_Rep_Specialty & " via " & Mid(RS2!EDI_Msg_Format, InStr(1, RS2!EDI_Msg_Format, ",") + 1), _
                          23, _
                          23)
   RS2.Close
   
   If RS!Replist_Practice <> RS!Index_Practice Then
      Set nd = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              "Temp", _
                                              "None"), _
                             "Individual index practice = " & RS!Index_Practice, _
                             23, _
                             23)
      nd.ForeColor = BPRED
   End If
   
   If RS!Replist_Practice <> RS!HP_Practice Then
      Set nd = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              "Temp", _
                                              "None"), _
                             "Health Party practice = " & RS!HP_Practice, _
                             23, _
                             23)
      nd.ForeColor = BPRED
   End If
   
   If IsNull(RS!NatCode_Index) Then
      Set nd = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              "Temp", _
                                              "None"), _
                             "Recipient = " & RS!NatCode_902, _
                             23, _
                             23)
      nd.ForeColor = BPBLUE
   Else
      Set nd = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              "Temp", _
                                              "None"), _
                             "Recipient = " & RS!NatCode_Index, _
                             23, _
                             23)
      If RS!NatCode_Index <> RS!NatCode_902 Then
         nd.ForeColor = BPRED
         nd.Text = "Index Recipient = " & RS!NatCode_Index
         
         Set nd = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(CStr(vData(0)), _
                                                 "Temp", _
                                                 "None"), _
                                "HP Recipient = " & RS!NatCode_902, _
                                23, _
                                23)
         nd.ForeColor = BPBLUE
      Else
         nd.ForeColor = BPGREEN
      End If
   End If
   
   If RS!NatCode_902 <> RS!NatCode_906 Then
      Set nd = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              "Temp", _
                                              "None"), _
                             "Requestor = " & RS!NatCode_906, _
                             23, _
                             23)
      nd.ForeColor = BPBLUE
   End If
   
   RS.Close
   Set RS2 = Nothing
   Set RS = Nothing
   Set iceCmd = Nothing
End Sub

Public Sub SendTo(Index As Long, NodeId As MSComctlLib.Node)
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   strSQL = "SELECT * FROM EDI_Recipient_Individuals " & _
            "WHERE Individual_Index = " & Index
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.RecordCount > 0 Then
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, RS!EDI_NatCode), "National GP Code - " & RS!EDI_NatCode, 4, 4
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, RS!EDI_GP_Name), "GP Name - " & RS!EDI_GP_Name, 4, 4
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, RS!EDI_Org_NatCode), "National Practice Code - " & RS!EDI_Org_NatCode, 4, 4
   Else
      RS.Close
      strSQL = "SELECT * " & _
               "FROM EDI_Rep_List " & _
               "WHERE EDI_Report_Index = " & objTV.NodeLevel(NodeId)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, "None"), "National GP Code - N/A", 4, 4
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, "NoGP"), "GP Name - Not specified", 4, 4
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, RS!EDI_Loc_Nat_Code_To), "National Practice Code - " & RS!EDI_Loc_Nat_Code_To, 4, 4
      
   End If
   RS.Close
   
   strSQL = "SELECT EDI_Loc_Specialties.EDI_Msg_Format, EDI_Loc_Specialties.EDI_Korner_Code, EDI_Loc_Specialties.EDI_Specialty " & _
            "FROM EDI_Rep_List " & _
               "INNER JOIN Service_Reports ON " & _
               "EDI_Rep_List.EDI_Report_Index = Service_Reports.Service_Report_Index " & _
                  "INNER JOIN EDI_Loc_Specialties ON " & _
                  "Left(Service_Reports.Service_Report_Type, 3) = EDI_Loc_Specialties.EDI_Korner_Code " & _
                     "AND EDI_Rep_List.EDI_Loc_Nat_Code_To = EDI_Loc_Specialties.EDI_Nat_Code " & _
            "WHERE (EDI_Rep_List.EDI_Report_Index = " & objTV.NodeLevel(NodeId) & ")"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, "MsgType"), "Format - " & RS!EDI_Msg_Format, 4, 4
      tv1.Nodes.Add NodeId, tvwChild, mCtrl.NewNodeKey(1, "Specialty"), "Specialty - " & RS!EDI_Specialty & " (" & RS!EDI_Korner_Code & ")", 4, 4
   Else
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, "MsgType"), _
                    "Format not available - Message will not be sent", _
                    4, _
                    4
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(1, "MsgType"), _
                    "Specialty not available - Message will not be sent", _
                    4, _
                    4
      NodeId.ForeColor = BPRED
      NodeId.Parent.ForeColor = BPRED
   End If
   RS.Close
   Set RS = Nothing
End Sub
