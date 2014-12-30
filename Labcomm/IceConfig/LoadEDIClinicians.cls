VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadEDIClinicians"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Private newIceNum As Long
Private curNode As Node

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim i As Integer
   
   Set tv1 = frmMain.TreeView1
'   tv1.Style = tvwPictureText
'
'   tv1.Indentation = 500
   
   Set plist = frmMain.ediPr
   mCtrl.TreeViewUpdate = True
   
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"
      .PropertyItems.Add "SUBICON", "Second Icon to use", plpsString, , "Hidden"

'     Add control items as required
      .PropertyItems.Add "CNAT", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Nat", "Clinician"
      
'     Set up the required property items
      With .PropertyItems
         .Add "NCODE", "National Code", plpsString, , "An ICE number indicates national code not known."
         .Add "SURNAME", "Surname", plpsString, , "Mandatory - The Clinician's Name"
         .Add "CSPEC", "Specialty", plpsCustom, , "Read only - The specialty of this clinician"
      End With
      
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 4
      .PropertyItems("SUBICON").PageKeys = "Control"
      .PropertyItems("SUBICON").Icon = 5

'     Set other Control items to the page key
      .PropertyItems("CNAT").PageKeys = "Control"
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
      
      With .PropertyItems("NCODE")
         .Tag = "Clinician_National_Code"
         .DefaultExtension = "Item"
         .ReadOnly = True
      End With
      With .PropertyItems("SURNAME")
         .Tag = "Clinician_Surname"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("CSPEC")
         .Tag = "Clinician_Speciality_Code"
         .DefaultExtension = "Item"
         .ReadOnly = True
      End With
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
   
   plist.Pages("Nat").Caption = "ICECONFIG_UpdateClinicians"
'   plist.Pages("Loc").Caption = "ICECONFIG_UpdateClinicians"
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.Class.Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub FirstView()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   
   tv1.Nodes.Clear
   
   strSQL = "SELECT DISTINCT cli.EDI_LTS_Index, EDI_Msg_Type, EDI_OrgCode " & _
            "FROM Clinician_local_Id cli " & _
               "INNER JOIN EDI_Local_Trader_Settings et " & _
               "ON cli.EDI_LTS_Index = et.EDI_LTS_Index"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!EDI_LTS_Index, _
                                                 RS!EDI_Msg_Type, _
                                                 "ShowClinicians"), _
                                RS!EDI_OrgCode & " - " & RS!EDI_Msg_Type, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Temp", _
                                     "Temp", _
                                     "None"), _
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
   eClass.CurrentProcedure = "LoadEDIClinicians.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub ShowClinicians(NodeId As Node)
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim sampTxt As String
   Dim i As Integer
   
'   tv1.Nodes.Clear
   For i = 1 To objTV.ActiveNode.Children
      tv1.Nodes.Remove objTV.ActiveNode.child.Index
   Next i
   
   vData = objTV.ReadNodeData(NodeId)
   
   strSQL = "SELECT DISTINCT c.Clinician_National_Code, " & _
               "Clinician_Surname, " & _
               "ei1.EDI_NatCode As ClinNat, " & _
               "ei2.EDI_NatCode As Key3Nat, " & _
               "Substring(Clinician_Local_Code,7,10) As LocalId " & _
            "FROM Clinician c " & _
               "LEFT JOIN EDI_Recipient_Individuals ei1 " & _
               "ON c.Clinician_National_Code = ei1.EDI_NatCode " & _
               "INNER JOIN Clinician_Local_Id cl " & _
                  "LEFT JOIN EDI_Matching em " & _
                     "INNER JOIN EDI_Recipient_Individuals ei2 " & _
                     "ON em.Individual_Index = ei2.Individual_Index " & _
                  "ON (Substring(Clinician_Local_Code,7,10) = EDI_Local_Key3 " & _
                     "AND EDI_LTS_Index = EDI_Local_Key2) " & _
               "ON c.Clinician_Index = cl.Clinician_Index " & _
            "WHERE EDI_LTS_Index = " & vData(0) & _
            " ORDER BY Clinician_Surname, LocalId"
      
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
'      With plist
'         .PropertyItems("<Key>").DefaultValue = "<Default>"
'      End With
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Clinician_National_Code, _
                                                 Trim(RS!LocalId & ""), _
                                                 "Details", _
                                                 RS!Clinician_Surname & "", _
                                                 ms_DISABLED), _
                                Trim(RS!Clinician_Surname & "") & " (" & Trim(RS!LocalId & "") & ")", _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
      If IsNull(RS!ClinNat) Then
         If IsNull(RS!Key3Nat) Then
            nd(0).ForeColor = BPBLUE
         Else
            nd(0).ForeColor = BPRED
'            Set nd(1) = tv1.Nodes.Add(nd(0), _
                                      tvwChild, _
                                      mCtrl.NewNodeKey(RS!Clinician_National_Code, _
                                                       Trim(RS!Key3Nat & ""), _
                                                       "ShowKey3"), _
                                      RS!localId, _
                                      plist("ICON").Icon, _
                                      plist("ICON").Icon)
'            nd(1).ForeColor = nd(0).ForeColor
         End If
      
      ElseIf RS!Clinician_National_Code = RS!ClinNat Then
         If RS!Clinician_National_Code = RS!Key3Nat Then
            nd(0).ForeColor = BPGREEN
         Else
            nd(0).ForeColor = BPBLUE
         End If
      
      Else
         nd(0).ForeColor = BPRED
      End If
      
      
'      Set nd(1) = tv1.Nodes.Add(nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Clinician_National_Code, _
                                                 Trim(RS!Key3Nat & ""), _
                                                 "ShowKey3"), _
                                RS!localId, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
'      nd(1).Bold = True
'      nd(1).ForeColor = nd(0).ForeColor
      
'      If IsNull(RS!EDI_NatCode) Then
'         nd(0).ForeColor = BPBLUE
'      ElseIf RS!Clinician_National_Code <> RS!EDI_NatCode Then
'         nd(0).ForeColor = BPRED
'      Else
'         nd(0).ForeColor = BPGREEN
'      End If
'      If RS!Clinician_Active Then
'         nd(0).ForeColor = BPGREEN
'      Else
'         nd(0).ForeColor = BPRED
'      End If
'
'      tv1.Nodes.Add nd(0), _
'                    tvwChild, _
'                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
'                    "Please wait...", _
'                    2, _
'                    2
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
   eClass.CurrentProcedure = "LoadClinicians.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node, _
                   Refresh As Boolean)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
'   NodeId.Expanded = True
   vData = objTV.ReadNodeData(NodeId)

   plist("CNAT").value = vData(0)
   
   If NodeId.ForeColor = BPBLUE Then
      plist("SURNAME").ReadOnly = False
      plist("SURNAME").Description = "Mandatory - The Clinician's Name"
   ElseIf NodeId.ForeColor = BPGREEN Then
      plist("SURNAME").ReadOnly = True
      plist("SURNAME").Description = "Registered EDI Individual - READ ONLY"
   End If
      
   frmEDIClinicians.ClinicianId = vData(0)
   frmEDIClinicians.ClinicianName = vData(4)
'   If Not curNode Is Nothing Then
'      For i = 0 To curNode.Children - 1
'         tv1.Nodes.Remove curNode.Child.Index
'      Next i
'   End If
   
   Set curNode = NodeId
   
   objTV.SQL = "SELECT * " & _
               "FROM Clinician " & _
               "WHERE Clinician_National_Code = '" & vData(0) & "'"
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Nat", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Nat", mCtrl, NodeId
   End If
   
   If NodeId.ForeColor = BPGREEN Then
      plist("NCODE").ReadOnly = True
      plist("NCODE").Description = "Read only - set from EDI Recipients"
   ElseIf NodeId.ForeColor = BPRED Then
      With frmLogon
         .ValidateInForm = True
         .Show 1
         .Hide
         If .Text2.Text = Format(Now(), "yyyymmdd") Then
            Unload frmLogon
            frmEDIClinicians.Show 1
            FirstView
         Else
            Unload frmLogon
            MsgBox "The Admin password incorrect - access denied." & vbCrLf & vbCrLf & _
                   "Please call Sunquest if you need access", vbInformation, "Protected Feature"
         End If
      End With
      
   Else
'      plist("NCODE").ReadOnly = False
      plist("NCODE").Description = "An ICE number indicates national code not known."
   End If
'   plist.Caption = "WHERE Clinician_National_Code = '" & vData(0) & "'"
   
'   Set nd(0) = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey("HDR", _
                                              "HDR", _
                                              "None"), _
                             "Local Identity(ies) for: " & vData(0))
'   nd(0).Bold = True
   
'   strSQL = "SELECT c1.Clinician_National_Code, " & _
               "Substring(Clinician_Local_code,7,10) As Local_Id, " & _
               "c1.Clinician_Surname, " & _
               "g1.Clinician_Name As EDI_Code_Name, " & _
               "g2.Clinician_Name As Clin_Code_Name, " & _
               "c1.Clinician_Speciality_Code, " & _
               "ei2.EDI_NatCode, " & _
               "IsNull(EDI_OP_Name, ei2.EDI_GP_Name) As GP_Name " & _
            "FROM Clinician c1 " & _
               "LEFT JOIN EDI_Recipient_Individuals ei1 " & _
               "ON c1.Clinician_National_Code = ei1.EDI_NatCode " & _
               "INNER JOIN Clinician_Local_Id cl " & _
                  "LEFT JOIN EDI_Matching em " & _
                     "INNER JOIN EDI_Recipient_Individuals ei2 " & _
                     "ON em.Individual_Index = ei2.Individual_Index " & _
                  "ON Substring(Clinician_Local_Code,7,10) = EDI_Local_Key3 " & _
               "ON c1.Clinician_National_Code = cl.Clinician_National_Code " & _
               "LEFT JOIN National_GPs g1 " & _
               "ON ei2.EDI_NatCode = g1.Clinician_National_Code " & _
               "LEFT JOIN National_GPs g2 " & _
               "ON c1.Clinician_National_Code = g2.Clinician_National_Code " & _
            "WHERE c1.Clinician_National_Code = '" & vData(0) & "'"

'   strSQL = "SELECT Substring(Clinician_Local_Code,7,10) as Local_Id, " & _
               "EDI_NatCode, " & _
               "EDI_Org_NatCode, " & _
               "EDI_GP_Name " & _
            "FROM Clinician_Local_Id " & _
               "LEFT JOIN EDI_Matching em " & _
                  "INNER JOIN EDI_Recipient_Individuals ei " & _
                  "ON em.Individual_Index = ei.Individual_Index " & _
               "ON Substring(Clinician_Local_Code,7,10) = EDI_Local_Key3 " & _
            "WHERE Clinician_National_Code = '" & vData(0) & "'"
'   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
'
'   Do Until RS.EOF
'      Set nd(0) = tv1.Nodes.Add(NodeId, _
'                                tvwChild, _
'                                mCtrl.NewNodeKey("LKey", _
'                                                 "LKey", _
'                                                 "None"), _
'                                RS!Local_Id)
'      If IsNull(RS!EDI_NatCode) Then
'         nd(0).ForeColor = BPBLUE
'      ElseIf vData(0) <> RS!EDI_NatCode Then
'         nd(0).ForeColor = BPRED
'      Else
'         nd(0).ForeColor = BPGREEN
'      End If
'
'      If RS!EDI_NatCode <> vData(0) Then
'         Set nd(1) = tv1.Nodes.Add(nd(0), _
'                                   tvwChild, _
'                                   mCtrl.NewNodeKey("EDI", _
'                                                    "EDI", _
'                                                    "None"), _
'                                   RS!EDI_NatCode)
'         nd(1).ForeColor = BPRED
'         nd(0).Expanded = True
'      End If
'      RS.MoveNext
'   Loop
'   RS.Close
   Set RS = Nothing
   
'  Specific node actions eg set active colour
      
'   Set nd(0) = objTV.NodeFromKey("<KeyId>", NodeId)
'   nd(0).EnsureVisible
'   If plist("ACT").value = True Then
'      NodeId.ForeColor = BPGREEN
'   Else
'      NodeId.ForeColor = BPRED
'   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubNodes(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   vData = objTV.ReadNodeData(NodeId.Parent)
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   strSQL = "SELECT * " & _
            "FROM Clinician_Local_Id " & _
            "WHERE Clinician_National_Code = '" & vData(0) & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Clinician_National_Code, _
                                                 RS!Clinician_Local_Code, _
                                                 "SubHeader", _
                                                 , _
                                                 ms_DELETE), _
                                RS!Clinician_Local_Code, _
                                plist("SUBICON").Icon, _
                                plist("SUBICON").Icon)
'      tv1.Nodes.Add nd(0), _
'                    tvwChild, _
'                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
'                    "Please wait...", _
'                    plist("SUBICON").Icon, _
'                    plist("SUBICON").Icon
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
   eClass.CurrentProcedure = "LoadClinicians.SubNodes"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubHeader(NodeId As Node)
   On Error GoTo procEH
   Dim pos As Integer
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim iconId As Long
   
'  Read owner data and set up control items
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   plist("CNAT").value = vData(0)
'   plist("<Control key>").value = vData(1)
   
'  Read this node details and set up controls
   vData = objTV.ReadNodeData(NodeId)
'   pos = InStr(1, vData(1), "-")
   plist("CLOC").value = vData(1)
   plist("OLDLOC").value = vData(1)
   
'  The sql string to be executed
   strSQL = "SELECT * " & _
            "FROM Clinician_Local_Id " & _
            "WHERE Clinician_National_Code = '" & vData(0) & "' " & _
               "AND Clinician_Local_Code = '" & vData(1) & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   plist("LCODE").value = Mid(RS!Clinician_Local_Code, 7)
   plist.ActivePage = "Loc"
   
   RS.Close
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.SubHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Public Function Refresh() As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = False

'  Specific refresh activities
   
   If objTV.newNode Then
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   If etrap.UpdateDb Then
      objTV.ActiveNode.Text = plist("SURNAME").value
      objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("NCODE").value, _
                                              plist("CSPEC").value, _
                                              "Details", _
                                              , _
                                              ms_DELETE)
'      If plist("ACT").value = False Then
'         objTV.ActiveNode.ForeColor = BPRED
'      Else
'         objTV.ActiveNode.ForeColor = BPRED
'      End If
   End If
   
   objTV.ActiveNode = objTV.RefreshNode
   Set newNode = Nothing
End Function

Public Sub RunWhat(NodeId As Node, _
                   Optional Refresh As Boolean = False)
   On Local Error GoTo procEH
   Dim vData As Variant
   Dim lCode As String
   Dim tNode As MSComctlLib.Node
   Dim RefreshId As String
   Dim strArray() As String
   
'   plist("LCODE").PageKeys = "Loc"
'   plist("LDETS").PageKeys = "Nat"
   If TidyUp = False Then
      tv1.Visible = False
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "FirstView"
            FirstView
         
         Case "Details"
            mCtrl.TreeViewUpdate = False
            objTV.RefreshNode = NodeId
            Details NodeId, Refresh
            
         Case "ShowClinicians"
            ShowClinicians NodeId
         
         Case "SubNodes"
            objTV.RefreshNode = NodeId.Parent
            SubNodes NodeId
            
         Case "Item"
            objTV.RefreshNode = NodeId.Parent
            Details NodeId.Parent, Refresh
            
         Case "SubHeader"
'            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId.Parent
            SubHeader NodeId
'            mCtrl.TreeViewUpdate = False
            
         Case "SubItem"
            objTV.RefreshNode = NodeId.Parent.Parent
            SubHeader NodeId.Parent
         
         Case "ShowKey3"
            frmEDIClinicians.Show 1
            FirstView
            
      End Select
   End If
   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function TidyUp() As Boolean
   TidyUp = False
   If Not newNode Is Nothing Then
      If objTV.ActiveNode.Text <> "Add new Clinician..." Then
         If objTV.nodeKey(newNode) = "NewNode" Then
            tv1.Nodes.Remove newNode.Index
            tv1.SelectedItem = tv1.Nodes(2)
            TidyUp = True
         End If
         Set newNode = Nothing
      End If
   End If
End Function

Public Function Update(PageId As String) As String
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim hpInd As Integer
   
'  Run a stored procedure or some special SQL
   With iceCmd
      .ActiveConnection = iceCon
      .CommandType = adCmdStoredProc
      .CommandText = "ICECONFIG_Amend_Clinician_Name"
      .Parameters.Append .CreateParameter("Name", adVarChar, adParamInput, 35, plist("SURNAME").value)
      .Parameters.Append .CreateParameter("natCode", adVarChar, adParamInput, 8, plist("NCODE").value)
      .Execute
   End With

   Set iceCmd = Nothing
   
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   If iceCmd.Parameters("Return").value <> 0 Then
      eClass.FurtherInfo = iceCmd.Parameters("Stage").value
   End If
   eClass.CurrentProcedure = "LoadClinicians.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function


Private Sub Class_Terminate()
   tv1.Style = tvwTreelinesPlusMinusPictureText
   tv1.Indentation = 200
End Sub
