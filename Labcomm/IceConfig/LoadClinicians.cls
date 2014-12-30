VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadClinicians"
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
Private newIceNum As Long

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim i As Integer
   
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   mCtrl.TreeViewUpdate = True
   
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"
      .PropertyItems.Add "SUBICON", "Second Icon to use", plpsString, , "Hidden"

'     Add control items as required
      .PropertyItems.Add "CNAT", "", plpsString, , "Hidden"
      .PropertyItems.Add "CLOC", "", plpsString, , "Hidden"
      .PropertyItems.Add "OLDLOC", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Nat", "Clinician"
      .Pages.Add "Loc", "Clinician_Local_ID"
      
'     Set up the required property items
      With .PropertyItems
         .Add "NAT", "National Code", plpsString, , "If no National Code is supplied, the system will allocate the next available ICE number"
         .Add "LDETS", "Local details", plpsString, "<Click to expand>", "Local Details"
         .Add "LCODE", "Local Code", plpsString, , "Mandatory - The Local code for this clinician"
         .Add "FLOC", "Full Local Code", plpsString, , "Hidden"
'         .Add "ROLE", "Role of Clinician", plpsString, , "What is the role of this clinician?"
         .Add "ORG", "Organisation", plpsString, , "This clinician's organisation (Read only)"
         .Add "SURNAME", "Surname", plpsString, , "Mandatory - This Clinician's surname"
         .Add "FORENAME", "Forename", plpsString, , "This Clinician's forname"
         .Add "MIDNAME", "Middlename", plpsString, , "This Clinician's middlename"
         .Add "TITLE", "Title", plpsString, , "This Clinician's Title"
'         .Add "GRADE", "Grade", plpsString, , "This Clinician's grade"
         .Add "ACT", "Active", plpsBoolean, , "Is this clinician active?"
         .Add "CSPEC", "Specialty", plpsCustom, , "Mandatory - The specialty of this clinician"
'         .Add "FUNC", "Specialty function", plpsString, , "The function of this specialty"
'         .Add "LOCSPEC", "Local specialty Code", plpsString, , "The local specialty code"
         .Add "DISCHARGE", "Discharge Workgroup", plpsString, , "The discharge workgroup"
         .Add "NDATE", "Date added", plpsString, , "Hidden"
         .Add "MAIL", "Mail Address", plpsString, , "Mail address"
      End With
      
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 4
      .PropertyItems("SUBICON").PageKeys = "Control"
      .PropertyItems("SUBICON").Icon = 5

'     Set other Control items to the page key
      .PropertyItems("CNAT").PageKeys = "Control"
      .PropertyItems("CLOC").PageKeys = "Control"
      .PropertyItems("OLDLOC").PageKeys = "Control"
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
      
      With .PropertyItems("NAT")
         .Tag = "Clinician_National_Code"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("LDETS")
         .Tag = "MENU"
         .DefaultExtension = "SubNodes"
         .defaultValue = "<Click '+' to expand>"
         .Flags = 1
         .PageKeys = "Nat"
      End With
      With .PropertyItems("LCODE")
'         .Tag = "Clinician_Local_Code"
         .DefaultExtension = "SubNodes"
         .max = 10
      End With
      With .PropertyItems("FLOC")
         .Tag = "Clinician_Local_Code"
         .Visible = False
      End With
'      With .PropertyItems("ROLE")
'         .Tag = "Role_of_Clinician"
'         .DefaultExtension = "Item"
'      End With
      With .PropertyItems("ORG")
         .Tag = "Clinician_Org_ID"
         .DefaultExtension = "Item"
         .defaultValue = Left(frmMain.OrgList.Text, 6)
         .ReadOnly = True
      End With
      With .PropertyItems("SURNAME")
         .Tag = "Clinician_Surname"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("FORENAME")
         .Tag = "Clinician_Forename"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("MIDNAME")
         .Tag = "Clinician_Middlenames"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("TITLE")
         .Tag = "Clinician_Title"
         .DefaultExtension = "Item"
      End With
'      With .PropertyItems("GRADE")
'         .Tag = "Clinician_Grade"
'         .DefaultExtension = "Item"
'      End With
      With .PropertyItems("ACT")
         .Tag = "Clinician_Active"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("CSPEC")
         .Tag = "Clinician_Speciality_Code"
         .DefaultExtension = "Item"
         .defaultValue = "999"
      End With
'      With .PropertyItems("FUNC")
'         .Tag = "Speciality_Function_Code"
'         .DefaultExtension = "Item"
'      End With
'      With .PropertyItems("LOCSPEC")
'         .Tag = "Local_Speciality_Code"
'         .DefaultExtension = "Item"
'      End With
      With .PropertyItems("DISCHARGE")
         .Tag = "Discharge_Workgroup"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("NDATE")
         .Tag = "Date_Added"
         .DefaultExtension = "Item"
         .Visible = False
         .PageKeys = "Nat"
      End With
'      With .PropertyItems("LDATE")
'         .Tag = "Date_Added"
'         .DefaultExtension = "SubItem"
'         .Visible = False
'         .PageKeys = "Loc"
'      End With
      With .PropertyItems("MAIL")
         .Tag = "Mail_Address"
         .DefaultExtension = "Item"
      End With
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
   
   plist.Pages("Nat").Caption = "ICECONFIG_UpdateClinicians"
   plist.Pages("Loc").Caption = "ICECONFIG_UpdateClinicians"
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.Class.Initialize"
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
                                              "NewMapping", _
                                              , _
                                              , _
                                              True), _
                             "Add new Clinician...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Clinician " & _
            "ORDER BY Clinician_Surname"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
'      With plist
'         .PropertyItems("<Key>").DefaultValue = "<Default>"
'      End With
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Clinician_National_Code, _
                                                 Trim(RS!Clinician_Speciality_Code & ""), _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE), _
                                Trim(RS!Clinician_Surname & "") & " (" & RS!Clinician_National_Code & ")", _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
      If RS!Clinician_Active Then
         nd(0).ForeColor = BPGREEN
      Else
         nd(0).ForeColor = BPRED
      End If
            
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    2, _
                    2
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

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim tNode As Node
      
   strSQL = ""
   vData = objTV.ReadNodeData(NodeId)
   If vData(2) = "Details" Then
      strSQL = "SELECT Count(Service_HP_Index) AS RecsUsing " & _
               "FROM Service_Health_Parties " & _
               "WHERE (Service_HP_Type = '902' OR Service_HP_Type = '906')" & _
                  "AND Service_HP_Nat_Code = '" & vData(0) & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS!RecsUsing > 0 Then
         RS.Close
         Set RS = Nothing
         MsgBox "This clinician is referenced in Health Parties and cannot be deleted", vbOKOnly, "Unable to delete " & NodeId.Text
         Exit Sub
      Else
         RS.Close
         strSQL = "SELECT Count(EDI_HP_Index) AS RecsUsing " & _
                  "FROM EDI_Health_Parties " & _
                  "WHERE (EDI_HP_Type = '902' OR EDI_HP_Type = '906')" & _
                     "AND EDI_HP_Nat_Code = '" & vData(0) & "'"
         
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         If RS!RecsUsing > 0 Then
            MsgBox "This clinician is referenced in Health Parties and cannot be deleted", vbOKOnly, "Unable to delete " & NodeId.Text
            RS.Close
            Set RS = Nothing
            Exit Sub
         Else
            strSQL = "DELETE FROM Clinician_Local_Id " & _
                     "WHERE Clinician_National_Code = '" & vData(0) & "'; " & vbCrLf & _
                     "DELETE FROM Clinician " & _
                     "WHERE Clinician_National_Code = '" & vData(0) & "'"
         End If
      End If
   
   Else
      strSQL = "SELECT Count(Clinician_Local_Code) AS RecsPresent " & _
               "FROM Clinician_Local_Id " & _
               "WHERE Clinician_National_Code = '" & vData(0) & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      If RS!RecsPresent = 1 Then
         MsgBox "A clinician must have at least one Local Code.", vbExclamation, "Unable to delete " & vData(1)
         RS.Close
         Set RS = Nothing
         Exit Sub
      Else
         strSQL = "DELETE FROM Clinician_Local_Id " & _
                  "WHERE Clinician_Local_Code = '" & vData(1) & "'"
      End If
   End If
   RS.Close
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.Delete"
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
      vData = objTV.ReadNodeData(NodeId)
   
'     Set Control values
      plist("CNAT").value = vData(0)
'     SQL string to be executed
'   End If
   
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
   
   plist.Caption = "WHERE Clinician_National_Code = '" & vData(0) & "'"
   
'  Specific node actions eg set active colour
      
'   Set nd(0) = objTV.NodeFromKey("<KeyId>", NodeId)
'   nd(0).EnsureVisible
   If plist("ACT").value = True Then
      NodeId.ForeColor = BPGREEN
   Else
      NodeId.ForeColor = BPRED
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim i As Integer
   
   plist("LCODE").value = ""
'      For i = 1 To .PropertyItems.Count
'         .PropertyItems(i).value = .PropertyItems(i).DefaultValue
'      Next i
'   End With
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode)
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(CStr(vData(0)), _
                                                   "NewNode", _
                                                   "SubNodes", _
                                                   , _
                                                   , _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
   End If
   
   objTV.RefreshNode = objTV.ActiveNode
   objTV.ActiveNode = newNode
   plist.ActivePage = "Loc"
   
'  Specific preparations for a new item
'   <Currently none>
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadClinicians.MenuAddEntry"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewMapping()
   On Error GoTo procEH
   Dim i As Integer
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   
   plist("LCODE").PageKeys = "Nat"
   plist("LDETS").PageKeys = "Loc"
   
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
                                                   "Details", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
   End If
   
   plist.ActivePage = "Nat"
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
   eClass.CurrentProcedure = "LoadClinicians.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub PB_AfterEdit(prop As PropertyItem, nVal As Variant)
   Dim RS As New ADODB.Recordset
   Dim tNode As Node
   
   nVal = UCase(nVal)
   If prop.Key = "NAT" Then
      strSQL = "SELECT * " & _
               "FROM Clinician " & _
               "WHERE clinician_National_Code = '" & nVal & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      If RS.EOF Then
         If Left(nVal, 3) = "ICE" Then
            MsgBox "The system will allocate the next available 'ICE' number.", vbInformation, "Invalid National Code"
            nVal = ""
         End If
      Else
         Set tNode = objTV.NodeFromText(RS!Clinician_Surname & " (" & nVal & ")")
         Set tv1.SelectedItem = tNode
         
         If Not (newNode Is Nothing) Then
            plist("LCODE").PageKeys = "Loc"
            plist("LDETS").PageKeys = "Nat"
            tv1.Nodes.Remove objTV.ActiveNode.Index
            Set newNode = Nothing
         End If
         
         objTV.ActiveNode = tNode
         objTV.RefreshNode = tNode
         mCtrl.TreeViewUpdate = False
         Details tNode, True
         mCtrl.TreeViewUpdate = True
         plist.Refresh
      End If
      
   End If
End Sub

Public Sub PB_Click(prop As PropertyItem)
   frmSpecs.CalledFrom = "TESTSPEC"
   frmSpecs.EDI_WriteBack = "CSPEC"
   frmSpecs.Show 1
   frmSpecs.EDI_WriteBack = ""
   frmSpecs.CalledFrom = ""
   Unload frmSpecs
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
   
   mCtrl.TreeViewUpdate = True

'  Specific refresh activities
   
   If objTV.newNode Then
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   If etrap.UpdateDb Then
      objTV.ActiveNode.Text = plist("SURNAME").value & " (" & plist("NAT").value & ")"
      objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("NAT").value, _
                                              plist("CSPEC").value, _
                                              "Details", _
                                              , _
                                              ms_DELETE)
      If plist("ACT").value = False Then
         objTV.ActiveNode.ForeColor = BPRED
      Else
         objTV.ActiveNode.ForeColor = BPRED
      End If
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
   
   plist("LCODE").PageKeys = "Loc"
   plist("LDETS").PageKeys = "Nat"
   If TidyUp = False Then
'      tv1.Visible = False
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "FirstView"
            FirstView NodeId
         
         Case "Details"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId
            Details NodeId, Refresh
            mCtrl.TreeViewUpdate = False
         
         Case "NewMapping"
            NewMapping
         
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
         
      End Select
   End If
'   tv1.Visible = True
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
   Dim RS As New ADODB.Recordset
   Dim newValue As String
   Dim intRet As Long
   Dim failInfo As String
   Dim strArray() As String
   
'  Run a stored procedure or some special SQL
   If PageId = "Nat" Then
      If plist("NAT").value = "" Then
'         ICECon.BeginTrans
         With iceCmd
            .ActiveConnection = iceCon
            .CommandType = adCmdStoredProc
            .CommandText = "ICEIMP_Ice_Nums"
            Set RS = .Execute
         End With
         newIceNum = (Val(RS!ICE_Clinician_Num) + 1)
         
         With iceCmd
            .CommandText = "ICEIMP_Ice_Nums_Update"
            .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , 1)
            .Parameters.Append .CreateParameter("ClinNum", adInteger, adParamInput, , newIceNum)
            .Parameters.Append .CreateParameter("GPNum", adInteger, adParamInput, , RS!ICE_GP_Num)
            .Parameters.Append .CreateParameter("RQNum", adInteger, adParamInput, , RS!ICE_Request_Num)
            .Parameters.Append .CreateParameter("PatNum", adInteger, adParamInput, , RS!ICE_Patient_Num)
            .Execute
         End With
'         ICECon.CommitTrans
      
         RS.Close
      
         Set RS = Nothing
         Set iceCmd = Nothing
         Set iceCmd = New ADODB.Command
         plist("NAT").value = "ICE" & newIceNum
      End If
      
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICECONFIG_Amend_Clinician"
         .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
         .Parameters.Append .CreateParameter("New", adBoolean, adParamInput, , objTV.newNode)
         .Parameters.Append .CreateParameter("cNat", adVarChar, adParamInput, 8, UCase(plist("NAT").value))
         .Parameters.Append .CreateParameter("oldNat", adVarChar, adParamInput, 8, plist("CNAT").value)
         .Parameters.Append .CreateParameter("Local", adVarChar, adParamInput, 16, plist("ORG").value & " " & plist("LCODE").value)
         .Parameters.Append .CreateParameter("Role", adVarChar, adParamInput, 50, "")
         .Parameters.Append .CreateParameter("Org", adVarChar, adParamInput, 6, Trim(Left(frmMain.OrgList.Text, 6)))
         .Parameters.Append .CreateParameter("Surname", adVarChar, adParamInput, 35, plist("SURNAME").value)
         .Parameters.Append .CreateParameter("Forename", adVarChar, adParamInput, 35, plist("FORENAME").value)
         .Parameters.Append .CreateParameter("Midname", adVarChar, adParamInput, 35, plist("MIDNAME").value)
         .Parameters.Append .CreateParameter("Title", adVarChar, adParamInput, 20, plist("TITLE").value)
         .Parameters.Append .CreateParameter("Grade", adVarChar, adParamInput, 30, "")
         .Parameters.Append .CreateParameter("Act", adBoolean, adParamInput, , plist("Act").value)
         .Parameters.Append .CreateParameter("Spec", adSmallInt, adParamInput, , Val(plist("CSPEC").value))
         .Parameters.Append .CreateParameter("Func", adSmallInt, adParamInput, , 0)
         .Parameters.Append .CreateParameter("lSpec", adVarChar, adParamInput, 50, "")
         .Parameters.Append .CreateParameter("dWork", adInteger, adParamInput, , Val(plist("DISCHARGE").value))
'         .Parameters.Append .CreateParameter("dAdd", adDate, adParamInput, , IIf(plist("NDATE").value = "", #1/1/1900#, plist("NDATE").value))
         .Parameters.Append .CreateParameter("Mail", adVarChar, adParamInput, 140, Left(plist("MAIL").value, 140))
         
         Dim i As Integer
         For i = 0 To .Parameters.Count - 1
            Debug.Print .Parameters(i).Name & ":" & .Parameters(i).value & ": (" & .Parameters(i).Size & ")"
         Next i
         .Execute
         If .Parameters("Return").value <> 0 Then
            Err.Raise .Parameters("Return").value, "Stored Proc", "Unable to update clinician local id"
         End If
      End With
   
   Else
      plist("FLOC").value = Trim(Left(frmMain.OrgList.Text, 6)) & " " & plist("LCODE").value
      
      If objTV.newNode Then
         strSQL = "SET dateformat dmy;" & _
                  "INSERT INTO Clinician_Local_Id (" & _
                     "Clinician_National_Code, " & _
                     "Clinician_Local_Code, " & _
                     "Date_Added) " & _
                  "VALUES ('" & _
                     plist("CNAT").value & "', " & _
                     "'" & plist("FLOC").value & "', " & _
                     "'" & Format(Now(), "dd/mm/yyyy") & "')"
      Else
         strSQL = "UPDATE Clinician_Local_Id SET " & _
                     "Clinician_Local_Code = '" & plist("FLOC").value & "' " & _
                  "WHERE Clinician_National_Code = '" & plist("CNAT").value & "' " & _
                     "AND Clinician_Local_Code = '" & plist("OLDLOC").value & "'"
         
      End If
      iceCon.Execute strSQL
      plist.ActivePage = "Nat"
   End If
   Set iceCmd = Nothing
   
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadClinicians.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function
