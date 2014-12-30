VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadMySettings"
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
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   mCtrl.TreeViewUpdate = True
   
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"

'     Add control items as required
'      .PropertyItems.Add "LTS", "", plpsString, , "Hidden"
      .PropertyItems.Add "TRADER", "", plpsString, , "Hidden"
      .PropertyItems.Add "FREE", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Main", "EDI_Local_Trader_Settings"
      
'     Set up the required property items
      With .PropertyItems
         .Add "LTS", "LTS Index", plpsNumber, , "Hidden"
         .Add "ORG", "Organisation", plpsString, , "Organisation Code (Same as National Code) - Read Only"
         .Add "NC", "National Code", plpsString, , "Mandatory - The Organisation code for the Hospital"
         .Add "SY", "Data Stream", plpsString, , "The originating Lab System"
         .Add "TC", "Trader Code", plpsString, , "The Trader code for THIS (sending) organisation"
         .Add "FP", "Free Part", plpsString, , "Mandatory - The Free Part for THIS (sending) organisation"
         .Add "PN", "Provider Name", plpsString, , "Sending organisation and data type (Pathology, radiology etc.)"
         .Add "A1", "Address Line 1", plpsString, , "This organisation's address - Line 1"
         .Add "A2", "Address Line 2", plpsString, , "This organisation's address - Line 2"
         .Add "A3", "Address Line 3", plpsString, , "This organisation's address - Line 3"
         .Add "PC", "Post Code", plpsString, , "The Post Code"
         .Add "TN", "Telephone No.", plpsString, , "This organisation's telephone number"
         .Add "DS", "Default Specialty", plpsString, , "The default specialty for this organisation"
         .Add "PK", "PKI Key", plpsString, , "The PKI key for this organisation (Read Only)", "RS!EDI_PKI_Key"
         .Add "AA", "Accept Acknowledgements", plpsBoolean, , "Do we accept acknowledgements?"
         .Add "CF", "Conformance Frequency", plpsNumber, , "After how many messages should a conformance interchange be sent?"
         .Add "CN", "Conformance National Code", plpsCustom, , "Edi reference of conformance authority"
         .Add "CT", "Conformance Running Total", plpsNumber, , "Interchanges sent since last conformance delivery (Read Only)"
      End With
   
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 2

'     Set other Control items to the page key
 '     .PropertyItems("LTS").PageKeys = "Control"
      .PropertyItems("TRADER").PageKeys = "Control"
      .PropertyItems("FREE").PageKeys = "Control"
      
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
      With .PropertyItems("LTS")
         .Tag = "EDI_LTS_Index"
         .Visible = False
      End With
      With .PropertyItems("ORG")
         .Tag = "Organisation"
         .defaultValue = frmMain.cboTrust.Text
         .ReadOnly = True
         .Visible = False
      End With
      With .PropertyItems("NC")
         .DefaultExtension = "Item"
         .Tag = "EDI_OrgCode"
      End With
      With .PropertyItems("SY")
         .DefaultExtension = "Item"
         .Tag = "EDI_Msg_Type"
      End With
      With .PropertyItems("TC")
         .DefaultExtension = "Item"
         .Tag = "EDI_Trader_Code"
      End With
      With .PropertyItems("FP")
         .DefaultExtension = "Item"
         .Tag = "EDI_Free_Part"
'         .DefaultValue = "99999"
      End With
      With .PropertyItems("PN")
         .DefaultExtension = "Item"
         .Tag = "EDI_Name"
      End With
      With .PropertyItems("A1")
         .DefaultExtension = "Item"
         .Tag = "EDI_Address1"
      End With
      With .PropertyItems("A2")
         .DefaultExtension = "Item"
         .Tag = "EDI_Address2"
      End With
      With .PropertyItems("A3")
         .DefaultExtension = "Item"
         .Tag = "EDI_Address3"
      End With
      With .PropertyItems("PC")
         .DefaultExtension = "Item"
         .Tag = "EDI_PostCode"
      End With
      With .PropertyItems("TN")
         .DefaultExtension = "Item"
         .Tag = "EDI_TelNo"
      End With
      With .PropertyItems("DS")
         .DefaultExtension = "Item"
         .Tag = "EDI_Default_Dept"
      End With
      With .PropertyItems("PK")
         .DefaultExtension = "Item"
         .Tag = "EDI_PKI_Key"
         .ReadOnly = True
      End With
      With .PropertyItems("AA")
         .DefaultExtension = "Item"
         .Tag = "EDI_Accept_Acks"
         .defaultValue = True
      End With
      With .PropertyItems("CF")
         .DefaultExtension = "Item"
         .Tag = "Conformance_Frequency"
         .defaultValue = 500
      End With
      With .PropertyItems("CN")
         .DefaultExtension = "Item"
         .Tag = "Conformance_Nat_Code"
      End With
      With .PropertyItems("CT")
         .DefaultExtension = "Item"
         .Tag = "Conformance_Total"
         .ReadOnly = True
      End With
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
   frmMain.OrgList.Visible = False
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMySettings.Class.Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub FirstView(Optional NodeId As Node = Nothing)
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim sampTxt As String
   
   tv1.Nodes.Clear
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("New", _
                                              "NewNode", _
                                              "NewMapping", _
                                              , _
                                              , _
                                              True), _
                             "Add new Local Trader...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM EDI_Local_Trader_Settings " & _
            "WHERE Organisation = '" & frmMain.cboTrust.Text & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   plist("ORG").value = frmMain.CurrentLTSOrg
   Do Until RS.EOF
      With plist
         .PropertyItems("TC").defaultValue = RS!EDI_Trader_Code
         .PropertyItems("A1").defaultValue = RS!EDI_Address1
         .PropertyItems("A2").defaultValue = RS!EDI_Address2
         .PropertyItems("A3").defaultValue = RS!EDI_Address3
         .PropertyItems("PC").defaultValue = RS!EDI_PostCode
         .PropertyItems("TN").defaultValue = RS!EDI_TelNo
      End With
      
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Organisation, _
                                                 RS!EDI_LTS_Index, _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE), _
                                RS!EDI_Name & " - " & RS!EDI_Msg_Type, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
'      If RS!EDI_Op_Active Then
'         nd(0).ForeColor = BPBLUE
'      Else
'         nd(0).ForeColor = BPRED
'      End If
            
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    2, _
                    2
      RS.MoveNext
   Loop
   RS.Close
   
   fView.Show Fra_HELP
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMySettings.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
   Dim RS As New ADODB.Recordset
   
   
   strSQL = "SELECT TOP 1 " & _
               "Case " & _
                  "When s.EDI_LTS_Index is null then 0 " & _
                  "Else s.EDI_LTS_Index " & _
               "End " & _
               "+ " & _
               "Case " & _
                  "When e.EDI_LTS_Index is Null then 0 " & _
                  "Else e.EDI_LTS_Index " & _
               "End As InUse " & _
            "FROM EDI_Local_Trader_Settings l " & _
               "LEFT JOIN Service_ImpExp_Headers s " & _
               "ON l.EDI_LTS_Index = s.EDI_LTS_Index " & _
               "LEFT JOIN EDI_InvTest_Codes e " & _
               "ON l.EDI_LTS_Index = e.EDI_LTS_Index " & _
            "WHERE l.EDI_LTS_Index = 3"
'   strSQL = "SELECT Count(*) AS totRecs " & _
            "FROM Service_ImpExp_Headers " & _
            "WHERE EDI_LTS_Index = " & plist("LTS").value
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   If RS!InUse = 0 Then
      strSQL = "DELETE FROM EDI_Local_Trader_Settings " & _
               "WHERE EDI_LTS_Index = " & plist("LTS").value
      iceCon.Execute strSQL
      tv1.Nodes.Remove NodeId.Index
      frmMain.GetOrganisations
   Else
      MsgBox "Local Trader has reports and/or read codes associated with it. Deletion refused", _
             vbInformation, "Database Consistency Check"
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadMySettings.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node, _
                   Refresh As Boolean)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
   fView.FrameToShow = Fra_EDI
   If Refresh Then
      vData = objTV.ReadNodeData(NodeId)
   
'     Set Control values
'      plist("ORGID").value = vData(0)
      plist("ORG").value = vData(0)
      plist("LTS").value = vData(1)
'      plist("FREE").value = vData(4)
'     SQL string to be executed
   End If
   
   objTV.SQL = "SELECT * " & _
               "FROM EDI_Local_Trader_Settings " & _
               "WHERE EDI_LTS_Index = " & plist("LTS").value
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Main", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Main", mCtrl, NodeId
   End If
   
   plist.Caption = "WHERE EDI_LTS_Index = " & plist("LTS").value
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMySettings.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim i As Integer
   
   With plist
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
      Next i
   End With
   
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
                                  2, _
                                  2)
   End If
   
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
   eClass.CurrentProcedure = "IceConfig.LoadMySettings.MenuAddEntry"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewMapping()
   On Error GoTo procEH
   Dim i As Integer
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   With plist
      .Redraw = False
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
      Next i
      .Redraw = True
   End With

   If newNode Is Nothing Then
      strSQL = "SELECT Max(EDI_LTS_Index) AS topLTS " & _
               "FROM EDI_Local_Trader_Settings"
      RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      
      plist("LTS").value = Val(RS!topLTS & "") + 1
      Set newNode = tv1.Nodes.Add(, _
                                  , _
                                  mCtrl.NewNodeKey("New", _
                                                   plist("LTS").value, _
                                                   "Details", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  2, _
                                  2)
   End If
   
   plist.ActivePage = "Main"
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
   eClass.CurrentProcedure = "LoadMySettings.NewMapping"
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
            "FROM " & _
            "WHERE "
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey("<Lookup_1>", _
                                                 "<Lookup_2>", _
                                                 "SubHeader", _
                                                 , _
                                                 ms_DELETE), _
                                "<Descriptive Text>", _
                                "<0>", _
                                "<0>")
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    "<0>", _
                    "<0>"
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
   eClass.CurrentProcedure = "LoadMySettings.SubNodes"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubHeader(NodeId As Node)
   On Error GoTo procEH
   Dim pos As Integer
   Dim i As Integer
   Dim iconId As Long
   
'  Read owner data and set up control items
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   plist("<Control key>").value = vData(0)
   plist("<Control key>").value = vData(1)
   
'  Read this node details and set up controls
   vData = objTV.ReadNodeData(NodeId)
   pos = InStr(1, vData(1), "-")
   plist("<Control key>").value = vData(0)
   
'  The sql string to be executed
   objTV.SQL = "SELECT * " & _
               "FROM " & _
               "WHERE "
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "<Page Id>", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "<Page Id>", mCtrl, NodeId
   End If
   
'  The SQL condition (if not a stored procedure)
   plist.Caption = " WHERE  = '" & plist("<Key1>").value & "' " & _
                     "AND = '" & plist("Key2>").value & "'"

   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMySettings.SubHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Public Function Refresh() As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = True

'  Specific refresh activities
   
   If objTV.newNode Then
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   
   plist("TRADER").value = plist("TC").value
   plist("FREE").value = plist("FP").value
   frmMain.GetOrganisations
   
   FirstView
   objTV.RefreshNode = tv1.Nodes(2)
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
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId.Parent
            SubHeader NodeId
            mCtrl.TreeViewUpdate = False
            
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
   eClass.CurrentProcedure = "LoadMySettings.RunWhat"
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
   Dim pl As PropertiesList
   Dim iceCmd As New ADODB.Command
   Dim newValue As String
   Dim intRet As Long
   Dim failInfo As String
   Dim strArray() As String
   
'  Run a stored procedure or some special SQL
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadMySettings.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function



