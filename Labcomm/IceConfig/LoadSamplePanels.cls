VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadSamplePanels"
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
      .PropertyItems.Add "SUBICON", "Second Icon to use", plpsString, , "Hidden"

'     Add control items as required
      .PropertyItems.Add "PANEL", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Main", "Request_Sample_Panels"
      
'     Set up the required property items
      With .PropertyItems
         .Add "PANELID", "Panel Id", plpsString, , "Hidden"
         .Add "PANELNAME", "Panel Name", plpsString, , "Mandatory - The name of the panel"
         .Add "DANGER", "Danger", plpsBoolean, , "Is this sample dangerous?"
         .Add "OPTIONS", "Panel Options", plpsString, "<Click '+' to view>", "Options allowed with this panel"
         .Add "RULES", "Panel Rules", plpsString, "<Click '+' to view>", "Rules attached to thsi panel"
      End With
   
      .PropertyItems("PANEL").PageKeys = "Control"
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 2
      .PropertyItems("SUBICON").PageKeys = "Control"
      .PropertyItems("SUBICON").Icon = 2

'     Set other Control items to the page key
'      .PropertyItems("<Control Id>").PageKeys = "Control"
      
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
      With .PropertyItems("PANELID")
         .Tag = "Sample_Panel_Id"
         .DefaultExtension = "Item"
         .PageKeys = "Main"
         .Visible = False
      End With
      With .PropertyItems("PANELNAME")
         .Tag = "Panel_Name"
         .Flags = 2
         .DefaultExtension = "Item"
         .PageKeys = "Main"
      End With
      With .PropertyItems("DANGER")
         .Tag = "Danger"
         .DefaultExtension = "Item"
         .PageKeys = "Main"
      End With
      With .PropertyItems("OPTIONS")
         .Tag = "MENU"
         .DefaultExtension = "SubNodes"
         .Flags = 1
         .PageKeys = "Main"
      End With
      With .PropertyItems("RULES")
         .Tag = "MENU"
         .DefaultExtension = "SubNodes"
         .Flags = 1
         .PageKeys = "Main"
      End With
   End With
'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSamplePanels.Class.Initialize"
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
                             "Add new Sample Panel...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Request_Sample_Panels"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      With plist
         .PropertyItems("PANEL").defaultValue = RS!Sample_Panel_Id
      End With
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Sample_Panel_Id, _
                                                 RS!Danger, _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE), _
                                RS!Panel_NAme, _
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
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSamplePanels.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
      
   vData = objTV.ReadNodeData(NodeId)
   If NodeId = objTV.TopLevelNode(NodeId) Then
      strSQL = "DELETE FROM Request_Sample_Panels " & _
               "WHERE Sample_Panel_Id = " & vData(0) & "; " & _
               "DELETE FROM Request_Sample_Prompts " & _
               "WHERE Sample_Panel_Id = " & vData(0) & "; " & _
               "DELETE FROM Request_Sample_Panels_Options " & _
               "WHERE Sample_Panel_Id = " & vData(0)
               
   ElseIf NodeId.Parent.Text = "Panel Rules" Then
      strSQL = "DELETE FROM Request_Sample_Prompts " & _
               "WHERE Prompt_Index = " & vData(1) & _
                  " AND Sample_Panel_Id = " & vData(0)
   
   Else
      strSQL = "DELETE FROM Request_Sample_Panels_Options " & _
               "WHERE Option_Id = " & vData(1) & _
                  " AND Sample_Panel_Id = " & vData(0)
   End If
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSamplePanels.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node, _
                   Refresh As Boolean)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
   plist("PANELID").PageKeys = "Control"
'   If Refresh Then
      vData = objTV.ReadNodeData(NodeId)
   
'     Set Control values
      plist("PANEL").value = vData(0)
'     SQL string to be executed
'   End If
   
   objTV.SQL = "SELECT * " & _
               "FROM Request_Sample_Panels " & _
               "WHERE Sample_Panel_Id = " & vData(0)
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Main", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Main", mCtrl, NodeId
   End If
   
   plist.Caption = " WHERE Sample_Panel_Id = " & vData(0)
   
'  Specific node actions eg set active colour
      
'   Set nd(0) = objTV.NodeFromKey("<KeyId>", NodeId)
'   nd(0).EnsureVisible
'   If pList("<KeyId>").value = True Then
'      If pList.PropertyItems("<TopKeyId>").value = "" Then
'         NodeId.ForeColor = BPBLUE
'      Else
'         NodeId.ForeColor = BPGREEN
'      End If
'   Else
'      NodeId.ForeColor = BPRED
'   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSamplePanels.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim i As Integer
   
'   With plist
'      For i = 1 To .PropertyItems.Count
'         .PropertyItems(i).Value = .PropertyItems(i).DefaultValue
'      Next i
'   End With
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode)
   If objTV.ActiveNode.Text = "Panel Rules" Then
      frmRules.PanelIdentifier = vData(0)
      frmRules.Show 1
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(CStr(vData(0)), _
                                                   Mid(frmRules.lvRules.SelectedItem.Key, 5), _
                                                   "SubItem", _
                                                   , _
                                                   , _
                                                   False), _
                                  frmRules.lvRules.SelectedItem.Text, _
                                  2, _
                                  2)
      Unload frmRules
      
   Else
      frmSamplePanel.PanelIdentifier = vData(0)
      frmSamplePanel.Show 1
      
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(CStr(vData(0)), _
                                                   frmSamplePanel.PanelOption, _
                                                   "SubItem", _
                                                   , _
                                                   , _
                                                   False), _
                                  frmSamplePanel.PanelOptionName, _
                                  2, _
                                  2)
      Unload frmSamplePanel
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
   eClass.CurrentProcedure = "IceConfig.LoadSamplePanels.MenuAddEntry"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewMapping()
   On Error GoTo procEH
   Dim i As Integer
   Dim RS As New ADODB.Recordset

   With plist
      .Redraw = False
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
      Next i
      .Redraw = True
   End With
   If newNode Is Nothing Then
      plist("PANELID").PageKeys = "Main"
      strSQL = "SELECT Max(Sample_Panel_Id) " & _
               "FROM Request_Sample_Panels"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      Set newNode = tv1.Nodes.Add(, _
                                  , _
                                  mCtrl.NewNodeKey(CStr(Val(RS(0)) + 1), _
                                                   "NewNode", _
                                                   "Details", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
      plist("PANELID").value = Val(RS(0) + 1)
      RS.Close
   End If
   
   plist.ActivePage = "Main"
   objTV.ActiveNode = newNode
   objTV.RefreshNode = newNode
'   objTView.ActiveNode = newNode
   fView.Show Fra_EDI
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSamplePanels.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubNodes(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim i As Integer
   
   vData = objTV.ReadNodeData(NodeId.Parent)
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   If NodeId.Text = "Panel Rules" Then
      strSQL = "SELECT * " & _
               "FROM Request_Sample_Prompts " & _
               "WHERE Sample_Panel_Id = " & vData(0)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      Do Until RS.EOF
         strSQL = "SELECT Prompt_Index, Prompt_Desc " & _
                  "FROM Request_Prompt " & _
                  "WHERE Prompt_Index = " & RS!Prompt_Index
         RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey(CStr(vData(0)), _
                                                    RS!Prompt_Index, _
                                                    "SubItem", _
                                                    , _
                                                    ms_DELETE), _
                                   RS2!Prompt_Desc, _
                                   plist("SUBICON").Icon, _
                                   plist("SUBICON").Icon)
         RS2.Close
         
         RS.MoveNext
      Loop
      
      RS.Close
      
   Else
      strSQL = "SELECT * " & _
               "FROM Request_Sample_Panels_Options " & _
               "WHERE Sample_Panel_Id = " & vData(0)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      Do Until RS.EOF
         strSQL = "SELECT Description " & _
                  "FROM Request_Sample_Panel_Collection_Day_Options " & _
                  "WHERE Option_Id = " & RS!Option_Id
         RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey(CStr(vData(0)), _
                                                    RS!Option_Id, _
                                                    "SubItem", _
                                                    , _
                                                    ms_DELETE), _
                                   RS2!Description, _
                                   plist("SUBICON").Icon, _
                                   plist("SUBICON").Icon)
         RS2.Close
   '      tv1.Nodes.Add nd(0), _
   '                    tvwChild, _
   '                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
   '                    "Please wait...", _
   '                    plist("SUBICON").Icon, _
   '                    plist("SUBICON").Icon
         RS.MoveNext
      Loop
      RS.Close
   End If
   Set RS2 = Nothing
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSamplePanels.SubNodes"
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
               "FROM <TableId>" & _
               "WHERE <Condition>"
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
   eClass.CurrentProcedure = "LoadSamplePanels.SubHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Public Function Refresh() As String
'   Dim tNode As node
'
'   mCtrl.TreeViewUpdate = True
'
''  Specific refresh activities
'
   objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("PANELID").value, _
                                           plist("DANGER").value, _
                                           "Details", _
                                           , _
                                           ms_DELETE)

   objTV.ActiveNode.Text = plist("PANELNAME").value
'   objTV.ActiveNode.Key = mCtrl.NewNodeKey("<New Value>")
   objTV.RefreshNode = objTV.TopLevelNode(objTV.ActiveNode)
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
   
   If TidyUp Then
      fView.FrameToShow = Fra_HELP
   Else
      fView.FrameToShow = Fra_EDI
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
            Details NodeId.Parent.Parent, True
         
      End Select
   End If
'   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSamplePanels.RunWhat"
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
   eClass.CurrentProcedure = "LoadSamplePanels.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function
