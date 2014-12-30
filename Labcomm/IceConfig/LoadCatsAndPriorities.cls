VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadCatsAndPriorities"
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
Private OLESourceNode As Node
Private OLERelative As Long
Private OLEMoveKey As String
Private OLEMoveText As String

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
      .PropertyItems.Add "DBCODE", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Cat", "Service_Requests_Categories"
      .Pages.Add "Pri", "Service_Requests_Priorities"
      
'     Set up the required property items
      With .PropertyItems
         .Add "P_CODE", "Priority Code", plpsString, , "Mandatory - The Priority code"
         .Add "P_DESC", "Priority Description", plpsString, , "Mandatory - The description"
         .Add "P_ORDER", "Priority Order", plpsString, , "Hidden"
         .Add "P_ORG", "Organisation", plpsString, , "Hidden"
         .Add "C_CODE", "Category Code", plpsString, , "Mandatory - The property Code"
         .Add "C_EDI", "EDI Category", plpsString, , "The EDI Category description"
         .Add "C_ORDER", "Category priority", plpsString, , "Hidden"
      End With
   
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 2
      .PropertyItems("SUBICON").PageKeys = "Control"
      .PropertyItems("SUBICON").Icon = "<IconId>"

'     Set other Control items to the page key
      .PropertyItems("DBCODE").PageKeys = "Control"
      
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
      With .PropertyItems("P_CODE")
         .Tag = "Priority_Code"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("P_DESC")
         .Tag = "Priority_Description"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("P_ORDER")
         .Tag = "Screen_Order"
         .Visible = False
      End With
      With .PropertyItems("P_ORG")
         .Tag = "Org_Code"
         .defaultValue = frmMain.cboTrust.Text
         .Visible = False
      End With
      With .PropertyItems("C_CODE")
         .Tag = "Category_Code"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("C_EDI")
         .Tag = "Category_EDI"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("C_ORDER")
         .Tag = "Priority"
         .Visible = False
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
   eClass.CurrentProcedure = "Load???.Class.Initialize"
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
                             mCtrl.NewNodeKey("0", _
                                              "CAT", _
                                              "None", _
                                              , _
                                              ms_ADD), _
                             "Categories", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   
'   Set nd(1) = tv1.Nodes.Add(nd(0), _
'                             tvwChild, _
'                             mCtrl.NewNodeKey("New", _
'                                              "NewNode", _
'                                              "NewMapping", _
'                                              , _
'                                              , _
'                                              True), _
'                             "Add new Category...", _
'                             plist("ICON").Icon, _
'                             plist("ICON").Icon)
'   nd(1).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Service_Requests_Categories " & _
            "ORDER BY Priority"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
'      With plist
'         .PropertyItems("<Key>").DefaultValue = "<Default>"
'      End With
      Set nd(1) = tv1.Nodes.Add(nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Category_Code, _
                                                 RS!Category_EDI, _
                                                 "Details", _
                                                 "DRAG", _
                                                 ms_DELETE), _
                                RS!Category_Code, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
'      If RS!EDI_Op_Active Then
'         nd(0).ForeColor = BPBLUE
'      Else
'         nd(0).ForeColor = BPRED
'      End If
            
      tv1.Nodes.Add nd(1), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    2, _
                    2
      RS.MoveNext
   Loop
   RS.Close
   
'  Priorites
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("0", _
                                              "PRI", _
                                              "Priorities", _
                                              , _
                                              ms_ADD), _
                             "Priorities", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   
'   Set nd(1) = tv1.Nodes.Add(nd(0), _
'                             tvwChild, _
'                             mCtrl.NewNodeKey("New", _
'                                              "NewNode", _
'                                              "NewMapping", _
'                                              , _
'                                              , _
'                                              True), _
'                             "Add new Priority...", _
'                             plist("ICON").Icon, _
'                             plist("ICON").Icon)
'   nd(1).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Service_Requests_Priorities " & _
            "ORDER BY Screen_Order"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
'      With plist
'         .PropertyItems("<Key>").DefaultValue = "<Default>"
'      End With
      Set nd(1) = tv1.Nodes.Add(nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Priority_Code, _
                                                 RS!Priority_Description, _
                                                 "Details", _
                                                 "DRAG", _
                                                 ms_DELETE), _
                                RS!Priority_Code & " (" & RS!Priority_Description & ")", _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
'      If RS!EDI_Op_Active Then
'         nd(0).ForeColor = BPBLUE
'      Else
'         nd(0).ForeColor = BPRED
'      End If
            
      tv1.Nodes.Add nd(1), _
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
   eClass.CurrentProcedure = "Load???.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
      
   vData = objTV.ReadNodeData(NodeId)
   
   If objTV.nodeKey(objTV.TopLevelNode) = "CAT" Then
      strSQL = "DELETE FROM Service_Requests_Categories " & _
               "WHERE Category_Code = '" & vData(0) & "'"
   Else
      strSQL = "DELETE FROM Service_Requests_Priorities " & _
               "WHERE Priority_Code = '" & vData(0) & "'"
   End If
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Load???.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node, _
                   Refresh As Boolean)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   Dim PageId As String
   
   If Refresh Then
      vData = objTV.ReadNodeData(NodeId)
   
'     Set Control values
      plist("DBCODE").value = vData(0)
'     SQL string to be executed
   End If
   
   If objTV.nodeKey(NodeId.Parent) = "CAT" Then
      objTV.SQL = "SELECT * " & _
                  "FROM Service_Requests_Categories " & _
                  "WHERE Category_Code = '" & vData(0) & "'"
      PageId = "Cat"
   Else
      objTV.SQL = "SELECT * " & _
                  "FROM Service_Requests_Priorities " & _
                  "WHERE Priority_Code = '" & vData(0) & "'"
      PageId = "Pri"
   End If
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, PageId, mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, PageId, mCtrl, NodeId
   End If
   
   If PageId = "Cat" Then
      plist.Caption = "WHERE Category_Code = '" & vData(0) & "'"
   Else
      plist.Caption = "WHERE Priority_Code = '" & vData(0) & "'"
   End If
   
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
   eClass.CurrentProcedure = "Load???.Details"
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
      If vData(1) = "CAT" Then
         plist.ActivePage = "Cat"
      Else
         plist.ActivePage = "Pri"
      End If
      
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey("New", _
                                                   "NewNode", _
                                                   "Details", _
                                                   , _
                                                   , _
                                                   True), _
                                  "New", _
                                  2, _
                                  2)
      tv1.Nodes.Add newNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    2, _
                    2
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
   eClass.CurrentProcedure = "IceConfig.Load???.MenuAddEntry"
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
                                                   "Details", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
   End If
   
   If objTV.nodeKey(objTV.ActiveNode) = "CAT" Then
      plist.ActivePage = "Cat"
   Else
      plist.ActivePage = "Pri"
   End If
   
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
   eClass.CurrentProcedure = "Load???.NewMapping"
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
                                plist("SUBICON").Icon, _
                                plist("SUBICON").Icon)
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    plist("SUBICON").Icon, _
                    plist("SUBICON").Icon
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
   eClass.CurrentProcedure = "Load???.SubNodes"
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
   eClass.CurrentProcedure = "Load???.SubHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Public Function Refresh() As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = True

'  Specific refresh activities
   
'   If objTV.newNode Then
'      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
'   End If
   
   objTV.ActiveNode = objTV.TopLevelNode(objTV.ActiveNode, True)
   
   If objTV.nodeKey(objTV.ActiveNode.Parent) = "CAT" Then
      objTV.ActiveNode.Text = plist("C_CODE").value
      objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("C_CODE").value, _
                                              plist("C_EDI").value, _
                                              "Details")
   Else
      objTV.ActiveNode.Text = plist("P_CODE").value & " (" & plist("P_DESC").value & ")"
      objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("P_CODE").value, _
                                              plist("P_DESC").value, _
                                              "Details")
   End If
   
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
   eClass.CurrentProcedure = "Load???.RunWhat"
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

Public Function TV_DragDrop(NodeId As Node, _
                            Optional data As MSComctlLib.DataObject = Nothing) As Long
   On Error GoTo procEH
   Dim TVEffect As Long
   Dim srcDets As String
   Dim blnFirst As Boolean
   Dim msgStr As String
   
   vData = objTV.ReadNodeData(NodeId)
   srcDets = data.GetData(vbCFText)
   
   OLEMoveKey = objTV.ActiveNode.Key
   OLEMoveText = objTV.ActiveNode.Text
   
   TV_DragDrop = vbDropEffectNone
   
   msgStr = "Are you sure you wish to move '"
   
   Set OLESourceNode = NodeId
   
   If vData(1) = "CAT" Or vData(1) = "PRI" Then
      msgStr = msgStr & objTV.ActiveNode.Text & "' above '" & NodeId.child.FirstSibling.Text & "'?"
      Set OLESourceNode = NodeId.child.FirstSibling
      OLERelative = tvwFirst
   
   Else
      msgStr = msgStr & objTV.ActiveNode.Text & "' below '" & NodeId.Text & "'?"
      OLERelative = tvwNext
   End If
   
   If MsgBox(msgStr, vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
      tv1.Nodes.Remove objTV.ActiveNode.Index
      objTV.ActiveNode = Nothing
      tv1.Nodes.Add OLESourceNode, _
                    OLERelative, _
                    OLEMoveKey, _
                    OLEMoveText, _
                    2
      
      WriteCPSequence objTV.TopLevelNode(NodeId)
   End If
   
   TV_DragDrop = vbDropEffectNone
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadCatsAndPriorities.TV_DragDrop2"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function TV_DragOver(NodeId As Node) As Long
   Dim dhData As Variant
   Dim pData As Variant
   Dim dhp As String
   Dim anp As String
   Dim lngEffect As Long
   
   dhData = objTV.ReadNodeData(NodeId)
   vData = objTV.ReadNodeData(objTV.ActiveNode)
   dhp = objTV.nodeKey(objTV.TopLevelNode(NodeId))
   anp = objTV.nodeKey(objTV.TopLevelNode)
   lngEffect = vbDropEffectNone
   
   If anp = dhp Then
      lngEffect = vbDropEffectMove
   End If
'   If vData(1) = "PANEL" Then
'      If dhData(1) = "PANEL" Then
'         lngEffect = vbDropEffectMove
'      End If
'   ElseIf vData(1) = "PAGE" Then
'      If dhData(0) = vData(0) Then
'         lngEffect = vbDropEffectMove
'      End If
'   End If
   TV_DragOver = lngEffect
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
   eClass.CurrentProcedure = "Load???.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Sub WriteCPSequence(NodeId As Node)
   On Error GoTo procEH
   Dim intSeq As Integer
   Dim tNode As Node
   Dim pNode As Node
   Dim tableId As String
   Dim fieldId As String
   Dim keyId As String
   
   intSeq = 1
   strSQL = ""
   
   If objTV.nodeKey(objTV.TopLevelNode(NodeId)) = "CAT" Then
      tableId = "Service_Requests_Categories"
      fieldId = "Priority"
      keyId = "Category_Code"
   Else
      tableId = "Service_Requests_Priorities"
      fieldId = "Screen_Order"
      keyId = "Priority_Code"
   End If
   
   Set tNode = NodeId.child.FirstSibling
   Do Until tNode Is Nothing
      strSQL = strSQL & _
               "UPDATE " & tableId & " SET " & _
               fieldId & " = " & intSeq & _
               " WHERE " & keyId & " = '" & objTV.NodeLevel(tNode) & "'; "
      intSeq = intSeq + 1

      Set tNode = tNode.Next
   Loop
   iceCon.Execute strSQL
   tv1.Nodes.Clear
   FirstView
   tv1.Visible = True
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadCatsAndPriorities.WriteCPSequence"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
