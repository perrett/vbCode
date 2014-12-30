VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadUsers"
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
      .PropertyItems.Add "PCLASS", "", plpsString, , "Hidden"
      .PropertyItems.Add "PSUBCLASS", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Main", "Service_Permissions"
      .Pages.Add "Sub", "Service_Permissions"
      
'     Set up the required property items
      With .PropertyItems
         .Add "CLASS", "Permission Class", plpsString, , "Mandatory - The Permission class"
         .Add "SUB", "Subclass Details", plpsString, "Click '+' to view", "Subclasses associated with class"
         .Add "SUBCLASS", "Permission Sub-Class", plpsString, , "Mandatory - The Permission sub-class"
         .Add "SRC", "Source Table", plpsList, , "Mandatory - from which table are these items read"
         .Add "SRCDATA", "Source Data column", plpsString, , "Which column to use?"
         .Add "SRCDESC", "Source Description Column", plpsString, , "Which column(s) contain the description"
         .Add "PTYPE", "Permission Type", plpsString, , "Mandatory - The variable type of the permission data"
         .Add "DESC", "Description", plpsString, , "How is this item used"
      End With
   
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 4
      .PropertyItems("SUBICON").PageKeys = "Control"
      .PropertyItems("SUBICON").Icon = 4

'     Set other Control items to the page key
      .PropertyItems("PCLASS").PageKeys = "Control"
      .PropertyItems("PSUBCLASS").PageKeys = "Control"
      
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
      With .PropertyItems("CLASS")
         .Tag = "Permission_Class"
         .DefaultExtension = "Details"
      End With
      With .PropertyItems("SUB")
         .Tag = "MENU"
         .PageKeys = "Main"
         .Flags = 1
         .DefaultExtension = "SubNodes"
      End With
      With .PropertyItems("SUBCLASS")
         .Tag = "Permission_Subclass"
         .PageKeys = "Sub"
      End With
      With .PropertyItems("SRC")
         .Tag = "Source_Table"
         .DefaultExtension = "SubHeader"
         .PageKeys = "Sub"
      End With
      With .PropertyItems("SRCDATA")
         .Tag = "Source_Data_Column"
         .DefaultExtension = "SubItem"
         .PageKeys = "Sub"
      End With
      With .PropertyItems("SRCDESC")
         .Tag = "Source_Description_Column"
         .DefaultExtension = "SubItem"
         .PageKeys = "Sub"
      End With
      With .PropertyItems("PTYPE")
         .Tag = "Permission_Type"
         .DefaultExtension = "SubItem"
         .PageKeys = "Sub"
      End With
      With .PropertyItems("DESC")
         .Tag = "Description"
         .DefaultExtension = "SubItem"
         .PageKeys = "Sub"
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
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("New", _
                                              "NewNode", _
                                              "NewMapping", _
                                              , _
                                              , _
                                              True), _
                             "Add new Permission...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT DISTINCT Permission_Class " & _
            "FROM Service_Permissions"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      With plist
         .PropertyItems("CLASS").defaultValue = RS!Permission_Class
      End With
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Permission_Class, _
                                                 "", _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE), _
                                RS!Permission_Class, _
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
   eClass.CurrentProcedure = "Load???.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
      
   strSQL = "DELETE FROM <TableId> " & _
            "WHERE <Condition>"
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
   
   If Refresh Then
      vData = objTV.ReadNodeData(NodeId)
   
'     Set Control values
      plist("PCLASS").value = vData(0)
'      plist("PSUBCLASS").value = vData(1)
'     SQL string to be executed
   End If
   
   objTV.SQL = "SELECT * " & _
                  "FROM Service_Permissions " & _
                  "WHERE Permission_Class = '" & plist("PCLASS").value & "'"
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Main", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Main", mCtrl, NodeId
   End If
   
'   plist.Caption = "<Update condition>"
   
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
   plist.Caption = "WHERE Permission_Class = " & plist("PCLASS").value
   
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
   plist.ActivePage = "Sub"
   
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
            "FROM Service_permissions " & _
            "WHERE Permission_Class = '" & vData(0) & "'" & _
               "AND Permission_SubClass = '" & vData(1) & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Permission_Class, _
                                                 RS!Permission_SubClass, _
                                                 "SubHeader", _
                                                 , _
                                                 ms_DELETE), _
                                RS!Permission_SubClass, _
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
'   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   
'  Read this node details and set up controls
   vData = objTV.ReadNodeData(NodeId)
   pos = InStr(1, vData(1), "-")
   plist("PCLASS").value = vData(0)
   plist("PSUBCLASS").value = vData(1)
   
'  The sql string to be executed
   objTV.SQL = "SELECT * " & _
               "FROM Service_Permissions " & _
               "WHERE Permission_Class = '" & vData(0) & "'"
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Sub", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Sub", mCtrl, NodeId
   End If
   
'  The SQL condition (if not a stored procedure)
   plist.Caption = " WHERE Permission_Class = '" & plist("<Key1>").value & "'" & _
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
   
   If objTV.newNode Then
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   
   objTV.ActiveNode.Text = plist("CLASS").value
   objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("CLASS").value, _
                                           plist("SUBCLASS").value, _
                                           "Details")
   
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
