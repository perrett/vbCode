VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadLocations"
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
      .PropertyItems.Add "LOCALATION", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Main", "Location"
      
'     Set up the required property items
      With .PropertyItems
         .Add "LOCCODE", "Localation Code", plpsString, , "Hidden"
         .Add "LOCAL", "Local Location Code", plpsString, , "Mandatory - The name for this location"
         .Add "QUAL", "Location Qualifier", plpsString, , "Any further qualifiers?"
         .Add "CLINIC", "Clinic Name", plpsString, , "The Clinic Name for this location"
         .Add "PURPOSE", "Purpose", plpsString, , "What is the purpose of this Location"
         .Add "AGE", "Intended age group", plpsString, , "The intended age group for the location"
         .Add "INTENSIVE", "Intensive Care Intensity", plpsString, , "Intensive care rating"
         .Add "SEX", "Patient Sex", plpsString, , "The sex of the patients in this location"
         .Add "NIGHT", "Night availability", plpsString, , "Night availability"
         .Add "DAY", "Day availability", plpsString, , "Day availability"
         .Add "ACTIVE", "Active", plpsBoolean, , "is this location active?"
         .Add "ADDED", "Date Added", plpsDate, , "Hidden"
         .Add "LOCNATCODE", "Location National Code", plpsString, , "The national code for this location"
         .Add "MAIL", "Mail active", plpsBoolean, , "Is mail active for this location"
      End With
   
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 22
      .PropertyItems("SUBICON").PageKeys = "Control"
      .PropertyItems("SUBICON").Icon = "<IconId>"

'     Set other Control items to the page key
      .PropertyItems("LOCALATION").PageKeys = "Control"
      
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

      With .PropertyItems("LOCCODE")
         .Tag = "Localation_Code"
         .Visible = False
         .defaultValue = ""
      End With
      With .PropertyItems("LOCAL")
         .Tag = "Local_Location_Code"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("QUAL")
         .Tag = "Location_Qualifier"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("CLINIC")
         .Tag = "Clinic_Name"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("PURPOSE")
         .Tag = "Clinic_Purpose"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("AGE")
         .Tag = "Age_Grp_Intended"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("INTENSIVE")
         .Tag = "Int_Care_Intensity"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("SEX")
         .Tag = "Sex_ofPatients"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("NIGHT")
         .Tag = "Night_Availability"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("DAY")
         .Tag = "Day_Availability"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("ACTIVE")
         .Tag = "Location_Active"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("ADDED")
         .Tag = "Date_Added"
         .Visible = False
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("LOCNATCODE")
         .Tag = "Location_Nat_Code"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("MAIL")
         .Tag = "Mail_Active"
         .DefaultExtension = "Item"
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
                             "Add new Location...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Location " & _
            "WHERE rtrim(Left(Localation_Code,6)) = '" & Trim(Left(frmMain.OrgList.Text, 6)) & "' " & _
            "ORDER BY Local_Location_Code"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      With plist
         .PropertyItems("LOCCODE").defaultValue = RS!Localation_Code
      End With
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Localation_Code, _
                                                 Trim(RS!Clinic_Name & ""), _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE), _
                                RS!Local_Location_Code & " - " & RS!Clinic_Name, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
      If RS!Location_Active Then
         nd(0).ForeColor = BPGREEN
      Else
         nd(0).ForeColor = BPRED
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
   eClass.CurrentProcedure = "Load???.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
      
   strSQL = "DELETE FROM Location " & _
            "WHERE Localation_Code = '" & objTV.NodeLevel(objTV.ActiveNode) & "'"
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
      plist("LOCCODE").value = vData(0)
'     SQL string to be executed
   End If
   
   objTV.SQL = "SELECT * " & _
               "FROM Location " & _
               "WHERE Localation_Code = '" & plist("LOCCODE").value & "'"
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Main", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Main", mCtrl, NodeId
   End If
   
   plist.Caption = "WHERE Localation_Code = '" & plist("LOCCODE").value & "'"
   
'  Specific node actions eg set active colour
      
'   Set nd(0) = objTV.NodeFromKey("<KeyId>", NodeId)
'   nd(0).EnsureVisible
   If plist("ACTIVE").value = True Then
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
   
   plist("ADDED").value = Format(Now(), "yyyy/mm/dd hh:nn:ss")
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
   
   If objTV.newNode Then
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   
   Set tNode = objTV.TopLevelNode
   tNode.Text = plist("LOCAL").value & "-" & plist("CLINIC").value
   tNode.Key = mCtrl.NewNodeKey(plist("LOCCODE").value, _
                                plist("CLINIC").value, _
                                "Details", _
                                , _
                                ms_DELETE)
   
   If plist("ACTIVE").value = True Then
      tNode.ForeColor = BPGREEN
   Else
      tNode.ForeColor = BPRED
   End If
   objTV.ActiveNode = tNode
   objTV.RefreshNode = tNode
   
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
