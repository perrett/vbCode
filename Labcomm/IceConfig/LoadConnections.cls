VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadConnections"
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
   Dim RS As New ADODB.Recordset
   
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   mCtrl.TreeViewUpdate = True
   
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 6

'     Add control items as required
      .PropertyItems.Add "CONINDEX", "Connection Index", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Main", "Connections"
      
'     Set up the required property items
      With .PropertyItems
         .Add "ORG", "Organisation", plpsString, frmMain.cboTrust.Text, "Read only - The organisation"
         .Add "INDEX", "Connection Index", plpsString, , "Mandatory - The database Key"
         .Add "NAME", "Connection Name", plpsList, , "Mandatory - Connection Identifier"
         .Add "DIRECT", "Direction", plpsList, , "Inbound or Outbound"
         .Add "COLL", "Data Source", plpsList, , "How is the source data collected?"
         .Add "MASK", "Collection mask", plpsString, , "Path and file mask for source files"
         .Add "FREQ", "Connection Frequency", plpsString, , "Module run frequency"
         .Add "VALID", "Validation Type", plpsList, , "Validation method"
         .Add "ACK", "Acknowlegment method", plpsList, , "Acknowledgement method required"
         .Add "IFM", "In-flight Mapping", plpsString, , "Mandatory - Process Key"
         .Add "OUT", "The output directory", plpsFolder, , "The directory to place any output files"
         .Add "OUTMASK", "Output file structure", plpsString, , "The output file mask"
         .Add "HIST", "History Folder", plpsFolder, , "The history directory filepath"
         .Add "PEND", "Pending Folder", plpsFolder, , "The Pending directory filepath"
         .Add "ERROR", "Error Folder", plpsFolder, , "The error directory filepath"
         .Add "LOG", "Log Folder", plpsFolder, , "The Log directory filepath"
         .Add "SPEC", "Special Filters", plpsString, , "Special filters"
         .Add "ACT", "Active", plpsBoolean, , "Is this connection active?"
      End With
      

'     Set other Control items to the page key
      .PropertyItems("CONINDEX").PageKeys = "Control"
      
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
      With plist("ORG")
         .Tag = "Organisation"
         .defaultValue = frmMain.cboTrust.Text
         .Visible = False
      End With
      With plist("INDEX")
         .Tag = "Connection_Index"
         .DefaultExtension = "Item"
      End With
      With plist("NAME")
         .Tag = "Connection_Name"
         strSQL = "SELECT * " & _
                  "FROM Connect_Modules"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         With .ListItems
            Do Until RS.EOF
               .Add RS!Module_Name, RS!Connection_Name
               RS.MoveNext
            Loop
         End With
         RS.Close
         Set RS = Nothing
         .DefaultExtension = "Item"
      End With
      With plist("DIRECT")
         .Tag = "Connection_Direction"
         With .ListItems
            .Add "Inbound", "I"
            .Add "Outbound", "O"
         End With
         .DefaultExtension = "Item"
      End With
      With plist("COLL")
         .Tag = "Connection_DataSource"
         With .ListItems
            .Add "TCP/IP", "TCPIP"
            .Add "File Transfer", "FTP"
            .Add "Folder", "FOLDER"
            .Add "Printed", "PRINTER"
            .Add "Terminal", "TERMINAL"
         End With
         .defaultValue = "FOLDER"
         .DefaultExtension = "Item"
      End With
      With plist("MASK")
         .Tag = "Connection_CollectHow"
         .DefaultExtension = "Item"
      End With
      With plist("FREQ")
         .Tag = "Connection_Frequency"
         .defaultValue = 10
         .DefaultExtension = "Item"
      End With
      With plist("VALID")
         .Tag = "Connection_ValidationType"
         With .ListItems
            .Add "7", "HL7V2.3"
            .Add "2", "NHS002"
            .Add "3", "NHS003"
            .Add "N", "NONE"
         End With
         .defaultValue = "NONE"
         .DefaultExtension = "Item"
      End With
      With plist("ACK")
         .Tag = "Connection_Acknowledge"
         With .ListItems
            .Add "S", "FILESCRIPT"
            .Add "A", "HL7ACK"
            .Add "N", "NONE"
         End With
         .defaultValue = "NONE"
         .DefaultExtension = "Item"
      End With
      With plist("IFM")
         .Tag = "Connection_InFlightMapping"
         .DefaultExtension = "Item"
      End With
      With plist("OUT")
         .Tag = "Connection_TargetDirectory"
         .DefaultExtension = "Item"
      End With
      With plist("OUTMASK")
         .Tag = "Connection_TargetFileMap"
         .DefaultExtension = "Item"
      End With
      With plist("HIST")
         .Tag = "Connection_HistoryDirs"
         .DefaultExtension = "Item"
      End With
      With plist("PEND")
         .Tag = "Connection_PendingDirs"
         .DefaultExtension = "Item"
      End With
      With plist("ERROR")
         .Tag = "Connection_ErrorDirs"
         .DefaultExtension = "Item"
      End With
      With plist("LOG")
         .Tag = "Connection_LogDirs"
         .DefaultExtension = "Item"
      End With
      With plist("SPEC")
         .Tag = "Connection_Filters"
         .DefaultExtension = "Item"
      End With
      With plist("ACT")
         .Tag = "Connection_Active"
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
   eClass.CurrentProcedure = "LoadConnections.Class.Initialize"
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
                             "Add new connection...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Connections " & _
            "WHERE Organisation like '" & frmMain.cboTrust.Text & "' " & _
            "ORDER BY Connection_Direction"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Connection_Index, _
                                                 RS!Connection_Name, _
                                                 "Detail", _
                                                 , _
                                                 ms_DELETE), _
                                RS!Connection_Index, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
      If RS!Connection_Active Then
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
   eClass.CurrentProcedure = "LoadConnections.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
      
   strSQL = "DELETE FROM Connections " & _
            "WHERE Connection_Index = '" & plist("CONINDEX").value & "'"
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadConnections.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
   vData = objTV.ReadNodeData(NodeId)
   
'  Set Control values
   plist("CONINDEX").value = vData(0)
   
'  SQL string to be executed
   objTV.SQL = "SELECT * " & _
               "FROM Connections " & _
               "WHERE Connection_Index = '" & plist("CONINDEX").value & "'"
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Main", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Main", mCtrl, NodeId
   End If
   
'  Specific node actions eg set active colour
      
   Set nd(0) = objTV.NodeFromKey("ACT", NodeId)
   nd(0).EnsureVisible
   If plist("ACT").value = True Then
      NodeId.ForeColor = BPGREEN
   Else
      NodeId.ForeColor = BPRED
   End If
   
   plist.Caption = "WHERE Connection_Index = '" & plist("CONINDEX").value & "'"
   
   fView.Show Fra_EDI
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadConnections.Details"
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
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
   End If
   
   plist.ActivePage = "Main"
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
   eClass.CurrentProcedure = "IceConfig.LoadConnections.MenuAddEntry"
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
   eClass.CurrentProcedure = "LoadConnections.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function Refresh() As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = True
   plist("CONINDEX").value = plist("INDEX").value
'  Specific refresh activities
   objTV.ActiveNode.Text = plist("INDEX").value
   objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("INDEX").value, _
                                           plist("NAME").value, _
                                           "Detail", _
                                           , _
                                           ms_DELETE, _
                                           False)
   Set newNode = Nothing
End Function

Public Sub RunWhat(NodeId As Node, _
                   Optional Refresh As String = "")
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
         
         Case "Detail"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId
            Details NodeId
            objTV.ActiveNode = objTV.RefreshNode
            mCtrl.TreeViewUpdate = False
         
         Case "NewMapping"
            NewMapping
         
'         Case "SubNodes"
'            objTV.RefreshNode = NodeId.Parent
'            SubNodes NodeId
'
         Case "Item"
            objTV.RefreshNode = NodeId.Parent
            Details NodeId.Parent
            
'         Case "SubHeader"
'            mCtrl.TreeViewUpdate = True
'            objTV.RefreshNode = NodeId.Parent
'            SubHeader NodeId
'            mCtrl.TreeViewUpdate = False
'
'         Case "SubItem"
'            objTV.RefreshNode = NodeId.Parent.Parent
'            SubHeader NodeId.Parent
         
      End Select
   End If
'   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadConnections.RunWhat"
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
   eClass.CurrentProcedure = "LoadConnections.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

