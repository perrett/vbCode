VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadConfiguration"
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
Private iceCmd As New ADODB.Command

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

'     Add control items as required
      .PropertyItems.Add "PROGRAMINDEX", "", plpsString, , "Hidden"
      .PropertyItems.Add "CONFIGINDEX", "", plpsString, , "Hidden"
      .PropertyItems.Add "OVERINDEX", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Prog", "Configuration_Programs"
      .Pages.Add "Config", "Configuration_IDs"
      .Pages.Add "Over", "Configuration_Overrides"
      
'     Set up the required property items
      .PropertyItems.Add "PROGID", "Program Name", plpsString, , "The program to configure"
      .PropertyItems.Add "CFGNAME", "Configuration Option", plpsString, , "The option in this program to configure"
      .PropertyItems.Add "CFGSTYLE", "Variable Type", plpsList, , "The variable type for the default value"
      .PropertyItems.Add "CFGDEF", "Default Value", plpsString, , "The default value"
      .PropertyItems.Add "CFGNOTES", "Notes", plpsString, , "Descriptive text for this configuration option"
      .PropertyItems.Add "OVER", "Overrides", plpsString, "Select the Category and item", "Overrides applicable to this configuration"
      .PropertyItems.Add "OVERSTYLE", "Variable Type", plpsList, , "The type of value to store for this override"
      .PropertyItems.Add "OVERNAME", "Override Description", plpsString, , "The name for this override value"
      .PropertyItems.Add "OVERVALUE", "Value", plpsString, , "The value of the override option"
      .PropertyItems.Add "OVERNOTES", "Notes", plpsString, , "Any descriptive text for this value"
   End With
   
   With plist
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 2

'     Set other Control items to the page key
      .PropertyItems("PROGRAMINDEX").PageKeys = "Control"
      .PropertyItems("CONFIGINDEX").PageKeys = "Control"
      .PropertyItems("OVERINDEX").PageKeys = "Control"
      
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
      With .PropertyItems("PROGID")
         .Tag = "Program_Name"
         .Flags = -1
      End With
      With .PropertyItems("CFGNAME")
         .Tag = "Config_Name"
         .Flags = -1
         .DefaultExtension = "CItem"
      End With
      With .PropertyItems("CFGNOTES")
         .Tag = "Config_Notes"
         .DefaultExtension = "CItem"
      End With
      With .PropertyItems("CFGSTYLE")
         .Tag = "Value_Type"
         With .ListItems
            .Add "Numeric", 5
'            .Add "Date", 7
            .Add "Text", 8
            .Add "True/False", 11
            .Add "Colour", 104
         End With
         .DefaultExtension = "CItem"
         .defaultValue = 8
      End With
      With .PropertyItems("CFGDEF")
         .Tag = "Default_Value"
         .DefaultExtension = "CItem"
      End With
      With .PropertyItems("OVER")
         .Tag = "MENU"
         .PageKeys = "Config"
         .ReadOnly = True
         .DefaultExtension = "SubNodes"
         .Flags = 2
      End With
      With .PropertyItems("OVERSTYLE")
         .Tag = "Override_Type"
         With .ListItems
            .Add "Numeric", 5
'            .Add "Date", 7
            .Add "Text", 8
            .Add "True/False", 11
            .Add "Colour", 104
         End With
         .DefaultExtension = "OItem"
         .defaultValue = 8
      End With
      With .PropertyItems("OVERNAME")
         .Tag = "Override_Description"
         .Flags = 2
         .DefaultExtension = "OItem"
      End With
      With .PropertyItems("OVERVALUE")
         .Tag = "Override_Value"
         .DefaultExtension = "OItem"
      End With
      With .PropertyItems("OVERNOTES")
         .Tag = "Override_Notes"
         .DefaultExtension = "OItem"
      End With
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   plist.Pages("Prog").Caption = "IceConfig_Amend_Configuration"
   plist.Pages("Config").Caption = "IceConfig_Amend_Configuration"
   plist.Pages("Over").Caption = "IceConfig_Amend_Configuration"
   
   With iceCmd
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("TableId", adVarChar, adParamInput, 6) ', pList.Pages(pList.ActivePage).Key)
      .Parameters.Append .CreateParameter("UpdateType", adInteger, adParamInput) ', , uType)
      .Parameters.Append .CreateParameter("ProgIndex", adInteger, adParamInput) ', , Val(pList("PROGRAMINDEX").value))
      .Parameters.Append .CreateParameter("ProgName", adVarChar, adParamInput, 255) ', pList("PROGID").value)
      .Parameters.Append .CreateParameter("OrgId", adVarChar, adParamInput, 6) ', frmMain.OrgList.Text)
      .Parameters.Append .CreateParameter("CfgIndex", adInteger, adParamInput) ', , Val(pList("CONFIGINDEX").value))
      .Parameters.Append .CreateParameter("CfgName", adVarChar, adParamInput, 50) ', pList("CFGNAME").value)
      .Parameters.Append .CreateParameter("OverIndex", adInteger, adParamInput)
      .Parameters.Append .CreateParameter("TypeIndex", adInteger, adParamInput) ', , Val(pList("OVERSTYLE").value))
      .Parameters.Append .CreateParameter("Notes", adVarChar, adParamInput, 255) ', "")
      .Parameters.Append .CreateParameter("VType", adSmallInt, adParamInput) ', , 0)
      .Parameters.Append .CreateParameter("Value", adVarChar, adParamInput, 255) ', "")
      .Parameters.Append .CreateParameter("Desc", adVarChar, adParamInput, 50) ', pList("OVERNAME").value)
   End With
   
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
                                              "NewMapping"), _
                             "Add new program...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Configuration_Programs"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Program_Index, _
                                                 "Prog", _
                                                 "ConfigItems", _
                                                 "PRG", _
                                                 ms_BOTH), _
                                RS!Program_Name, _
                                1, _
                                1)
      
'     Colour code the node?
'      If RS!EDI_Op_Active Then
'         If Trim(RS!EDI_Read_Code & "") = "" Then
'            nd(0).ForeColor = BPBLUE
'         Else
'            nd(0).ForeColor = BPGREEN
'         End If
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
   eClass.CurrentProcedure = "Load???.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub ConfigItems(NodeId As Node)
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   vData = objTV.ReadNodeData(NodeId)
   plist("PROGRAMINDEX").value = vData(0)
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
'   strSQL = "SELECT * " & _
'            "FROM Configuration_Programs " & _
'            "WHERE Program_index = " & vData(0)
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'   objTV.ActiveNode.Text = RS!Program_Name
'   RS.Close
   strSQL = "SELECT * " & _
            "FROM Configuration_IDs " & _
            "WHERE Program_Index = " & plist("PROGRAMINDEX").value
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(CStr(vData(0)), _
                                                 RS!Config_Index, _
                                                 "Detail", _
                                                 "CFG", _
                                                 ms_DELETE), _
                                RS!Config_Name, _
                                2, _
                                2)
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", _
                                     "Tmp"), _
                    "Please wait...", _
                    2, _
                    2
      RS.MoveNext
   Loop
   objTV.SQL = "SELECT * " & _
               "FROM Configuration_Programs " & _
               "WHERE Program_index = " & plist("PROGRAMINDEX").value
   objTV.UpdatePropList plist, "Prog", mCtrl, NodeId
   NodeId.Text = plist("PROGID").value
   RS.Close
   Set RS = Nothing
   plist.Caption = " WHERE Program_Index = " & plist("PROGRAMINDEX").value
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
   Dim pIndex As Long
   Dim cIndex As Long
   
   vData = objTV.ReadNodeData(NodeId)
   If vData(4) = "PRG" Then
      pIndex = vData(0)
   ElseIf vData(4) = "CFG" Then
      pIndex = vData(0)
      cIndex = vData(1)
   Else
      vData = objTV.ReadNodeData(NodeId.Parent)
      pIndex = vData(0)
      cIndex = vData(1)
   End If
   
   objTV.ActiveNode = NodeId.Parent
   With iceCmd
      .ActiveConnection = iceCon
      .CommandText = plist.Pages(plist.ActivePage).Caption
      .Parameters("TableId").value = plist.Pages(plist.ActivePage).Key
      .Parameters("UpdateType").value = 0
      .Parameters("ProgIndex").value = pIndex
      .Parameters("CfgIndex").value = cIndex
      .Parameters("OverIndex").value = Val(plist("OVERINDEX").value)
      .Execute
   End With
   tv1.Nodes.Remove NodeId.Index
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.Load???.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
   vData = objTV.ReadNodeData(NodeId)
   
'  Set Control values
   plist("PROGRAMINDEX").value = vData(0)
   plist("CONFIGINDEX").value = vData(1)
   
   strSQL = "SELECT Value_Type " & _
            "FROM Configuration_IDs " & _
            "WHERE Program_Index = " & plist("PROGRAMINDEX").value & _
               "AND Config_Index = " & vData(1)
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   plist("CFGDEF").Style = RS!Value_Type
   RS.Close
   
'  SQL string to be executed
   strSQL = "SELECT * " & _
            "FROM Configuration_IDs " & _
            "WHERE Program_Index = " & plist("PROGRAMINDEX").value & _
               "AND Config_Index = " & vData(1)
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
   
   objTV.SQL = strSQL
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Config", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Config", mCtrl, NodeId
   End If
   NodeId.Text = plist("CFGNAME").value
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
   
'   fView.Show Fra_EDI
   plist.Caption = " WHERE Program_Index = " & plist("PROGRAMINDEX").value & _
                     " AND Config_Index = " & vData(1)
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
   Dim RS As New ADODB.Recordset
   Dim rWhat As String
   Dim iconId As Integer
   
   With plist
      For i = 1 To .PropertyItems.Count
         If .PropertyItems(i).PageKeys <> "Control" Then
            .PropertyItems(i).value = .PropertyItems(i).defaultValue
         End If
      Next i
   End With
   
   If objTV.ActiveNode = objTV.TopLevelNode Then
      rWhat = "Details"
      plist.ActivePage = "Config"
      iconId = 2
      vData = objTV.ReadNodeData(objTV.TopLevelNode)
   Else
      rWhat = "SubHeader"
      plist.ActivePage = "Over"
      iconId = 4
      vData = objTV.ReadNodeData(objTV.ActiveNode)
      If vData(4) = 2 Then
         plist("OVERNAME").Style = plpsList
         With plist("OVERNAME").ListItems
            .Clear
            strSQL = "SELECT Local_Location_Code, Clinic_Name " & _
                     "FROM Location " & _
                     "ORDER BY Local_Location_Code"
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            Do While Not RS.EOF
               .Add "(" & Trim(RS!Clinic_Name) & ")", Trim(UCase(RS!Local_Location_Code))
               RS.MoveNext
            Loop
            RS.Close
         End With
      Else
         plist("OVERNAME").Style = plpsString
      End If
      
   End If
   
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(CStr(vData(0)), _
                                                   "NewNode", _
                                                   rWhat, _
                                                   CStr(vData(4)), _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  iconId, _
                                  iconId)
   End If
   
   objTV.RefreshNode = objTV.ActiveNode.Parent
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
                                  2, _
                                  2)
   End If
   
   objTV.ActiveNode = newNode
   objTV.RefreshNode = newNode
'   fView.Show Fra_EDI
   plist.ActivePage = "Prog"
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Load???.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub OverList(NodeId As Node)
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   vData = objTV.ReadNodeData(NodeId)
   strSQL = "SELECT * " & _
            "FROM Configuration_Overrides " & _
            "WHERE Config_Type_Index = " & vData(4) & _
               " AND Config_Index = " & vData(1)
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!Override_Index, _
                                                 RS!Override_Type, _
                                                 "SubHeader", _
                                                 CStr(vData(4)), _
                                                 ms_DELETE), _
                                RS!Override_Description, _
                                4, _
                                4)
'      tv1.Nodes.Add nd(0), _
'                    tvwChild, _
'                    mCtrl.NewNodeKey("Tmp", "tmp"), _
'                    "Please wait...", _
'                    4, _
'                    4
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
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
            "FROM Configuration_Types"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(CStr(vData(0)), _
                                                 CStr(vData(1)), _
                                                 "OverList", _
                                                 RS!Config_Type_Index, _
                                                 ms_ADD), _
                                RS!Config_Type_Name, _
                                4, _
                                4)
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    4, _
                    4
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
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim iconId As Long
   
'  Read owner data and set up control items
'   vData = objTV.ReadNodeData(objTV.TopLevelNode(nodeId))
'   pList("PROGRAMINDEX").value = vData(0)
   
'  Read this node details and set up controls
   vData = objTV.ReadNodeData(NodeId)
   plist("OVERVALUE").Style = vData(1)
   plist("OVERINDEX").value = vData(0)
'   pList("CONFIGINDEX").value = vData(1)
   
'  The sql string to be executed
   objTV.SQL = "SELECT * " & _
               "FROM Configuration_Overrides " & _
               "WHERE Override_Index = " & vData(0) ' & _
                  " AND Config_Type_Index = " & vData(4) & _
                  " AND Override_Description = '" & NodeId.Text & "'"
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Over", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Over", mCtrl, NodeId
      If vData(4) = 2 Then
         plist("OVERNAME").Style = plpsList
         With plist("OVERNAME").ListItems
            .Clear
            strSQL = "SELECT Local_Location_Code, Clinic_Name " & _
                     "FROM Location " & _
                     "ORDER BY Local_Location_Code"
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            Do While Not RS.EOF
               .Add "(" & Trim(RS!Clinic_Name) & ")", Trim(UCase(RS!Local_Location_Code))
               RS.MoveNext
            Loop
            RS.Close
         End With
      Else
         plist("OVERNAME").Style = plpsString
      End If
   End If
   
'  The SQL condition (if not a stored procedure)
   plist.Caption = " WHERE  = '" & plist("CONFIGINDEX").value & "' "

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
'      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   
   Set newNode = Nothing
   objTV.RefreshNode = objTV.ActiveNode.Parent
   objTV.ActiveNode = objTV.RefreshNode
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
      eClass.FurtherInfo = vData(2)
      Select Case vData(2)
         Case "FirstView"
            FirstView NodeId
         
         Case "Detail"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId
            Details NodeId
            mCtrl.TreeViewUpdate = False
         
         Case "NewMapping"
            NewMapping
         
         Case "ConfigItems"
            ConfigItems NodeId
         
         Case "SubNodes"
            objTV.RefreshNode = NodeId.Parent
            SubNodes NodeId
            Details NodeId.Parent
            
         Case "CItem"
            objTV.RefreshNode = NodeId.Parent
            Details NodeId.Parent
            
         Case "OverList"
            OverList NodeId
            mCtrl.TreeViewUpdate = False
            Details NodeId.Parent.Parent
         
         Case "SubHeader"
            mCtrl.TreeViewUpdate = False
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
   Dim i As Integer
   Dim pl As PropertiesList
   Dim uType As Long
   Dim intRet As Long
   Dim failInfo As String
   Dim strArray() As String
   
'   If objTV.newNode Then
'      uType = 1
'   Else
'      uType = 2
'   End If
   
   With iceCmd
      .Parameters("TableId").value = plist.Pages(plist.ActivePage).Key
      If objTV.newNode Then
         .Parameters("UpdateType").value = 1
      Else
         .Parameters("UpdateType").value = 2
      End If
      .Parameters("ProgIndex").value = Val(plist("PROGRAMINDEX").value)
      .Parameters("ProgName").value = plist("PROGID").value
      .Parameters("OrgId").value = frmMain.cboTrust.Text
      .Parameters("CfgIndex").value = Val(plist("CONFIGINDEX").value)
      .Parameters("CfgName").value = plist("CFGNAME").value
      .Parameters("TypeIndex").value = Val(plist("OVERSTYLE").value)
      
'      .Parameters.Append .CreateParameter("TableId", adVarChar, adParamInput, 6, pList.Pages(pList.ActivePage).Key)
'      .Parameters.Append .CreateParameter("UpdateType", adInteger, adParamInput, , uType)
'      .Parameters.Append .CreateParameter("ProgIndex", adInteger, adParamInput, , Val(pList("PROGRAMINDEX").value))
'      .Parameters.Append .CreateParameter("ProgName", adVarChar, adParamInput, 255, pList("PROGID").value)
'      .Parameters.Append .CreateParameter("OrgId", adVarChar, adParamInput, 6, frmMain.OrgList.Text)
'      .Parameters.Append .CreateParameter("CfgIndex", adInteger, adParamInput, , Val(pList("CONFIGINDEX").value))
'      .Parameters.Append .CreateParameter("CfgName", adVarChar, adParamInput, 50, pList("CFGNAME").value)
'      .Parameters.Append .CreateParameter("TypeIndex", adInteger, adParamInput, , Val(pList("OVERSTYLE").value))
      
      Select Case plist.Pages(plist.ActivePage).Key
'         Case "Prog"
'            .Parameters.Append .CreateParameter("Notes", adVarChar, adParamInput, 255, "")
'            .Parameters.Append .CreateParameter("VType", adSmallInt, adParamInput, , 0)
'            .Parameters.Append .CreateParameter("Value", adVarChar, adParamInput, 255, "")
         
         Case "Config"
            .Parameters("Notes").value = plist("CFGNOTES").value
            .Parameters("VType").value = Val(plist("CFGSTYLE").value)
            .Parameters("Value").value = plist("CFGDEF").value
'            .Parameters.Append .CreateParameter("Notes", adVarChar, adParamInput, 255, pList("CFGNOTES").value)
'            .Parameters.Append .CreateParameter("VType", adSmallInt, adParamInput, , Val(pList("CFGSTYLE").value))
'            .Parameters.Append .CreateParameter("Value", adVarChar, adParamInput, 255, pList("CFGDEF").value)
         
         Case "Over"
            vData = objTV.ReadNodeData(objTV.ActiveNode)
            .Parameters("OverIndex").value = vData(0)
            .Parameters("TypeIndex").value = vData(4)
            .Parameters("Notes").value = plist("OVERNOTES").value
            .Parameters("VType").value = Val(plist("OVERSTYLE").value)
            .Parameters("Value").value = plist("OVERVALUE").value
            .Parameters("Desc").value = plist("OVERNAME").value
'            .Parameters.Append .CreateParameter("Notes", adVarChar, adParamInput, 255, pList("OVERNOTES").value)
'            .Parameters.Append .CreateParameter("VType", adSmallInt, adParamInput, , Val(pList("OVERSTYLE").value))
'            .Parameters.Append .CreateParameter("Value", adVarChar, adParamInput, 255, pList("OVERVALUE").value)
         
      End Select
'      .Parameters.Append .CreateParameter("Desc", adVarChar, adParamInput, 50, pList("OVERNAME").value)
      .CommandText = plist.Pages(plist.ActivePage).Caption
'      .CommandType = adCmdStoredProc
      .ActiveConnection = iceCon
      For i = 0 To .Parameters.Count - 1
         Debug.Print .Parameters(i).Name & ":" & vbTab & vbTab & .Parameters(i).value & vbTab & .Parameters(i).Type
      Next i
      .Execute
   End With
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Load???.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function



