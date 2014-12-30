VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadNewSpecimenCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private pList As PropertiesList
Private newNode As node
Private strSQL As String
Private vData As Variant
Private mCtrl As New ManageControls
Private nd(2) As node
Private blnReadOnly As Boolean
'Private plSysBut As New PropertiesListCtl.SysButtonTemplate
'Private myBut As PropertiesListCtl.Buttons

Private Sub Class_Initialize()
   On Error GoTo procEH
   Set tv1 = frmMain.TreeView1
   Set pList = frmMain.edipr
   mCtrl.TreeViewUpdate = True
   
   With pList
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"

'     Add control items as required
      .PropertyItems.Add "SAMPDESC", "", plpsString, , "Hidden"
      .PropertyItems.Add "LOCAL_COUNT", "", plpsNumber, -1, "Hidden"
      .PropertyItems.Add "SAMPCODE", "", plpsString, , "Hidden"
      .PropertyItems.Add "ANATCODE", "", plpsString, , "Hidden"
      .PropertyItems.Add "COLLCODE", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Main", "Service_Sample_Codes"
      .Pages.Add "Desc", "None"
      
'     Set up the required property items
      .PropertyItems.Add "SDESC", "Sample Description", plpsString, , "The National sample description"
      .PropertyItems.Add "SCODE", "Sample Details", plpsString, "<Expand to view details>", "The National Type code"
      .PropertyItems.Add "SANAT", "Anatomical Origin", plpsString, "<Expand to view details>", "National Anatomical Origin Code"
      .PropertyItems.Add "SCOLL", "Collection", plpsString, "<Expand to view details>", "The national collection code"
      .PropertyItems.Add "LOCAL", "Local description(s)", plpsString, , "Local sample type"
      .PropertyItems.Add "NATCODE", "National Code", plpsCustom, , "The relevant national code"
      .PropertyItems.Add "NATDESC", "National Description", plpsString, , "The national description"
      .PropertyItems.Add "ACTIVE", "Active", plpsBoolean, , "Is this sample type active?"
   End With
   
   With pList
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 2

'     Set other Control items to the page key
'      .PropertyItems("<Control key>").PageKeys = "Control"
      
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
      
      With .PropertyItems("SAMPDESC")
         .PageKeys = "Control"
         .Visible = False
      End With
      With .PropertyItems("LOCAL_COUNT")
         .PageKeys = "Control"
         .Visible = False
         .value = -1
         .DefaultValue = -1
      End With
      With .PropertyItems("SAMPCODE")
         .PageKeys = "Control"
         .Visible = False
      End With
      With .PropertyItems("ANATCODE")
         .PageKeys = "Control"
         .Visible = False
      End With
      With .PropertyItems("COLLCODE")
         .PageKeys = "Control"
         .Visible = False
      End With
      
      With .PropertyItems("SDESC")
         .Tag = "Sample_Description"
      End With
      With .PropertyItems("LOCAL")
         .Tag = "MENU"
         .PageKeys = "Control"
         .DefaultExtension = "SubItem"
         .Flags = 1
      End With
      With .PropertyItems("SCODE")
         .Tag = "MENU"
         .PageKeys = "Main"
         .DefaultExtension = "SubNodes"
         .ReadOnly = True
      End With
      With .PropertyItems("SANAT")
         .Tag = "MENU"
         .PageKeys = "Main"
         .DefaultExtension = "SubNodes"
         .ReadOnly = True
      End With
      With .PropertyItems("SCOLL")
         .Tag = "MENU"
         .PageKeys = "Main"
         .DefaultExtension = "SubNodes"
         .ReadOnly = True
      End With
      With .PropertyItems("NATCODE")
         .Tag = "None"
         .PageKeys = "Desc"
      End With
      With .PropertyItems("NATDESC")
         .Tag = "None"
         .PageKeys = "Desc"
         .ReadOnly = True
      End With
      With .PropertyItems("ACTIVE")
         .Tag = "Active"
      End With
      
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures pList
   pList.Pages("Main").Caption = "Specific"
'   pList.Pages("Samp").Caption = "Specific"
'   pList.Pages("Anat").Caption = "Specific"
'   pList.Pages("Coll").Caption = "Specific"
   
'   pList.Buttons.Add
'   pList.Buttons(0).Caption = "hi"
'   pList.Buttons(0).Visible = True
'   Set plSysBut.ButtonControl = pList.Buttons(0).ButtonControl
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Load???.Class.Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let EvNoEdit(blnNewValue As Boolean)
   blnReadOnly = blnNewValue
End Property

Public Sub EvaluateRecords(NodeId As node, _
                           tableId As String, _
                           itemValue As String)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim rc As Integer
   
   With pList
      .Redraw = False
      If pList("LOCAL_COUNT").value > -1 Then
         For i = 0 To pList("LOCAL_COUNT").value
            .PropertyItems.Remove "LKEY_" & i
         Next i
      End If
      
      If Not (NodeId Is Nothing) Then
         For i = 1 To NodeId.Children
            tv1.Nodes.Remove NodeId.Child.Index
         Next i

         tv1.Nodes.Add NodeId, _
                       tvwChild, _
                       mCtrl.NewNodeKey(itemValue, _
                                        "TYPE", _
                                        "Item", _
                                        , _
                                        ms_DELETE), _
                       itemValue & " - " & pList("NATDESC").value, _
                       pList("ICON").Icon, _
                       pList("ICON").Icon
      
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey("LOCAL", _
                                                    "LOCAL", _
                                                    "SubHeader", _
                                                    , _
                                                    ms_ADD), _
                                   "Local Description(s)", _
                                   pList("ICON").Icon, _
                                   pList("ICON").Icon)
      End If
      
      strSQL = "SELECT * " & _
               "FROM " & tableId & _
               " WHERE National_Code = '" & itemValue & "'"
      RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'      pList("LOCAL_COUNT").value = RS(0)
'      RS.Close
      i = 0
      Do Until RS.EOF
         .PropertyItems.Add "LKEY_" & i, "Local Description " & i, plpsString, RS!Local_Text, "Local Description"
         .PropertyItems("LKEY_" & i).Tag = "Local_Text"
         .PropertyItems("LKEY_" & i).PageKeys = "Desc"
         .PropertyItems("LKEY_" & i).DefaultValue = RS!Local_Text
         
         If NodeId Is Nothing Then
            .PropertyItems("LKEY_" & i).ReadOnly = blnReadOnly
         Else
            tv1.Nodes.Add nd(0), _
                          tvwChild, _
                          mCtrl.NewNodeKey(itemValue, _
                                           "LDESC", _
                                           "SubItem", _
                                           tableId, _
                                           ms_DELETE), _
                          RS!Local_Text, _
                          pList("ICON").Icon, _
                          pList("ICON").Icon
         End If
         
         i = i + 1
         RS.MoveNext
      Loop
      RS.Close
      pList("LOCAL_COUNT").value = i - 1
'      pList.Caption = " WHERE National_Code = '" & ItemValue & "'"
      .Redraw = True
   End With
   blnReadOnly = False
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSpecimenCodes.EvaluateRecords"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub FirstView(Optional NodeId As node = Nothing)
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
                                              ms_DELETE, _
                                              True), _
                             "Add new Sample Type...", _
                             pList("ICON").Icon, _
                             pList("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT * " & _
            "FROM Service_Sample_Codes " & _
            "ORDER BY Sample_Description"
   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Type_Code, _
                                                 RS!AnatomicalOrigin_Code, _
                                                 "Detail", _
                                                 RS!Collection_Code, _
                                                 ms_DELETE), _
                                RS!Sample_Description, _
                                pList("ICON").Icon, _
                                pList("ICON").Icon)
      
'     Colour code the node?
      If RS!Active Then
'         If Trim(RS!EDI_Read_Code & "") = "" Then
            nd(0).ForeColor = BPBLUE
'         Else
'            nd(0).ForeColor = BPGREEN
'         End If
      Else
         nd(0).ForeColor = BPRED
      End If
      
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    pList("ICON").Icon, _
                    pList("ICON").Icon
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

Public Sub Delete(NodeId As node)
   On Error GoTo procEH
   Dim tNode As node
      
   If NodeId = objTV.TopLevelNode(NodeId) Then
      Set tNode = NodeId.Previous
   Else
      Set tNode = NodeId.Parent
   End If
   
   vData = objTV.ReadNodeData(NodeId)
   
   Select Case vData(1)
      Case "LDESC"
         strSQL = "DELETE FROM " & vData(4) & _
                  " WHERE Local_Text = '" & NodeId.Text & "'"
         
      Case "TYPE"
      
      Case Else
         strSQL = "DELETE FROM Service_Sample_Codes " & _
                  "WHERE Sample_Description = '" & NodeId.Text & "'"
                  
   End Select
   
   ICECon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   objTV.ActiveNode = tNode
   objTV.RefreshNode = tNode
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.Load???.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As node)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
   vData = objTV.ReadNodeData(NodeId)
   
'  Set Control values
   pList("SAMPDESC").value = NodeId.Text
   pList("SAMPCODE").value = vData(0)
   pList("ANATCODE").value = vData(1)
   pList("COLLCODE").value = vData(4)
'   pList("<Control key>").value = vData(1)
   
'  SQL string to be executed
   objTV.SQL = "SELECT * " & _
               "FROM Service_Sample_Codes " & _
               "WHERE Sample_Description = '" & NodeId.Text & "'"
   
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.Child.Index
      Next i
      objTV.UpdateTreeView pList, "Main", mCtrl, NodeId
   Else
      objTV.UpdatePropList pList, "Main", mCtrl, NodeId
   End If
   
'  Specific node actions eg set active colour
      
'   Set nd(0) = objTV.NodeFromKey("<KeyId>", NodeId)
'   nd(0).EnsureVisible
   If pList("ACTIVE").value = True Then
'      If pList.PropertyItems("<TopKeyId>").value = "" Then
         NodeId.ForeColor = BPBLUE
'      Else
'         NodeId.ForeColor = BPGREEN
'      End If
   Else
      NodeId.ForeColor = BPRED
   End If
   
'   pList("SCODE").value = vData(0)
'   pList("SANAT").value = vData(1)
'   pList("SCOLL").value = vData(4)
   
'   pList.Caption = " WHERE Sample_Description = '" & pList("SDESC").value & "' "
   pList.Caption = "Main"
   fView.Show Fra_EDI
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Load???.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get KeyControl() As ManageControls
   Set KeyControl = mCtrl
End Property

Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim i As Integer
   
   With pList
'     Preserve the value of 'LOCAL_COUNT' and replace after initialisation
'      For i = 1 To .PropertyItems.Count
'         If .PropertyItems(i).PageKeys <> "Control" Then
'            .PropertyItems(i).value = .PropertyItems(i).DefaultValue
'         End If
'      Next i
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
   
'   objTV.RefreshNode = objTV.ActiveNode
'   objTV.ActiveNode = newNode
   
'  Specific preparations for a new item
   pList("LOCAL_COUNT").value = pList("LOCAL_COUNT").value + 1
   pList.PropertyItems.Add "LKEY_" & pList("LOCAL_COUNT").value, _
                           "Local Description " & pList("LOCAL_COUNT").value, _
                           plpsString, _
                           , _
                           "A local description"
   pList("LKEY_" & pList("LOCAL_COUNT")).Selected = True
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

   With pList
      .Redraw = False
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).DefaultValue
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
   objTView.ActiveNode = newNode
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

Public Sub SubNodes(NodeId As node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim tableId As String
   Dim itemValue As String
   Dim pageId As String
   Dim tNode As node
   
   If mCtrl.TreeViewUpdate Then
      Set tNode = NodeId
   Else
      Set tNode = Nothing
   End If
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   objTV.RefreshNode = NodeId
   
'   For i = 1 To nodeId.Children
'      tv1.Nodes.Remove nodeId.Child.Index
'   Next i
'
   Select Case NodeId.Text
      Case "Sample Details"
         strSQL = "SELECT * " & _
                  "FROM CRIR_Sample_Type " & _
                  "WHERE Sample_Code = '" & vData(0) & "'"
         RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
         If RS.EOF Then
            With pList("NATCODE")
               .value = ""
               .DefaultValue = ""
            End With
            With pList("NATDESC")
               .value = ""
               .DefaultValue = ""
            End With
         Else
'            If mCtrl.TreeViewUpdate Then
'               tv1.Nodes.Add nodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              "SAMP", _
                                              "SubHeader", _
                                              , _
                                              ms_DELETE), _
                             vData(0) & " - " & RS!Sample_Text, _
                             pList("ICON").Icon, _
                             pList("ICON").Icon
'            End If
            With pList("NATCODE")
               .value = vData(0)
               .DefaultValue = vData(0)
            End With
            With pList("NATDESC")
               .value = RS!Sample_Text
               .DefaultValue = RS!Sample_Text
            End With
         End If
         
         RS.Close
         pageId = "Samp"
         tableId = "EDI_Local_Sample_Types"
         itemValue = vData(0)
'         pList.Caption = "Samp"
         
'         EvaluateRecords nodeId, "EDI_Local_Sample_Types", CStr(vData(0))
      
      Case "Anatomical Origin"
         strSQL = "SELECT * " & _
                  "FROM CRIR_Sample_AnatOrigin " & _
                  "WHERE Origin_Code = '" & vData(1) & "'"
         RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
         
         If RS.EOF Then
            With pList("NATCODE")
               .value = ""
               .DefaultValue = ""
            End With
            With pList("NATDESC")
               .value = ""
               .DefaultValue = ""
            End With
         Else
'            If mCtrl.TreeViewUpdate Then
'               tv1.Nodes.Add nodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(1)), _
                                              "ANAT", _
                                              "SubHeader", _
                                              , _
                                              ms_DELETE), _
                             vData(1) & " - " & RS!Origin_Text, _
                             pList("ICON").Icon, _
                             pList("ICON").Icon
'            End If
            With pList("NATCODE")
               .value = vData(1)
               .DefaultValue = vData(1)
            End With
            With pList("NATDESC")
               .value = RS!Origin_Text
               .DefaultValue = RS!Origin_Text
            End With
         End If
         
         RS.Close
'         EvaluateRecords nodeId, "EDI_Local_Sample_AnatOrigin", CStr(vData(1))
'         pList.Caption = "Anat"
         pageId = "Anat"
         tableId = "EDI_Local_Sample_AnatOrigin"
         itemValue = vData(1)
      
      Case "Collection"
         strSQL = "SELECT * " & _
                  "FROM CRIR_Sample_CollectionType " & _
                  "WHERE Collection_Code = '" & vData(4) & "'"
         RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
         
         If RS.EOF Then
            With pList("NATCODE")
               .value = ""
               .DefaultValue = ""
            End With
            With pList("NATDESC")
               .value = ""
               .DefaultValue = ""
            End With
         Else
         
'            tv1.Nodes.Add nodeId, _
                          tvwChild, _
                          mCtrl.NewNodeKey(CStr(vData(4)), _
                                           "COLL", _
                                           "SubHeader", _
                                           , _
                                           ms_DELETE), _
                          vData(4) & " - " & RS!Collection_Text, _
                          pList("ICON").Icon, _
                          pList("ICON").Icon
            With pList("NATCODE")
               .value = vData(4)
               .DefaultValue = vData(4)
            End With
            With pList("NATDESC")
               .value = RS!Collection_Text
               .DefaultValue = RS!Collection_Text
            End With
         End If
         RS.Close
'         EvaluateRecords nodeId, "EDI_Local_Sample_CollectionTypes", CStr(vData(4))
'         pList.Caption = "Coll"
         pageId = "Coll"
         tableId = "EDI_Local_Sample_CollectionTypes"
         itemValue = vData(4)
      
   End Select
   
   pList.Caption = pageId
   EvaluateRecords tNode, tableId, itemValue
   pList.ActivePage = "Desc"
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

Public Sub SubHeader(NodeId As node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim pos As Integer
   Dim i As Integer
   Dim pageId As String
   Dim iconId As Long
   
'  Read owner data and set up control items
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   pList("SAMPDESC").value = objTV.TopLevelNode(NodeId).Text
   
'  Read this node details and set up controls
'   vData = objTV.ReadNodeData(NodeId)
'   pos = InStr(1, vData(1), "-")
'   pList("<Control key>").value = vData(0)
   
'  The sql string to be executed
   Select Case NodeId.Text
      Case "Sample details"
         pageId = "Samp"
         objTV.SQL = "SELECT * " & _
                     "FROM CRIR_Sample_Type " & _
                     "WHERE Sample_Code = '" & vData(0) & "'"
      
      Case "Anatomical Origin Details"
         pageId = "Anat"
         objTV.SQL = "SELECT * " & _
                     "FROM CRIR_Sample_AnatOrigin " & _
                     "WHERE Origin_Code = '" & vData(1) & "'"
      
      Case "Collection details"
         pageId = "Coll"
         objTV.SQL = "SELECT * " & _
                     "FROM CRIR_Sample_CollectionType " & _
                     "WHERE Collection_Code = '" & vData(4) & "'"
      
   End Select
   pList("LOCAL").PageKeys = pageId
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.Child.Index
      Next i
      objTV.UpdateTreeView pList, pageId, mCtrl, NodeId
   Else
      objTV.UpdatePropList pList, pageId, mCtrl, NodeId
   End If
   pList.ActivePage = pageId
'  The SQL condition (if not a stored procedure)
'   pList.Caption = " WHERE  = '" & pList("SAMPDESC").value & "' " '& _
                     "AND = '" & pList("Key2>").value & "'"

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
   Dim tNode As node
   
   mCtrl.TreeViewUpdate = True

'  Specific refresh activities
   
   If objTV.newNode Then
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   
   Set newNode = Nothing
End Function

Public Sub RunWhat(NodeId As node, _
                   Optional Refresh As String = "")
   On Local Error GoTo procEH
   Dim vData As Variant
   Dim lCode As String
   Dim tNode As MSComctlLib.node
   Dim RefreshId As String
   Dim strArray() As String
   
   If TidyUp = False Then
'      tv1.Visible = False
      vData = objTV.ReadNodeData(NodeId)
      pList("LOCAL").PageKeys = "Control"
      mCtrl.TreeViewUpdate = False
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
         
         Case "SubNodes"
            mCtrl.TreeViewUpdate = True
'            objTV.RefreshNode = nodeId.Parent
            SubNodes NodeId
            mCtrl.TreeViewUpdate = False
            
         Case "Item"
            mCtrl.TreeViewUpdate = False
            SubNodes NodeId.Parent
            objTV.RefreshNode = NodeId.Parent
            mCtrl.TreeViewUpdate = True
            
         Case "SubHeader"
            mCtrl.TreeViewUpdate = False
            SubNodes NodeId.Parent
            mCtrl.TreeViewUpdate = True

'            objTV.RefreshNode = nodeId.Parent
'            SubHeader nodeId.Parent
'            mCtrl.TreeViewUpdate = False
            
         Case "SubItem"
            mCtrl.TreeViewUpdate = False
            pList.ActivePage = "Desc"
            SubNodes NodeId.Parent.Parent
         
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
      If objTV.NodeKey(newNode) = "NewNode" Then
         tv1.Nodes.Remove newNode.Index
         tv1.SelectedItem = tv1.Nodes(1)
         TidyUp = True
      End If
      Set newNode = Nothing
   End If
End Function

Public Function Update(pageId As String) As String
   On Error GoTo procEH
   Dim i As Integer
   Dim sCode As String
   Dim aCode As String
   Dim cCode As String
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode)
   sCode = vData(0)
   aCode = vData(1)
   cCode = vData(4)
'  Run a stored procedure or some special SQL
   
   Select Case pList.Caption
      Case "Main"
         If objTV.newNode Then
            strSQL = "INSERT INTO Service_Sample_Codes (" & _
                        "Sample_Description, Active) " & _
                     "VALUES (" & _
                        "'" & pList("SDESC").value & "', " & _
                        Abs(CInt(pList("ACTIVE").value)) & ")"
         Else
            strSQL = "UPDATE Service_Sample_Codes SET " & _
                        "Sample_Description = '" & pList("SDESC").value & "', " & _
                        "Active = " & Abs(CInt(pList("ACTIVE").value)) & _
                     " WHERE Sample_Description = '" & pList("SAMPDESC").value & "'"
         End If
      
      
      Case "Samp"
         strSQL = "UPDATE Service_Sample_Codes SET " & _
                     "Type_Code = '" & pList("NATCODE").value & "' " & _
                  "WHERE Sample_Description = '" & pList("SAMPDESC").value & "'; "
         
         sCode = pList("NATCODE").value
         If (pList("NATCODE").value = pList("SAMPCODE").value) Then
            strSQL = strSQL & _
                     "DELETE FROM EDI_Local_Sample_Types " & _
                     "WHERE National_Code = '" & pList("SAMPCODE").value & "'; "
            
            For i = 1 To pList("LOCAL_COUNT").value
               strSQL = strSQL & _
                        "INSERT INTO EDI_Local_Sample_Types (" & _
                           "National_Code, Local_Text) " & _
                        "VALUES(" & _
                           "'" & pList("NATCODE").value & "', " & _
                           "'" & pList("LKEY_" & i).value & "'); "
            Next i
         End If
         
      Case "Anat"
         strSQL = "UPDATE Service_Sample_Codes SET " & _
                     "AnatomicalOrigin_code = '" & pList("NATCODE").value & "' " & _
                  "WHERE Sample_Description = '" & pList("SAMPDESC").value & "'; "
         
         aCode = pList("NATCODE").value
         If (pList("NATCODE").value = pList("ANATCODE").value) Then
            strSQL = strSQL & _
                     "DELETE FROM EDI_Local_Sample_AnatOrigin " & _
                     "WHERE National_Code = '" & pList("ANATCODE").value & "'; "
            
            For i = 0 To pList("LOCAL_COUNT").value
               strSQL = strSQL & _
                        "INSERT INTO EDI_Local_Sample_AnatOrigin (" & _
                           "National_Code, Local_Text) " & _
                        "VALUES(" & _
                           "'" & pList("ANATCODE").value & "', " & _
                           "'" & pList("LKEY_" & i).value & "'); "
            Next i
         End If
      
      Case "Coll"
         strSQL = "UPDATE Service_Sample_Codes SET " & _
                     "Collection_code = '" & pList("NATCODE").value & "' " & _
                  "WHERE Sample_Description = '" & pList("SAMPDESC").value & "'; "
         
         cCode = pList("NATCODE").value
         If (pList("NATCODE").value = pList("COLLCODE").value) Then
            strSQL = strSQL & _
                     "DELETE FROM EDI_Local_Sample_CollectionTypes " & _
                     "WHERE National_Code = '" & pList("COLLCODE").value & "'; "
            
            For i = 0 To pList("LOCAL_COUNT").value - 1
               strSQL = strSQL & _
                        "INSERT INTO EDI_Local_Sample_CollectionTypes (" & _
                           "National_Code, Local_Text) " & _
                        "VALUES(" & _
                           "'" & pList("COLLCODE").value & "', " & _
                           "'" & pList("LKEY_" & i).value & "'); "
            Next i
         End If
      
   End Select
   
   eClass.FurtherInfo = strSQL
   ICECon.Execute strSQL
   
   objTV.TopLevelNode(objTV.ActiveNode).Text = pList("SDESC").value
   
   If objTV.NodeKey(objTV.ActiveNode) = "TYPE" Then
      objTV.ActiveNode.Text = pList("NATCODE").value & " - " & pList("NATDESC").value
      objTV.TopLevelNode(objTV.ActiveNode).Key = mCtrl.NewNodeKey(sCode, _
                                                                  aCode, _
                                                                  "Detail", _
                                                                  cCode, _
                                                                  ms_DELETE)
   End If
'   Details objTV.TopLevelNode(objTV.ActiveNode)
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "Load???.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

