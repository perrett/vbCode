VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadSpecimenCodes"
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
Private blnReadOnly As Boolean
Private OLESourceNode As Node
Private OLERelative As Long
Private OLEMoveKey As String
Private OLEMoveText As String
'Private plSysBut As New PropertiesListCtl.SysButtonTemplate
'Private myBut As PropertiesListCtl.Buttons

Private Sub Class_Initialize()
   On Error GoTo procEH
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   mCtrl.TreeViewUpdate = True
   fView.FrameToShow = Fra_EDI
   
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"

'     Add control items as required
      .PropertyItems.Add "SDESC", "Sample Description", plpsString, , "Hidden"
'      .PropertyItems.Add "LOCAL_COUNT", "", plpsNumber, -1, "Hidden"
      .PropertyItems.Add "SAMPCODE", "Sample Code", plpsString, , "National Sample Code"
      .PropertyItems.Add "ANATCODE", "Anatomical Origin Code", plpsCustom, , "National anatomical origin code"
      .PropertyItems.Add "COLLCODE", "Collection Code", plpsCustom, , "National collection code"

'     Add pages as required
      .Pages.Add "Main", "Service_Sample_Codes"
      .Pages.Add "Desc", "None"
      
'     Set up the required property items
'      .PropertyItems.Add "SDESC", "Sample Description", plpsString, , "Mandatory - The National sample description"
'      .PropertyItems.Add "SCODE", "Sample Details", plpsString, "<Expand to view details>", "The National Type code"
'      .PropertyItems.Add "SANAT", "Anatomical Origin", plpsString, "<Expand to view details>", "National Anatomical Origin Code"
'      .PropertyItems.Add "SCOLL", "Collection", plpsString, "<Expand to view details>", "The national collection code"
'      .PropertyItems.Add "LOCAL", "Local description(s)", plpsString, , "Local sample type"
'      .PropertyItems.Add "NATCODE", "National Code", plpsCustom, , "The relevant national code"
'      .PropertyItems.Add "NATDESC", "National Description", plpsString, , "The national description"
      .PropertyItems.Add "ACTIVE", "Active", plpsBoolean, , "Is this sample type active?"
   End With
   
   With plist
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 9

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
      
'      With .PropertyItems("SAMPDESC")
'         .PageKeys = "Control"
'         .Visible = False
'      End With
'      With .PropertyItems("LOCAL_COUNT")
'         .PageKeys = "Control"
'         .Visible = False
'         .value = -1
'         .defaultValue = -1
'      End With
      With .PropertyItems("SAMPCODE")
         .PageKeys = "Control"
'         .Tag = "Type_Code"
      End With
      With .PropertyItems("ANATCODE")
         .PageKeys = "Main"
         .Tag = "AnatomicalOrigin_Code"
      End With
      With .PropertyItems("COLLCODE")
         .PageKeys = "Main"
         .Tag = "Collection_Code"
      End With
      
      With .PropertyItems("SDESC")
         .ReadOnly = True
         .Tag = "Sample_Description"
         .DefaultExtension = "Desc"
      End With
'      With .PropertyItems("LOCAL")
'         .Tag = "MENU"
'         .PageKeys = "Control"
'         .DefaultExtension = "SubItem"
'         .Flags = 1
'      End With
'      With .PropertyItems("SCODE")
'         .Tag = "MENU"
'         .PageKeys = "Main"
'         .DefaultExtension = "SubNodes"
'         .defaultValue = "<Expand to view details>"
'         .ReadOnly = True
'      End With
'      With .PropertyItems("SANAT")
'         .Tag = "MENU"
'         .PageKeys = "Main"
'         .DefaultExtension = "SubNodes"
'         .defaultValue = "<Expand to view details>"
'         .ReadOnly = True
'      End With
'      With .PropertyItems("SCOLL")
'         .Tag = "MENU"
'         .PageKeys = "Main"
'         .DefaultExtension = "SubNodes"
'         .defaultValue = "<Expand to view details>"
'         .ReadOnly = True
'      End With
'      With .PropertyItems("NATCODE")
'         .Tag = "None"
'         .PageKeys = "Desc"
'      End With
'      With .PropertyItems("NATDESC")
'         .Tag = "None"
'         .PageKeys = "Desc"
'         .ReadOnly = True
'      End With
      With .PropertyItems("ACTIVE")
         .Tag = "Active"
         .DefaultExtension = "Desc"
      End With
      
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   plist.Pages("Main").Caption = "Specific"
   plist.Caption = "Main"
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
   eClass.CurrentProcedure = "LoadSpecimenCodes.Class.Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let EvNoEdit(blnNewValue As Boolean)
   blnReadOnly = blnNewValue
End Property

Public Sub EvaluateRecords(NodeId As Node, _
                           tableId As String, _
                           itemValue As String)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim rc As Integer
   
   With plist
      .Redraw = False
      If plist("LOCAL_COUNT").value > -1 Then
         For i = 0 To plist("LOCAL_COUNT").value
            .PropertyItems.Remove "LKEY_" & i
         Next i
      End If
      
      If Not (NodeId Is Nothing) Then
         For i = 1 To NodeId.Children
            tv1.Nodes.Remove NodeId.child.Index
         Next i

         tv1.Nodes.Add NodeId, _
                       tvwChild, _
                       mCtrl.NewNodeKey(itemValue, _
                                        "TYPE", _
                                        "Item"), _
                       itemValue & " - " & plist("NATDESC").value, _
                       plist("ICON").Icon, _
                       plist("ICON").Icon
      
         Set nd(0) = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey("LOCAL", _
                                                    "LOCAL", _
                                                    "SubHeader", _
                                                    , _
                                                    ms_ADD), _
                                   "Local Description(s)", _
                                   plist("ICON").Icon, _
                                   plist("ICON").Icon)
      End If
      
      strSQL = "SELECT * " & _
               "FROM " & tableId & _
               " WHERE National_Code = '" & itemValue & "'"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
'      pList("LOCAL_COUNT").value = RS(0)
'      RS.Close
      i = 0
      Do Until RS.EOF
         .PropertyItems.Add "LKEY_" & i, "Local Description " & i, plpsString, RS!Local_Text, "Local Description"
         .PropertyItems("LKEY_" & i).Tag = "Local_Text"
         .PropertyItems("LKEY_" & i).PageKeys = "Desc"
         .PropertyItems("LKEY_" & i).defaultValue = RS!Local_Text
         
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
                          plist("ICON").Icon, _
                          plist("ICON").Icon
         End If
         
         i = i + 1
         RS.MoveNext
      Loop
      RS.Close
      plist("LOCAL_COUNT").value = i - 1
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
                             "Add new Sample Type...", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True
'  SQL string to extract top level branches for node
   strSQL = "SELECT DISTINCT Sample_Text, Sample_Code " & _
            "FROM Service_Sample_Codes " & _
               "INNER JOIN CRIR_Sample_Type " & _
               "ON Type_Code = Sample_Code " & _
            "ORDER BY Sample_Text"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(Trim(RS!Sample_Code & ""), _
                                                 "DROP", _
                                                 "SubNodes", _
                                                 Trim(RS!Sample_Text), _
                                                 ms_DISABLED), _
                                RS!Sample_Text & " (" & RS!Sample_Code & ")", _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
'     Colour code the node?
'      If RS!Active Then
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
                    mCtrl.NewNodeKey("Tmp", _
                                     "Tmp"), _
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
   eClass.CurrentProcedure = "LoadSpecimenCodes.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
      
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
   
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   objTV.ActiveNode = tNode
   objTV.RefreshNode = tNode
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadSpecimenCodes.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
   Dim tNode As Node
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
   vData = objTV.ReadNodeData(NodeId)
   
   If vData(4) <> "DRAG" Then
      Set NodeId = NodeId.Parent
   End If
   
'  Set Control values
'   plist("SDESC").value = NodeId.Text
'   plist("SAMPCODE").value = vData(0)
'   plist("ANATCODE").value = vData(1)
'   plist("COLLCODE").value = vData(4)
'   pList("<Control key>").value = vData(1)
   
'  SQL string to be executed
   objTV.SQL = "SELECT * " & _
               "FROM Service_Sample_Codes " & _
               "WHERE Sample_Description = '" & NodeId.Text & "'"
   
   
   If mCtrl.TreeViewUpdate Then
      If NodeId.Children > 0 Then
         For i = 1 To NodeId.Children
            tv1.Nodes.Remove NodeId.child.Index
         Next i
         objTV.UpdateTreeView plist, "Main", mCtrl, NodeId
      Else
         objTV.UpdatePropList plist, "Main", mCtrl, NodeId
      End If
   Else
      objTV.UpdatePropList plist, "Main", mCtrl, NodeId
   End If
   
'  Specific node actions eg set active colour
      
'   Set nd(0) = objTV.NodeFromKey("<KeyId>", NodeId)
'   nd(0).EnsureVisible
   If plist("ACTIVE").value = True Then
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
   plist.Caption = "Main"
   fView.Show Fra_EDI
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSpecimenCodes.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get KeyControl() As ManageControls
   Set KeyControl = mCtrl
End Property

Public Sub LocalDesc(NodeId As Node)
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   vData = objTV.ReadNodeData(NodeId)
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   strSQL = "SELECT * " & _
            "FROM EDI_Local_Sample_Types " & _
            "WHERE National_Code = '" & vData(0) & "'"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   Do Until RS.EOF
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(CStr(vData(0)), _
                                     "Samp", _
                                     "ShowForm"), _
                    RS!Local_Text, _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
End Sub


Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim i As Integer
   
   With plist
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
   plist("LOCAL_COUNT").value = plist("LOCAL_COUNT").value + 1
   plist.PropertyItems.Add "LKEY_" & plist("LOCAL_COUNT").value, _
                           "Local Description " & plist("LOCAL_COUNT").value, _
                           plpsString, _
                           , _
                           "A local description"
   plist("LKEY_" & plist("LOCAL_COUNT")).Selected = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadSpecimenCodes.MenuAddEntry"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewMapping()
   On Error GoTo procEH
   Dim i As Integer

   If plist("LOCAL_COUNT").value > -1 Then
      For i = 0 To plist("LOCAL_COUNT").value
         plist.PropertyItems.Remove "LKEY_" & i
      Next i
   End If
   
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
                                                   "Detail", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
   End If
   
   objTV.ActiveNode = newNode
   objTV.RefreshNode = newNode
   plist.ActivePage = "Main"
'   objTView.ActiveNode = newNode
   fView.Show Fra_EDI
   
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSpecimenCodes.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubNodes(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim tableId As String
   Dim itemValue As String
   Dim PageId As String
   Dim tNode As Node
   
   If mCtrl.TreeViewUpdate Then
      Set tNode = NodeId
   Else
      Set tNode = Nothing
   End If
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   objTV.RefreshNode = NodeId
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i

'   With plist("NATCODE")
'      .value = vData(0)
'      .defaultValue = vData(0)
'   End With
'   With plist("NATDESC")
'      .value = NodeId.Text
'      .defaultValue = NodeId.Text
'   End With
   
   Set nd(0) = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(CStr(vData(0)), _
                                              "DRAG", _
                                              "LocalDesc"), _
                             "Local Sample Descriptions", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   
   nd(0).Bold = True
   
   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("tmp", _
                                  "tnp", _
                                  "None"), _
                 "Please wait...", _
                 plist("ICON").Icon, _
                 plist("ICON").Icon
                 
   strSQL = "SELECT Sample_Description, " & _
               "AnatomicalOrigin_Code, " & _
               "Origin_Text, " & _
               "ssc.Collection_Code, " & _
               "Collection_Text, " & _
               "Active " & _
            "FROM Service_Sample_Codes ssc " & _
               "LEFT JOIN CRIR_Sample_AnatOrigin " & _
               "ON AnatomicalOrigin_Code = Origin_Code " & _
               "LEFT JOIN CRIR_Sample_CollectionType cc " & _
               "ON ssc.Collection_Code = cc.Collection_Code " & _
            "WHERE Type_Code = '" & vData(0) & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(CStr(vData(0)), _
                                                 "Spec", _
                                                 "Detail", _
                                                 "DRAG", _
                                                 ms_DISABLED), _
                                RS!Sample_Description, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
      
      If RS!Active Then
         nd(0).ForeColor = BPBLUE
      Else
         nd(0).ForeColor = BPRED
      End If
      
      If (RS!AnatomicalOrigin_Code & "") <> "" Then
      
         tv1.Nodes.Add nd(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey(CStr(vData(0)), _
                                        RS!AnatomicalOrigin_Code, _
                                        "Detail"), _
                       RS!AnatomicalOrigin_Code & "( " & RS!Origin_Text & ")", _
                       plist("ICON").Icon, _
                       plist("ICON").Icon
      End If
      
      If (RS!Collection_Code & "") <> "" Then
      
         tv1.Nodes.Add nd(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey(CStr(vData(0)), _
                                        RS!Collection_Code, _
                                        "Detail"), _
                       RS!Collection_Code & " (" & RS!Collection_Text & ")", _
                       plist("ICON").Icon, _
                       plist("ICON").Icon
      End If
      
      RS.MoveNext
   
   Loop
   RS.Close
'         PageId = "Samp"
'         tableId = "EDI_Local_Sample_Types"
'         itemValue = vData(0)
'         pList.Caption = "Samp"
         
'         EvaluateRecords nodeId, "EDI_Local_Sample_Types", CStr(vData(0))
      
   plist.Caption = PageId
   
'   EvaluateRecords tNode, tableId, itemValue
   plist.ActivePage = "Desc"
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSpecimenCodes.SubNodes"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub xSubNodes(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim tableId As String
   Dim itemValue As String
   Dim PageId As String
   Dim tNode As Node
   
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
                  "FROM Service_Sample_Codes " & _
                  "WHERE Sample_Code = '" & vData(0) & "'"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         If RS.EOF Then
            With plist("NATCODE")
               .value = "<None Set>"
               .defaultValue = ""
            End With
            With plist("NATDESC")
               .value = "<None Set>"
               .defaultValue = ""
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
            With plist("NATCODE")
               .value = vData(0)
               .defaultValue = vData(0)
            End With
            With plist("NATDESC")
               .value = RS!Sample_Text
               .defaultValue = RS!Sample_Text
            End With
         End If
         
         RS.Close
         PageId = "Samp"
         tableId = "EDI_Local_Sample_Types"
         itemValue = vData(0)
'         pList.Caption = "Samp"
         
'         EvaluateRecords nodeId, "EDI_Local_Sample_Types", CStr(vData(0))
      
      Case "Anatomical Origin"
         strSQL = "SELECT * " & _
                  "FROM CRIR_Sample_AnatOrigin " & _
                  "WHERE Origin_Code = '" & vData(1) & "'"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         If RS.EOF Then
            With plist("NATCODE")
               .value = "<None Set>"
               .defaultValue = ""
            End With
            With plist("NATDESC")
               .value = "<None Set>"
               .defaultValue = ""
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
            With plist("NATCODE")
               .value = vData(1)
               .defaultValue = vData(1)
            End With
            With plist("NATDESC")
               .value = RS!Origin_Text
               .defaultValue = RS!Origin_Text
            End With
         End If
         
         RS.Close
'         EvaluateRecords nodeId, "EDI_Local_Sample_AnatOrigin", CStr(vData(1))
'         pList.Caption = "Anat"
         PageId = "Anat"
         tableId = "EDI_Local_Sample_AnatOrigin"
         itemValue = vData(1)
      
      Case "Collection"
         strSQL = "SELECT * " & _
                  "FROM CRIR_Sample_CollectionType " & _
                  "WHERE Collection_Code = '" & vData(4) & "'"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         If RS.EOF Then
            With plist("NATCODE")
               .value = "<None Set>"
               .defaultValue = ""
            End With
            With plist("NATDESC")
               .value = "<None Set>"
               .defaultValue = ""
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
            With plist("NATCODE")
               .value = vData(4)
               .defaultValue = vData(4)
            End With
            With plist("NATDESC")
               .value = RS!Collection_Text
               .defaultValue = RS!Collection_Text
            End With
         End If
         RS.Close
'         EvaluateRecords nodeId, "EDI_Local_Sample_CollectionTypes", CStr(vData(4))
'         pList.Caption = "Coll"
         PageId = "Coll"
         tableId = "EDI_Local_Sample_CollectionTypes"
         itemValue = vData(4)
      
   End Select
   
   plist.Caption = PageId
   EvaluateRecords tNode, tableId, itemValue
   plist.ActivePage = "Desc"
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSpecimenCodes.SubNodes"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub SubHeader(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim pos As Integer
   Dim i As Integer
   Dim PageId As String
   Dim iconId As Long
   
'  Read owner data and set up control items
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   plist("SAMPDESC").value = objTV.TopLevelNode(NodeId).Text
   
'  Read this node details and set up controls
'   vData = objTV.ReadNodeData(NodeId)
'   pos = InStr(1, vData(1), "-")
'   pList("<Control key>").value = vData(0)
   
'  The sql string to be executed
   Select Case NodeId.Text
      Case "Sample details"
         PageId = "Samp"
         objTV.SQL = "SELECT * " & _
                     "FROM CRIR_Sample_Type " & _
                     "WHERE Sample_Code = '" & vData(0) & "'"
      
      Case "Anatomical Origin Details"
         PageId = "Anat"
         objTV.SQL = "SELECT * " & _
                     "FROM CRIR_Sample_AnatOrigin " & _
                     "WHERE Origin_Code = '" & vData(1) & "'"
      
      Case "Collection details"
         PageId = "Coll"
         objTV.SQL = "SELECT * " & _
                     "FROM CRIR_Sample_CollectionType " & _
                     "WHERE Collection_Code = '" & vData(4) & "'"
      
   End Select
   plist("LOCAL").PageKeys = PageId
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, PageId, mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, PageId, mCtrl, NodeId
   End If
   plist.ActivePage = PageId
'  The SQL condition (if not a stored procedure)
'   pList.Caption = " WHERE  = '" & pList("SAMPDESC").value & "' " '& _
                     "AND = '" & pList("Key2>").value & "'"

   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSpecimenCodes.SubHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Public Function Refresh() As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = True

'  Specific refresh activities
   
   If objTV.newNode Then
      objTV.ActiveNode.Text = plist("SDESC").value
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("", "", "Detail", , ms_DELETE, False)
   End If
   objTV.RefreshNode = objTV.ActiveNode
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
'      plist("LOCAL").PageKeys = "Control"
      mCtrl.TreeViewUpdate = False
      Select Case vData(2)
         Case "FirstView"
            FirstView NodeId
         
         Case "Detail"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId
            Details NodeId
            mCtrl.TreeViewUpdate = False
         
         Case "Desc"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId
            objTV.ActiveNode = NodeId.Parent
            Details NodeId.Parent
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
            
         Case "LocalDesc"
            LocalDesc NodeId
            
         Case "SubHeader"
            mCtrl.TreeViewUpdate = False
            SubNodes NodeId.Parent
            mCtrl.TreeViewUpdate = True

'            objTV.RefreshNode = nodeId.Parent
'            SubHeader nodeId.Parent
'            mCtrl.TreeViewUpdate = False
            
         Case "SubItem"
            mCtrl.TreeViewUpdate = False
            plist.ActivePage = "Desc"
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
   eClass.CurrentProcedure = "LoadSpecimenCodes.RunWhat"
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
   Dim TVEffect As Long
   Dim srcDets As String
   Dim blnFirst As Boolean
   Dim msgStr As String
   Dim sCode As String
   
   Set OLESourceNode = objTV.ActiveNode
   srcDets = OLESourceNode.Parent.Text
   
   vData = objTV.ReadNodeData(OLESourceNode)
'   srcDets = data.GetData(vbCFText)
   
   OLEMoveKey = objTV.ActiveNode.Key
   OLEMoveText = objTV.ActiveNode.Text
   
   TV_DragDrop = vbDropEffectNone

   blnFirst = False
   msgStr = "Are you sure you wish to move " & objTV.ActiveNode.Text & " from " & srcDets & " to " & _
             NodeId.Text & "?"
   
   OLERelative = tvwChild
   
   If MsgBox(msgStr, vbYesNo + vbQuestion, "Sample Category amendment") = vbYes Then
      vData = objTV.ReadNodeData(NodeId)
      sCode = vData(0)
      tv1.Nodes.Remove objTV.ActiveNode.Index
      objTV.ActiveNode = Nothing
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                OLEMoveKey, _
                                OLEMoveText, _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      nd(0).ForeColor = BPBLUE
      NodeId.Expanded = True
      
      strSQL = "UPDATE Service_Sample_Codes SET " & _
                  "Type_Code = '" & sCode & "' " & _
               "WHERE Sample_Description = '" & OLEMoveText & "' " & _
                  "AND Type_Code = '" & vData(0) & "'"
      
      iceCon.Execute strSQL
      Debug.Print strSQL
   End If
   
   TV_DragDrop = vbDropEffectNone
End Function

Public Function TV_DragOver(NodeId As Node) As Long
   Dim dhData As Variant
   Dim lngEffect As Long
   
   vData = objTV.ReadNodeData(NodeId)
'   vData = objTV.ReadNodeData(objTV.ActiveNode)
   lngEffect = vbDropEffectNone
   
   If vData(1) = "DROP" Then
      lngEffect = vbDropEffectMove
   End If
   
   TV_DragOver = lngEffect
End Function

Public Function Update(PageId As String) As String
   On Error GoTo procEH
   Dim i As Integer
   Dim sCode As String
   Dim aCode As String
   Dim cCode As String
   Dim tNode As Node
   
   Set tNode = objTV.TopLevelNode(, True)
   tNode.Expanded = False
   vData = objTV.ReadNodeData(tNode)
   sCode = vData(0)
   aCode = vData(1)
   cCode = vData(4)
'  Run a stored procedure or some special SQL
   
   strSQL = "UPDATE Service_Sample_Codes SET " & _
               "Sample_Description = '" & plist("SDESC").value & "', " & _
               "AnatomicalOrigin_code = '" & plist("ANATCODE").value & "', " & _
               "Collection_code = '" & plist("COLLCODE").value & "', " & _
               "Active = " & CInt(plist("ACTIVE").value) & _
            " WHERE Sample_Description = '" & tNode.Text & "' " & _
               "AND Type_Code = '" & vData(0) & "'"

   iceCon.Execute strSQL
   
   If plist("ANATCODE").value = "" And plist("COLLCODE").value = "" Then
      For i = 1 To tNode.Children
         tv1.Nodes.Remove tNode.child.Index
      Next i
   Else
      tv1.Nodes.Add tNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey("tmp", _
                                     "tmp", _
                                     "None"), _
                    "Please wait...", _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
   End If
'   Select Case plist.Caption
'      Case "Main"
'         If objTV.newNode Then
'            strSQL = "INSERT INTO Service_Sample_Codes (" & _
'                        "Sample_Description, Active) " & _
'                     "VALUES (" & _
'                        "'" & plist("SDESC").value & "', " & _
'                        Abs(CInt(plist("ACTIVE").value)) & ")"
'         Else
'            strSQL = "UPDATE Service_Sample_Codes SET " & _
'                        "Sample_Description = '" & plist("SDESC").value & "', " & _
'                        "Active = " & Abs(CInt(plist("ACTIVE").value)) & _
'                     " WHERE Sample_Description = '" & plist("SAMPDESC").value & "'"
'         End If
'
'
'      Case "Samp"
'         strSQL = "UPDATE Service_Sample_Codes SET " & _
'                     "Type_Code = '" & Left(plist("NATCODE").value, 5) & "' " & _
'                  "WHERE Sample_Description = '" & plist("SAMPDESC").value & "'; "
'
'         sCode = plist("NATCODE").value
'         If (plist("NATCODE").value = plist("SAMPCODE").value) Then
'            strSQL = strSQL & _
'                     "DELETE FROM EDI_Local_Sample_Types " & _
'                     "WHERE National_Code = '" & plist("SAMPCODE").value & "'; "
'
'            For i = 0 To plist("LOCAL_COUNT").value
'               strSQL = strSQL & _
'                        "INSERT INTO EDI_Local_Sample_Types (" & _
'                           "National_Code, Local_Text) " & _
'                        "VALUES(" & _
'                           "'" & plist("NATCODE").value & "', " & _
'                           "'" & plist("LKEY_" & i).value & "'); "
'            Next i
'         End If
'         sCode = plist("NATCODE").value
'
'      Case "Anat"
'         strSQL = "UPDATE Service_Sample_Codes SET " & _
'                     "AnatomicalOrigin_code = '" & Left(plist("NATCODE").value, 5) & "' " & _
'                  "WHERE Sample_Description = '" & plist("SAMPDESC").value & "'; "
'
'         aCode = plist("NATCODE").value
'         If (plist("NATCODE").value = plist("ANATCODE").value) Then
'            strSQL = strSQL & _
'                     "DELETE FROM EDI_Local_Sample_AnatOrigin " & _
'                     "WHERE National_Code = '" & plist("ANATCODE").value & "'; "
'
'            For i = 0 To plist("LOCAL_COUNT").value
'               strSQL = strSQL & _
'                        "INSERT INTO EDI_Local_Sample_AnatOrigin (" & _
'                           "National_Code, Local_Text) " & _
'                        "VALUES(" & _
'                           "'" & plist("ANATCODE").value & "', " & _
'                           "'" & plist("LKEY_" & i).value & "'); "
'            Next i
'         End If
'         aCode = plist("NATCODE").value
'
'      Case "Coll"
'         strSQL = "UPDATE Service_Sample_Codes SET " & _
'                     "Collection_code = '" & Left(plist("NATCODE").value, 5) & "' " & _
'                  "WHERE Sample_Description = '" & plist("SAMPDESC").value & "'; "
'
'         cCode = plist("NATCODE").value
'         If (plist("NATCODE").value = plist("COLLCODE").value) Then
'            strSQL = strSQL & _
'                     "DELETE FROM EDI_Local_Sample_CollectionTypes " & _
'                     "WHERE National_Code = '" & plist("COLLCODE").value & "'; "
'
'            For i = 0 To plist("LOCAL_COUNT").value
'               strSQL = strSQL & _
'                        "INSERT INTO EDI_Local_Sample_CollectionTypes (" & _
'                           "National_Code, Local_Text) " & _
'                        "VALUES(" & _
'                           "'" & plist("COLLCODE").value & "', " & _
'                           "'" & plist("LKEY_" & i).value & "'); "
'            Next i
'         End If
'         cCode = plist("NATCODE").value
'
'   End Select
'
'   eClass.FurtherInfo = strSQL
'   iceCon.Execute strSQL
'
'   objTV.TopLevelNode(objTV.ActiveNode).Text = plist("SDESC").value
'
'   If plist.Caption <> "Main" Then
'      Set tNode = objTV.ActiveNode
'      vData = objTV.ReadNodeData(tNode)
'      Do Until vData(2) = "SubNodes"
'         Set tNode = tNode.Parent
'         vData = objTV.ReadNodeData(tNode)
'      Loop
'      objTV.ActiveNode = tNode
'      objTV.TopLevelNode.Key = mCtrl.NewNodeKey(sCode, _
'                                   aCode, _
'                                   "Details", _
'                                   cCode)
'   End If
'
'   If objTV.nodeKey(objTV.ActiveNode) = "TYPE" Then
'      objTV.ActiveNode.Text = plist("NATCODE").value & " - " & plist("NATDESC").value
'      objTV.TopLevelNode(objTV.ActiveNode).Key = mCtrl.NewNodeKey(sCode, _
'                                                                  aCode, _
'                                                                  "Detail", _
'                                                                  cCode, _
'                                                                  ms_DELETE)
'   End If
'   Details objTV.TopLevelNode(objTV.ActiveNode)
   Details tNode
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadSpecimenCodes.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

