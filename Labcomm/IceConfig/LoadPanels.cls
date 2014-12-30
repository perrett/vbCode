VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private plist As PropertiesList
Private nd(2) As Node
Private mCtrl As New ManageControls
Private strSQL As String
Private strArray() As String
Private vData As Variant
Private newNode As Node
Private OLESourceNode As Node
Private OLERelative As Long
Private OLEMoveKey As String
Private OLEMoveText As String

Private Sub Class_Initialize()
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   objTV.Origin = "P_DRAGDROP"
'   objctrl.PracticeId = "System"
'   objctrl.SetDbStructures "Request_Panels"
'   objctrl.SetDbStructures "Request_Panels_Pages"
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Panels", "Request_Panels"
      .Pages.Add "Pages", "Request_Panels_Pages"
      .PropertyItems.Add "PANELNAME", "Panel Name", plpsString, , "Mandatory - The name of this panel"
      .PropertyItems.Add "PANELTYPE", "Panel Type", plpsList, , "Read Only - The type of panel"
      .PropertyItems.Add "PAGENAME", "Page Name", plpsString, , "The title of the page"
      .PropertyItems.Add "PAGESEQ", "Panel Sequence", plpsNumber, , "Read Only - The position this page will appear on the panel"
      .PropertyItems.Add "PANELID", "This Is Hidden", plpsString, , "Hidden from User"
      
      With .PropertyItems("PANELNAME")
         .Tag = "PanelName"
         .defaultValue = "New Panel"
         .PageKeys = "Panels"
      End With
      With .PropertyItems("PANELTYPE")
         .Tag = "PanelType"
         .ListItems.Add "Type", 1
         .defaultValue = 1
         .ReadOnly = True
         .PageKeys = "Panels"
      End With
      With .PropertyItems("PAGENAME")
         .Tag = "PageName"
         .defaultValue = "New Page"
         .PageKeys = "Pages"
      End With
      With .PropertyItems("PAGESEQ")
         .Tag = "Sequence"
         .PageKeys = "Pages"
         .ReadOnly = True
      End With
      With .PropertyItems("PANELID")
         .Tag = "PanelID"
         .PageKeys = "Pages"
         .Visible = False
         .ReadOnly = True
      End With
   End With
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
End Sub

Public Sub Delete(NodeId As Node)
   Dim RS As New ADODB.Recordset
   Dim strPanels As String
   Dim noItems As Integer
   
   vData = objTV.ReadNodeData(NodeId)
   strSQL = "SELECT Test_Index, Screen_Caption " & _
            "FROM Request_Tests " & _
            "WHERE Screen_Panel = " & vData(0)
   If Not (NodeId.Parent Is Nothing) Then
      strSQL = strSQL & " AND Screen_Panel_Page = '" & NodeId.Text & "'"
   End If
   
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.RecordCount > 0 Then
      noItems = 0
      strPanels = NodeId.Text & " has the following tests assigned:" & vbCrLf & vbCrLf
      Do Until RS.EOF Or noItems > 10
         strPanels = strPanels & RS!Screen_Caption & vbCrLf
         noItems = noItems + 1
         RS.MoveNext
      Loop
      If noItems <> RS.RecordCount Then
         strPanels = strPanels & "... more " & vbCrLf
      End If
      
      strPanels = strPanels & vbCrLf & "Remove these tests before attempting to delete the panel"
      MsgBox strPanels, vbInformation, _
             "Unable to delete Panel/Page - (" & RS.RecordCount & ") tests still assigned"
   Else
      If NodeId.Parent Is Nothing Then
         strSQL = "DELETE FROM Request_Panels_Pages " & _
                  "WHERE PanelID = " & vData(0) & "; " & _
                  "DELETE FROM Request_Panels " & _
                  "WHERE PanelId = " & vData(0)
      Else
         strSQL = strSQL & "DELETE FROM Request_Panels_Pages " & _
                           "WHERE PanelId = " & vData(0) & _
                              " AND PageName = '" & NodeId.Text & "'"
      End If
      iceCon.Execute strSQL
      tv1.Nodes.Remove objTV.ActiveNode.Index
   End If
   RS.Close
   Set RS = Nothing
End Sub

Public Sub MenuAddEntry()
   Dim RS As New ADODB.Recordset
   
   vData = objTV.ReadNodeData(objTV.ActiveNode)
   strSQL = "SELECT COUNT(PanelId) as Tot " & _
            "FROM Request_Panels_Pages " & _
            "WHERE PanelID = " & vData(0)
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS!Tot >= 15 Then
      MsgBox "This panel already contains 15 pages. " & _
             "This is the maximum that can realistically be displayed", vbInformation, _
             "Unable to Insert New Page"
   Else
      RS.Close
      strSQL = "SELECT MAX(Sequence) " & _
               "FROM Request_Panels_Pages " & _
               "WHERE PanelID = " & vData(0)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      With frmMain.edipr
         .PropertyItems("PAGENAME").value = .PropertyItems("PAGENAME").defaultValue
         If IsNull(RS(0)) Then
            .PropertyItems("PAGESEQ").value = 1
         Else
            .PropertyItems("PAGESEQ").value = Val(RS(0) + 1)
         End If
         .PropertyItems("PAGENAME").Enabled = True
         .PropertyItems("PAGESEQ").Enabled = True
         .PropertyItems("PANELNAME").value = objTV.ActiveNode.Text
         .PropertyItems("PANELNAME").Enabled = False
         .PropertyItems("PANELTYPE").value = objTV.nodeKey(objTV.ActiveNode)
         .PropertyItems("PANELTYPE").Enabled = False
         .PropertyItems("PANELID").value = vData(0)
   '      .Caption = "WHERE PanelId = " & PanelId & " AND PageName = '" & objtv.ActiveNode.Text & "'"
         .ActivePage = "Pages"
      End With
      If newNode Is Nothing Then
         Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                     tvwChild, _
                                     mCtrl.NewNodeKey("New", _
                                                      "NewNode", _
                                                      "PanelDetails", _
                                                      , _
                                                      ms_BOTH, _
                                                      True), _
                                     "New", _
                                     2, _
                                     2)
      End If
      objTV.ActiveNode = newNode
   End If
   RS.Close
   Set RS = Nothing
   fView.Show Fra_EDI
End Sub

Public Sub NewPanel()
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   strSQL = "SELECT COUNT(PanelID) as Tot FROM Request_Panels"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS!Tot >= 10 Then
      MsgBox "There are already 10 panels on this system. " & _
             "This is the maximum that can realistically be displayed", vbInformation, _
             "Unable to Insert New Panel"
   Else
      With frmMain.edipr
         .Redraw = False
         For i = 1 To .PropertyItems.Count
            .PropertyItems(i).value = .PropertyItems(i).defaultValue
         Next i
         .PropertyItems("PAGENAME").Enabled = False
         .PropertyItems("PAGENAME").value = ""
         .PropertyItems("PAGESEQ").Enabled = False
         .PropertyItems("PAGESEQ").value = ""
         .PropertyItems("PANELNAME").Enabled = True
         .PropertyItems("PANELTYPE").Enabled = True
         .PropertyItems("PANELSEQ").value = Val(RS!Tot & "") + 1
         .ActivePage = "Panels"
         .Redraw = True
      End With
      If newNode Is Nothing Then
         Set newNode = tv1.Nodes.Add(, _
                                     , _
                                     mCtrl.NewNodeKey("New", _
                                                      "NewNode", _
                                                      "PanelDetails", _
                                                      , _
                                                      ms_BOTH, _
                                                      True), _
                                     "New", _
                                     2, _
                                     2)
      End If
      objTV.ActiveNode = newNode
   End If
   RS.Close
   Set RS = Nothing
   fView.Show Fra_EDI
End Sub

Public Sub PageDetails(PanelId As String)
   Dim RS As New ADODB.Recordset
   
   strSQL = "SELECT * " & _
            "FROM Request_Panels_Pages " & _
            "WHERE PanelID = " & PanelId & _
               " AND PageName = '" & objTV.ActiveNode.Text & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   With frmMain.edipr
      .PropertyItems("PAGENAME").value = RS!PageName
      .PropertyItems("PAGESEQ").value = RS!Sequence
      .PropertyItems("PAGENAME").Enabled = True
      .PropertyItems("PAGESEQ").Enabled = True
      .PropertyItems("PANELNAME").value = objTV.ActiveNode.Parent.Text
      .PropertyItems("PANELNAME").Enabled = False
      .PropertyItems("PANELTYPE").value = objTV.nodeKey(objTV.ActiveNode.Parent)
      .PropertyItems("PANELTYPE").Enabled = False
      .PropertyItems("PANELID").value = PanelId
      .Caption = "WHERE PanelId = " & PanelId & " AND PageName = '" & objTV.ActiveNode.Text & "'"
      .ActivePage = "Pages"
   End With
   RS.Close
   Set RS = Nothing
   Set newNode = Nothing
End Sub

Public Sub Panels()
   Dim RS As New ADODB.Recordset
   
'   panelName = strNewValue
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("New", _
                                              "PANEL", _
                                              "NewPanel", _
                                              "DROP"), _
                             "Add New Panel", _
                             2, _
                             2)
   nd(0).Bold = True
   
   strSQL = "SELECT * " & _
            "FROM Request_Panels " & _
            "ORDER BY Sequence"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!PanelId, _
                                                 "PANEL", _
                                                 "PanelDetails", _
                                                 "DRAG", _
                                                 ms_BOTH), _
                                RS!panelName, _
                                2, _
                                2)
      nd(0).Checked = True
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Temp", "Temp"), _
                    "Please wait...", _
                    2, _
                    2
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
End Sub

Public Sub PanelDetails(NodeId As Node, Optional Refresh As Boolean = False)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   vData = objTV.ReadNodeData(NodeId)
   strSQL = "SELECT * " & _
            "FROM Request_Panels " & _
            "WHERE PanelID = " & vData(0) & _
            "ORDER BY PanelID"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   With frmMain.edipr
      .PropertyItems("PAGENAME").Enabled = False
      .PropertyItems("PAGENAME").value = ""
      .PropertyItems("PAGESEQ").Enabled = False
      .PropertyItems("PAGESEQ").value = ""
      .PropertyItems("PANELNAME").value = RS!panelName
      .PropertyItems("PANELNAME").Enabled = True
      .PropertyItems("PANELTYPE").value = Format(RS!panelType)
      .PropertyItems("PANELTYPE").Enabled = True
      .Caption = "WHERE PanelID = " & vData(0)
      .ActivePage = "Panels"
   End With
   
   NodeId.Text = RS!panelName
   NodeId.Key = mCtrl.NewNodeKey(RS!PanelId, _
                                 "PANEL", _
                                 "PanelDetails", _
                                 "DRAG", _
                                 ms_BOTH)
   RS.Close
   strSQL = "SELECT * " & _
            "FROM Request_Panels_Pages " & _
            "WHERE PanelID = " & vData(0) & _
            " ORDER BY Sequence"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(CStr(vData(0)), _
                                                 "PAGE", _
                                                 "PageDetails", _
                                                 "DRAG", _
                                                 ms_DELETE), _
                                RS!PageName, _
                                2, _
                                2)
      nd(0).Checked = True
      RS.MoveNext
   Loop
   fView.Show Fra_EDI
   RS.Close
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadPanels.PanelDetails"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function Refresh() As String
   Dim RS As New ADODB.Recordset
   
   With frmMain.edipr
      If objTV.newNode Then
         If objTV.ActiveNode.Parent Is Nothing Then
            strSQL = "SELECT MAX(PanelID) " & _
                     "FROM Request_Panels"
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            objTV.ActiveNode.Text = .PropertyItems("PANELNAME").value
            objTV.ActiveNode.Key = mCtrl.NewNodeKey(RS(0), _
                                                    .PropertyItems("PANELTYPE").value, _
                                                    "PanelDetails", _
                                                    , _
                                                    ms_BOTH)
            RS.Close
         Else
            objTV.ActiveNode.Text = .PropertyItems("PAGENAME").value
            objTV.ActiveNode.Key = mCtrl.NewNodeKey(.PropertyItems("PANELID").value, _
                                                    .PropertyItems("PAGESEQ").value, _
                                                    "PageDetails", _
                                                    , _
                                                    ms_DELETE)
         End If
      Else
         If Not (objTV.ActiveNode.Parent Is Nothing) Then
            strSQL = "UPDATE Request_Tests " & _
                     "SET Screen_Panel_Page = '" & .PropertyItems("PAGENAME").value & "' " & _
                     "WHERE Screen_Panel_Page = '" & objTV.ActiveNode.Text & "'"
            iceCon.Execute strSQL
            objTV.ActiveNode.Text = .PropertyItems("PAGENAME").value
            objTV.ActiveNode.Key = mCtrl.NewNodeKey(.PropertyItems("PANELID").value, _
                                                    .PropertyItems("PAGESEQ").value, _
                                                    "PageDetails", _
                                                    , _
                                                    ms_DELETE)
         End If
      End If
   End With
   objTV.RefreshNode = objTV.ActiveNode
End Function

Public Sub RunWhat(NodeId As Node, Optional Refresh As String = "")
   If TidyUp Then
      fView.FrameToShow = Fra_HELP
   Else
      fView.FrameToShow = Fra_EDI
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "PanelDetails"
            If Not newNode Is Nothing Then
               If newNode.Text = "New" Then
                  tv1.Nodes.Remove newNode.Index
               End If
               Set newNode = Nothing
            End If
            PanelDetails NodeId
            
         Case "PageDetails"
            If Not newNode Is Nothing Then
               If newNode.Text = "New" Then
                  tv1.Nodes.Remove newNode.Index
               End If
               Set newNode = Nothing
            End If
            PageDetails CStr(vData(0))
            
         Case "NewPanel"
            NewPanel
            
      End Select
   End If
End Sub

Public Function TV_DragDrop(NodeId As Node, _
                            Optional data As MSComctlLib.DataObject = Nothing) As Long
   Dim TVEffect As Long
   Dim srcDets As String
   Dim blnFirst As Boolean
   Dim msgStr As String
   
   vData = objTV.ReadNodeData(NodeId)
   srcDets = data.GetData(vbCFText)
   
   OLEMoveKey = objTV.ActiveNode.Key
   OLEMoveText = objTV.ActiveNode.Text
   
   TV_DragDrop = vbDropEffectNone
   
   blnFirst = False
   msgStr = "Are you sure you wish to move '"
   
   Set OLESourceNode = NodeId
   If srcDets = "PANEL" Then
      If NodeId.Root = NodeId Then
         msgStr = msgStr & objTV.ActiveNode.Text & "' above '" & tv1.Nodes(1).Next.Text & "'?"
         OLERelative = tvwNext
      Else
         msgStr = msgStr & objTV.ActiveNode.Text & "' below '" & NodeId.Text & "'?"
         OLERelative = tvwNext
      End If
   
   Else
      If vData(1) = "PANEL" Then
         msgStr = msgStr & objTV.ActiveNode.Text & "' above '" & NodeId.child.FirstSibling.Text & "'?"
         Set OLESourceNode = NodeId.child.FirstSibling
         OLERelative = tvwFirst
      Else
         msgStr = msgStr & objTV.ActiveNode.Text & "' below '" & NodeId.Text & "'?"
         OLERelative = tvwNext
      End If
   End If
   
'   If blnFirst Then
'      If MsgBox("Are you sure you wish to move '" & objTV.ActiveNode.Text & _
'                "' above '" & nodeId.Child.FirstSibling.Text & "'?", _
'                vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
'         Set OLESourceNode = nodeId.Child
'         OLERelative = tvwFirst
'         TVEffect = vbDropEffectMove
'      End If
'   Else
'      If MsgBox("Are you sure you wish to move '" & objTV.ActiveNode.Text & _
'                "' below '" & nodeId.Text & "'?", _
'                vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
'         Set OLESourceNode = nodeId
'         OLERelative = tvwNext
'         TVEffect = vbDropEffectMove
'      End If
'   End If
   
'   If TVEffect = vbDropEffectMove Then
      If MsgBox(msgStr, vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
         tv1.Nodes.Remove objTV.ActiveNode.Index
         objTV.ActiveNode = Nothing
         tv1.Nodes.Add OLESourceNode, _
                       OLERelative, _
                       OLEMoveKey, _
                       OLEMoveText, _
                       1
         
         If srcDets = "PANEL" Then
            WritePanelSequence
         Else
'            If OLERelative = tvwFirst Then
'               WritePageSequence OLESourceNode
'            Else
               WritePageSequence OLESourceNode.Parent
'            End If
         End If
      End If
'   End If
   TV_DragDrop = vbDropEffectNone
End Function

Public Function TV_DragOver(NodeId As Node) As Long
   Dim dhData As Variant
   Dim lngEffect As Long
   
   dhData = objTV.ReadNodeData(NodeId)
   vData = objTV.ReadNodeData(objTV.ActiveNode)
   lngEffect = vbDropEffectNone
   
   If vData(1) = "PANEL" Then
      If dhData(1) = "PANEL" Then
         lngEffect = vbDropEffectMove
      End If
   ElseIf vData(1) = "PAGE" Then
      If dhData(0) = vData(0) Then
         lngEffect = vbDropEffectMove
      End If
   End If
   TV_DragOver = lngEffect
End Function

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

Public Sub WritePageSequence(NodeId As Node)
   Dim intSeq As Integer
   
   Set nd(0) = NodeId.child
   strSQL = "DELETE FROM Request_Panels_Pages " & _
            "WHERE PanelId = " & objTV.NodeLevel(NodeId)
   iceCon.Execute strSQL
   
   intSeq = 1
   Do Until nd(0) Is Nothing
      vData = objTV.ReadNodeData(nd(0))
      strSQL = "INSERT INTO Request_Panels_Pages " & _
               "(PanelID, PageName, Sequence) " & _
               "VALUES ( " & _
                  vData(0) & ", '" & _
                  nd(0).Text & "', " & _
                  intSeq & ")"
      iceCon.Execute strSQL
      intSeq = intSeq + 1
      Set nd(0) = nd(0).Next
   Loop
   
End Sub

Public Sub WritePanelSequence()
   On Error GoTo procEH
   Dim intSeq As Integer
   Dim tNode As Node
   Dim pNode As Node
   
   intSeq = 1
   strSQL = ""
   
   Set tNode = tv1.Nodes(1).Next
   Do Until tNode Is Nothing
      strSQL = strSQL & _
               "UPDATE Request_Panels SET " & _
               "Sequence = " & intSeq & _
               " WHERE PanelID = " & objTV.NodeLevel(tNode) & "; "
      intSeq = intSeq + 1

      Set tNode = tNode.Next
   Loop
   iceCon.Execute strSQL
   tv1.Nodes.Clear
   Panels
   tv1.Visible = True
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadPanels.WritePanelSequence"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

