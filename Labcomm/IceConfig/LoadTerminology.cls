VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadTerminology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private strSQL As String
Private vData As Variant
Private mCtrl As New ManageControls
Private nd(2) As Node
Private newNode As Node

Private Sub Class_Initialize()
   Set tv1 = frmMain.TreeView1
   objTV.Origin = "Term"
   fView.FrameToShow = Fra_HELP
End Sub

Public Sub FirstView()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim tNode As Node
   
   tv1.Nodes.Clear
   Set newNode = Nothing
'   Set newNode = Nothing
   Set tNode = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("CAT", _
                                              "NONE", _
                                              "Categories", _
                                              , _
                                              ms_ADD), _
                             "Category List", _
                             1, _
                             1)
'   Set nd(0) = tv1.Nodes.Add(tNode, _
'                             tvwChild, _
'                             mCtrl.NewNodeKey("CAT", _
'                                              "NONE", _
'                                              "NewEntry"), _
'                             "Add New Category", _
'                             1, _
'                             1)
'
'   nd(0).Bold = True
'
   strSQL = "SELECT * " & _
            "FROM Request_Category"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      tv1.Nodes.Add tNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey("CAT", _
                                     "EDIT", _
                                     , _
                                     , _
                                     ms_DELETE), _
                    RS!Category, _
                    1, _
                    1
      RS.MoveNext
   Loop
   RS.Close
   Set tNode = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("PRIORITY", _
                                              "NONE", _
                                               "Priorities", _
                                               , _
                                               ms_ADD), _
                             "Priority List", _
                             1, _
                             1)
   
'   Set nd(0) = tv1.Nodes.Add(tNode, _
'                             tvwChild, _
'                             mCtrl.NewNodeKey("PRIORITY", _
'                                              "NONE", _
'                                              "NewEntry"), _
'                             "Add New Priority", _
'                             1, _
'                             1)
'
'   nd(0).Bold = True
   
   strSQL = "SELECT * " & _
            "FROM Request_Priority"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      tv1.Nodes.Add tNode, _
                    tvwChild, _
                    mCtrl.NewNodeKey("PRIORITY", _
                                     "EDIT", _
                                     , _
                                     , _
                                     ms_DELETE), _
                    RS!Priority, _
                    1, _
                    1
      RS.MoveNext
   Loop
   If Not (objTV.ActiveNode Is Nothing) Then
      vData = objTV.ReadNodeData(objTV.ActiveNode)
      If vData(0) = "CAT" Then
         Set tNode = tv1.Nodes(1)
      Else
         Set tNode = tv1.Nodes(1).Next
      End If
      tNode.Expanded = True
      Set tNode = tNode.child
      Do Until tNode Is Nothing
         If tNode.Text = objTV.ActiveNode.Text Then
            Set tv1.SelectedItem = tNode
            Exit Do
         End If
         Set tNode = tNode.Next
      Loop
   End If
   RS.Close
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadTerminology.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   vData = objTV.ReadNodeData(NodeId)
   If vData(0) = "CAT" Then
      strSQL = "DELETE FROM Request_Category " & _
               "WHERE Category = '" & NodeId.Text & "'"
   Else
      strSQL = "DELETE FROM Request_Priority " & _
               "WHERE Priority = '" & NodeId.Text & "'"
   End If
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   objTV.ActiveNode = Nothing
End Sub

Public Sub MenuAddEntry()
   If newNode Is Nothing Then
      vData = objTV.ReadNodeData(objTV.ActiveNode)
      objTV.ActiveNode.Expanded = True
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(CStr(vData(0)), _
                                                   "EDIT", _
                                                   , _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New...", _
                                  1, _
                                  1)
   End If
   objTV.ActiveNode = newNode
   Set tv1.SelectedItem = newNode
   tv1.StartLabelEdit
End Sub

Public Sub Refresh()
   FirstView
End Sub

Public Sub RunWhat(NodeId As Node, _
                   Optional Refresh As String = "")
   On Local Error GoTo procEH
   Dim vData As Variant
   Dim lCode As String
   Dim tNode As MSComctlLib.Node
   Dim RefreshId As String
   Dim strArray() As String
   
   If TidyUp = False Then
'      vData = objTV.ReadNodeData(NodeId)
'      Select Case vData(2)
'         Case "NewEntry"
'            NewEntry NodeId
'
'      End Select
   End If
   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadTerminology.RunWhat"
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
'      Set newNode = Nothing
   End If
End Function


