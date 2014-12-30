VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadPickLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private nd(2) As Node
Private newNode As Node
Private plist As PropertiesList
Private vData As Variant
Private mCtrl As New ManageControls
Private thisIndex As Long

Private Sub Class_Initialize()
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   fView.FrameToShow = Fra_PICKLIST
   plist.PropertyItems.Clear
   plist.PropertyItems.Add "PICK", , plpsString, , "Hidden"
End Sub

Public Sub Delete(NodeId As Node)
   Dim strSQL As String
   Dim vData As Variant
   Dim ruleStr As String
   Dim RS As New ADODB.Recordset
   
   vData = objTV.ReadNodeData(NodeId)
   If vData(2) = "Details" Then
      strSQL = "SELECT Prompt_Text, Prompt_Desc " & _
               "FROM Request_Prompt " & _
               "WHERE Picklist_Index = " & vData(0)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount > 0 Then
         ruleStr = "This picklist is used in the following Data Entry rule(s):" & vbCrLf & vbCrLf
         Do Until RS.EOF
            ruleStr = ruleStr & RS!Prompt_Text & " (" & RS!Prompt_Desc & ")" & vbCrLf
            RS.MoveNext
         Loop
         RS.Close
         Set RS = Nothing
         ruleStr = ruleStr & vbCrLf & "Please amend the rules to use another picklist"
         MsgBox ruleStr, vbOKOnly, "Picklist in use"
         Exit Sub
      End If
      
      strSQL = "DELETE FROM Request_Picklist " & _
               "WHERE Picklist_Index = " & vData(0)
   ElseIf vData(2) = "Entries" Then
      strSQL = "DELETE FROM Request_Picklist_Data " & _
               "WHERE Picklist_Index = " & vData(0) & _
                  "AND Picklist_Value = '" & NodeId.Text & "'"
   
   End If
   If strSQL <> "" Then
      iceCon.Execute strSQL
      frmMain.TreeView1.Nodes.Remove NodeId.Index
   End If
   Set tv1.SelectedItem = Nothing
End Sub

Private Sub Details(NodeId As Node, _
                    Refresh As Boolean)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   vData = objTV.ReadNodeData(NodeId)
   With frmMain
      .fraPickOpt.Tag = "Details"
      .cmdPickOK.Caption = "Update"
      strSQL = "SELECT * " & _
               "FROM Request_Picklist " & _
               "WHERE Picklist_Index = " & vData(0) 'Format(Index)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      .Text5.Text = RS!PickList_Name
      If RS!Multichoice Then
         .Check1.value = vbChecked
      Else
         .Check1.value = vbUnchecked
      End If
      RS.Close
         
'         .SSListBar1.Enabled = True
      .fraPickOpt.Visible = True
      .Check1.Visible = True
   End With
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadPickLists.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Entries()
   With frmMain
      .fraPickOpt.Tag = "Entries"
      .fraPickOpt.Caption = "Picklist Entry"
      .Text5.Text = tv1.SelectedItem.Text
      .cmdPickOK.Caption = "Update"
      .Check1.Visible = False
      .fraPickOpt.Visible = True
   End With
End Sub

Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim PickId As String
   Dim thisNode As Node
   Dim vData As Variant
   Dim i As Integer
   
   With frmMain
      .Check1.Visible = False
      .fraPickOpt.Tag = "Entries"
      .Text5.Text = ""
      .cmdPickOK.Caption = "Add Picklist Value"
      .fraPickOpt.Visible = True
   End With
   Set thisNode = objTV.TopLevelNode(objTV.ActiveNode)
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(thisNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(objTV.NodeLevel(objTV.ActiveNode), _
                                                   "NewNode", _
                                                   "Entries", _
                                                   , _
                                                   ms_DELETE), _
                                  "New", _
                                  1, _
                                  1)
   End If
   objTV.ActiveNode = newNode
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadPickLists.MenuAddEntry"
   eClass.Add Err.Number, Err.Description, Err.Source

End Sub

Public Sub NewPickList()
   Set newNode = tv1.Nodes.Add(, _
                               , _
                               mCtrl.NewNodeKey("New", _
                                                "NewNode", _
                                                "Details", _
                                                , _
                                                ms_BOTH, _
                                                True), _
                              "New", _
                              1, _
                              1)
   objTV.ActiveNode = newNode
   
   With frmMain
      .cmdPickOK.Caption = "New Picklist"
      .fraPickOpt.Tag = "Details"
      .fraPickOpt.Visible = True
      .cmdPickOK.Caption = "Add Picklist"
      .Text5.Text = ""
   End With
End Sub
'
Public Sub FirstView()
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset

'  Set up the Database field, Data and PropertiesList objects

   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("NewList", _
                                              "New", _
                                              "NewPickList", _
                                              , _
                                              , True), _
                             "New Picklist", _
                             1, _
                             1)

   nd(0).Bold = True
   strSQL = "SELECT * " & _
            "FROM Request_Picklist " & _
            "ORDER BY Picklist_Name"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      Do While Not RS.EOF
         Set nd(0) = tv1.Nodes.Add(, _
                                   , _
                                   mCtrl.NewNodeKey(RS!Picklist_Index, _
                                                    RS!PickList_Name, _
                                                    "Details", _
                                                    , _
                                                    ms_BOTH), _
                                   RS!PickList_Name, _
                                   1, _
                                   1)

         strSQL = "SELECT * " & _
                  "FROM Request_Picklist_Data " & _
                  "WHERE Picklist_Index = " & Format(RS!Picklist_Index) & _
                  " ORDER BY Picklist_Value"
         RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         If RS2.RecordCount > 0 Then
            Do While Not RS2.EOF
               tv1.Nodes.Add nd(0), _
                             tvwChild, _
                             mCtrl.NewNodeKey(RS!Picklist_Index, _
                                              RS2!Picklist_Value, _
                                              "Entries", _
                                              , _
                                              ms_DELETE), _
                             RS2!Picklist_Value, _
                             1, _
                             1
               RS2.MoveNext
            Loop
         End If
         RS2.Close
         RS.MoveNext
      Loop
   End If
   RS.Close
   Set RS2 = Nothing
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadPickLists.PickLists"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function Refresh(Index As String, _
                        mStatus As Long) As String
'   vData = objTV.ReadNodeData(objTV.ActiveNode)
   objTV.ActiveNode.Text = frmMain.Text5.Text
   objTV.ActiveNode.Key = mCtrl.NewNodeKey(Index, _
                                           frmMain.Text5.Text, _
                                           frmMain.fraPickOpt.Tag, _
                                           , _
                                           mStatus)
   objTV.RefreshNode = objTV.ActiveNode
End Function

Public Sub RunWhat(NodeId As MSComctlLib.Node, _
                   Optional Refresh As Boolean)
   On Error GoTo procEH
   Dim vData As Variant
   
   If TidyUp = False Then
      fView.FrameToShow = Fra_PICKLIST
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "Details"
            Details NodeId, Refresh
'            If IsNumeric(Refresh) Then
'               Details Refresh
'            Else
'               Details CStr(vData(0))
'            End If
         
         Case "Entries"
            Entries
            
         Case "NewPickList"
            NewPickList
   
      End Select
'      fView.Show Fra_PICKLIST
   End If
Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadPRofiles.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let PickIndex(lngNewValue As Long)
   thisIndex = lngNewValue
End Property

Public Function TidyUp() As Boolean
   TidyUp = False
   If Not newNode Is Nothing Then
      If objTV.nodeKey(newNode) = "NewNode" Then
         If tv1.SelectedItem.Text = "New" Then
            tv1.Nodes.Remove newNode.Index
            objTV.ActiveNode = tv1.Nodes(1).LastSibling
         Else
            tv1.Nodes.Remove newNode.Index
            objTV.ActiveNode = tv1.SelectedItem
         End If
      End If
      vData = objTV.ReadNodeData(objTV.ActiveNode)
      If vData(2) = "Details" Then
         Details objTV.ActiveNode, True
      Else
         Entries
      End If
      TidyUp = True
'         Set tv1.SelectedItem = objTV.ActiveNode
      End If
      Set newNode = Nothing
'   End If
End Function
