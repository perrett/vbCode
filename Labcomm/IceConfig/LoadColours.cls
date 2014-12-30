VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadColours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private dForm As String
Private newNode As Node
Private nd(2) As Node
Dim vData As Variant
Private tv1 As TreeView
Private plist As PropertiesList
Private mCtrl As New ManageControls

Private Sub Class_Initialize()
   Dim RS As New ADODB.Recordset
   
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   objTV.Origin = "COL"
   
   strSQL = "SELECT dateformat " & _
            "FROM master.dbo.syslanguages " & _
            "WHERE langid = @@default_langid"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS!DateFormat = "DMY" Then
      dForm = "dd/mm/yyyy"
   Else
      dForm = "yyyy/mm/dd"
   End If
   
   RS.Close
   Set RS = Nothing
   With plist
      .Pages.Clear
      .PropertyItems.Clear
      .Pages.Add "Main", "Colours"
      .PropertyItems.Add "INDEX", "Hidden", plpsNumber, , "Hidden"
      .PropertyItems.Add "COL_NAME", "Colour Name", plpsString, , "How is this colour known?"
      .PropertyItems.Add "COL_CODE", "Colour Code", plpsColor, , "Numeric value for the colour"
      .PropertyItems.Add "DATE", "Date Added", plpsString, , "The date this colour was added to the collection"
          
'      With .PropertyItems("DESC")
'         .Tag = "Discipline_Expansion"
'         .DefaultValue = ""
'      End With
      With .PropertyItems("INDEX")
         .Tag = "Colour_Index"
         .Visible = False
      End With
      With .PropertyItems("COL_NAME")
         .Tag = "Colour_Name"
         .defaultValue = ""
      End With
      With .PropertyItems("COL_CODE")
         .Tag = "Colour_Code"
         .defaultValue = 0
      End With
      With .PropertyItems("DATE")
         .Tag = "Date_Added"
         .defaultValue = ""
         .ReadOnly = True
      End With
   End With
   fView.Show Fra_COLOURS
'   frmMain.labShowCol.Top = 6900
'   frmMain.labShowCol.Left = 2730
'   frmMain.labShowCol.BorderStyle = 0
   mCtrl.SetDbStructures plist
End Sub

Public Sub Delete(NodeId As Node)
   Dim thisColour As Integer
   Dim lCount As Integer
   Dim strDets As String
   Dim RS As New ADODB.Recordset
   
   thisColour = objTV.NodeLevel(NodeId)
   strSQL = "SELECT Test_Index, Screen_Caption " & _
            "FROM Request_Tests " & _
            "WHERE Screen_Colour = " & thisColour & _
               " OR Screen_Help_Backcolour = " & thisColour & _
            " ORDER BY Screen_Caption"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount = 0 Then
      RS.Close
      strSQL = "SELECT Tube_Index, Name " & _
               "FROM Request_Tubes " & _
               "WHERE Colour = " & thisColour & _
               " ORDER BY Name"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount = 0 Then
         strSQL = "DELETE FROM Colours " & _
                  "WHERE Colour_Index = " & objTV.NodeLevel(NodeId)
         iceCon.Execute strSQL
         tv1.Nodes.Remove NodeId.Index
      Else
         strDets = "This colour is still in use by the following Tubes: " & vbCrLf & vbCrLf
         Do Until RS.EOF
            strDets = strDets & RS!Name & vbCrLf
            RS.MoveNext
         Loop
         strDets = strDets & vbCrLf & "Please amend the tube(s) to use a different colour."
         MsgBox strDets, vbExclamation, "Unable to delete Colour"
      End If
   Else
      strDets = "This colour is still in use by the following Tests: " & vbCrLf & vbCrLf
      lCount = 0
      Do Until RS.EOF Or (lCount > 20)
         strDets = strDets & RS!Screen_Caption & vbCrLf
         lCount = lCount + 1
         RS.MoveNext
      Loop
      If RS.EOF = False Then
         strDets = strDets & "(... further tests use this colour ...)" & vbCrLf
      End If
      strDets = strDets & vbCrLf & "Please amend the test(s) to use a different colour."
      MsgBox strDets, vbExclamation, "Unable to delete Colour"
   End If
   RS.Close
   Set RS = Nothing
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   objTV.ActiveNode = NodeId
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   strSQL = "SELECT * " & _
            "FROM Colours " & _
            "WHERE Colour_Index = " & objTV.NodeLevel(NodeId)
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
'   Set nd(0) = tv1.Nodes.Add(nodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(RS!Colour_Index, _
                                              "COL_EXAMPLE", _
                                              "Item"), _
                             "         ", _
                             2, _
                             2)
'   nd(0).BackColor = Val(RS!Colour_Code)
   
'   pList.DescriptionPane.ForeColor = Val(RS!Colour_Code)
'   pList.DescriptionPane.BackColor = Val(RS!Colour_Code)
'   frmMain.fraPanel(3).BackColor = Val(RS!Colour_Code)
'   frmMain.OrgList.BackColor = Val(RS!Colour_Code)
   tv1.Nodes.Add NodeId, _
                 tvwChild, _
                 mCtrl.NewNodeKey(RS!Colour_Index, _
                                  RS!Colour_Code, _
                                  "Item"), _
                 "Name - " & RS!Colour_Name, _
                 2, _
                 2
   tv1.Nodes.Add NodeId, _
                 tvwChild, _
                 mCtrl.NewNodeKey(RS!Colour_Index, _
                                  RS!Colour_Code, _
                                  "Item"), _
                 "Code - " & RS!Colour_Code, _
                 2, _
                 2
   tv1.Nodes.Add NodeId, _
                 tvwChild, _
                 mCtrl.NewNodeKey(RS!Colour_Index, _
                                  RS!Colour_Code, _
                                  "Item"), _
                 "Date Added - " & IIf(IsNull(RS!Date_Added), Format("01/01/1900", dForm), Format(RS!Date_Added, dForm)), _
                 2, _
                 2
   
   With plist
      .PropertyItems("INDEX").value = RS!Colour_Index
      .PropertyItems("COL_NAME").value = RS!Colour_Name
      .PropertyItems("COL_CODE").value = Val(RS!Colour_Code)
      .PropertyItems("DATE").value = IIf(IsNull(RS!Date_Added), Format("01/01/1900", dForm), Format(RS!Date_Added, dForm))
      .Caption = "WHERE Colour_Index = " & RS!Colour_Index
   End With
   RS.Close
   fView.Show Fra_EDI
'   fView.Show Fra_COLOURS
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadColours.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewColour()
   Dim i As Integer
   Dim RS As New ADODB.Recordset
   
   With plist
      .Redraw = False
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
      Next i
      .Redraw = True
   End With
   
   If newNode Is Nothing Then
      strSQL = "SELECT MAX(Colour_Index) as MCI " & _
               "FROM Colours"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If IsNull(RS!MCI!) Then
         i = 1
      Else
         i = RS!MCI + 1
      End If
      Set newNode = tv1.Nodes.Add(, _
                                  , _
                                  mCtrl.NewNodeKey(CStr(i), _
                                                   "NewNode", _
                                                   "Details", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  2, _
                                  2)
      plist.PropertyItems("INDEX").value = i
      plist("DATE").value = Format(Now(), dForm)
      RS.Close
   End If
   objTV.ActiveNode = newNode
   objTV.RefreshNode = newNode
   
   Set RS = Nothing
'   fView.Show Fra_EDI
End Sub

Public Function Refresh()
   objTV.ActiveNode = objTV.TopLevelNode(objTV.ActiveNode)
   objTV.ActiveNode.Text = plist("COL_NAME").value
   objTV.ActiveNode.Key = mCtrl.NewNodeKey(objTV.NodeLevel(objTV.ActiveNode), _
                                           plist("COL_CODE").value, _
                                           "Details", _
                                           , _
                                           ms_DELETE)
'   objTView.ActiveNode = objTV.ActiveNode
   tv1.SelectedItem = objTV.ActiveNode
'   tv1.Sorted = True
'   objTV.ActiveNode.Expanded = True
'   objTV.ActiveNode.EnsureVisible
   objTV.RefreshNode = objTV.ActiveNode
'   tv1.Refresh
'   tv1.Sorted = False
End Function

Public Sub RunWhat(NodeId As Node, Optional Refresh As String = "")
   On Error GoTo procEH
   
   If TidyUp = False Then
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "Details"
            Details NodeId
            
         Case "NewColour"
            NewColour
            
         Case "Item"
            Details NodeId.Parent
            
      End Select
   End If
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadColours.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub LoadColours()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   
   fView.FrameToShow = Fra_EDI
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("New", _
                                              "NewNode", _
                                              "NewColour"), _
                             " Add New ", _
                             2, _
                             2)
   nd(0).Bold = True

   strSQL = "SELECT * " & _
            "FROM Colours " & _
            "ORDER BY Colour_Name"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Colour_Index, _
                                                 RS!Colour_Code, _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE), _
                                RS!Colour_Name, _
                                2, _
                                2)
'      nd(0).ForeColor = Val(RS!Colour_Code)
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
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadDisciplineMap.ServiceDisciplines"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function TidyUp() As Boolean
   Dim tNode As Node
   
   TidyUp = False
   Set tNode = tv1.Nodes(1).Next
   If Not newNode Is Nothing Then
      If objTV.nodeKey(newNode) = "NewNode" Then
         tv1.Nodes.Remove newNode.Index
         tv1.SelectedItem = tNode
         Details tNode
         TidyUp = True
      End If
      Set newNode = Nothing
   End If
End Function

Private Sub Class_Terminate()
   plist.DescriptionPane.BackColor = &H8000000F
   plist.DescriptionPane.ForeColor = &H80000012
'   frmMain.labShowCol.Visible = False
End Sub
