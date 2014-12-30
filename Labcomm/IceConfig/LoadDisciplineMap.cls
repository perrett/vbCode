VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadDisciplineMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private newNode As Node
Private nd(2) As Node
Dim vData As Variant
Private tv1 As TreeView
Private plist As PropertiesList
Private mCtrl As New ManageControls

Private Sub Class_Initialize()
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   fView.FrameToShow = Fra_EDI
   
   With plist
      .Caption = ""
      .Pages.Clear
      .PropertyItems.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "The Icon to use", plpsString, , "Hidden"
      .Pages.Add "Main", "Service_Discipline_Map"
      .PropertyItems.Add "DMSPEC", "National Specialty Code", plpsCustom, , "The National Korner Code"
      .PropertyItems.Add "NATDESC", "National Specialty Description", plpsString, , "(Read Only) The National Description this discipline maps to."
      .PropertyItems.Add "DESC", "Local Description", plpsString, , "The Local Description"
      .PropertyItems.Add "DCODE", "Discipline Code", plpsString, , "Mandatory - The Discipline Code"
      .PropertyItems.Add "TEXT", "Discipline Text", plpsString, , "The Local Discipline description"
      
      With .PropertyItems("ICON")
         .PageKeys = "Control"
         .Icon = 7
      End With
      With .PropertyItems("TEXT")
         .Tag = "Discipline_Text"
         .defaultValue = ""
      End With
      With .PropertyItems("DCODE")
         .Tag = "Specialty_Code"
         .defaultValue = ""
      End With
      With .PropertyItems("DESC")
         .Tag = "Discipline_Expansion"
         .defaultValue = ""
      End With
      With .PropertyItems("NATDESC")
         .PageKeys = "Main"
         .ReadOnly = True
      End With
   End With
   plist.ActivePage = "Main"
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
End Sub

Public Sub Delete(NodeId As Node)
   strSQL = "DELETE FROM Service_Discipline_Map " & _
            "WHERE Discipline_Index = " & objTV.NodeLevel(NodeId)
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   objTV.ActiveNode = NodeId
   vData = objTV.ReadNodeData(NodeId)
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   strSQL = "SELECT * " & _
            "FROM Service_Discipline_Map " & _
            "WHERE Discipline_Index = " & vData(0)
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly

   tv1.Nodes.Add NodeId, _
                 tvwChild, _
                 mCtrl.NewNodeKey(RS!Discipline_Index, _
                                  Trim(RS!Specialty_Code & ""), _
                                  "Item"), _
                 "Local Specialty Code - " & RS!Specialty_Code, _
                 plist("ICON").Icon, _
                 plist("ICON").Icon
   tv1.Nodes.Add NodeId, _
                 tvwChild, _
                 mCtrl.NewNodeKey(RS!Discipline_Index, _
                                  Trim(RS!Specialty_Code & ""), _
                                  "Item"), _
                 "Local discipline text - " & Trim(RS!Discipline_Text & ""), _
                 plist("ICON").Icon, _
                 plist("ICON").Icon
   
   With plist
      .PropertyItems("DESC").value = Trim(RS!Discipline_Expansion & "")
      .PropertyItems("TEXT").value = Trim(RS!Discipline_Text & "")
      .PropertyItems("DCODE").value = Trim(RS!Specialty_Code & "")
      .Caption = " WHERE Discipline_Index = " & vData(0)
   End With
   RS.Close
   
   strSQL = "SELECT * " & _
            "FROM CRIR_Specialty " & _
            "WHERE Specialty_Code = '" & Left(objTV.nodeKey(NodeId), 3) & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.EOF = False Then
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!Specialty, _
                                     "NATDESC", _
                                     "Item"), _
                    "National Description - " & RS!Specialty, _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!Specialty, _
                                     "NATCODE", _
                                     "Item"), _
                    "National (Korner) Code - " & RS!Specialty_Code, _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
   
      plist("NATDESC").value = RS!Specialty
      plist("DMSPEC").value = Left(RS!Specialty_Code, 3)
   Else
   
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey("NA", _
                                     "NATDESC", _
                                     "Item"), _
                    "National Description - <Not assigned>", _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey("NA", _
                                     "DMSPEC", _
                                     "Item"), _
                    "National (Korner) Code - <Not assigned>", _
                    plist("ICON").Icon, _
                    plist("ICON").Icon
   
      plist("NATDESC").value = ""
      plist("DMSPEC").value = ""
   End If
   RS.Close
   fView.Show Fra_EDI
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadDisciplineMap.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewDiscipline()
   Dim i As Integer
   Dim RS As New ADODB.Recordset
   
   If newNode Is Nothing Then
      strSQL = "SELECT Max(Discipline_Index) AS SDI " & _
               "FROM Service_Discipline_Map"
      RS.Open strSQL, iceCon
      Set newNode = tv1.Nodes.Add(, _
                                  , _
                                  mCtrl.NewNodeKey(RS!SDI + 1, _
                                                   "NewNode", _
                                                   "Details", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  plist("ICON").Icon, _
                                  plist("ICON").Icon)
      RS.Close
   End If
'   objTView.ActiveNode = newNode
   objTV.ActiveNode = newNode
   
   With plist
      .Redraw = False
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
      Next i
      .Redraw = True
   End With
   plist.ActivePage = "Main"
   fView.Show Fra_EDI
   Set RS = Nothing
End Sub

Public Function Refresh()
   objTV.ActiveNode = objTV.TopLevelNode(objTV.ActiveNode)
   objTV.ActiveNode.Text = plist("DESC").value & " - " & plist("TEXT").value
   objTV.ActiveNode.Key = mCtrl.NewNodeKey(objTV.NodeLevel(objTV.ActiveNode), _
                                           plist("DMSPEC").value, _
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

Public Sub PB_Click(prop As PropertyItem)
   frmSpecs.CalledFrom = "NATCODE"
   frmSpecs.EDI_WriteBack = "DMSPEC"
   frmSpecs.Show 1
   frmSpecs.EDI_WriteBack = ""
   frmSpecs.CalledFrom = ""
   Unload frmSpecs
   plist("DCODE").value = Left(plist("DMSPEC").value, 3)
End Sub

Public Sub RunWhat(NodeId As Node, Optional Refresh As String = "")
   On Error GoTo procEH
   
   If TidyUp = False Then
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "Details"
            Details NodeId
            
         Case "NewDiscipline"
            NewDiscipline
            
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
   eClass.CurrentProcedure = "IceConfig.LoadDiscipline<ap.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub ServiceDisciplines()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("New", _
                                              "NewNode", _
                                              "NewDiscipline"), _
                             "Add New Discipline Map", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   nd(0).Bold = True

   strSQL = "SELECT * " & _
            "FROM Service_Discipline_Map " & _
            "ORDER BY Discipline_Expansion"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Discipline_Index, _
                                                 Trim(RS!Specialty_Code & ""), _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE), _
                                Trim(RS!Discipline_Expansion & "") & " - " & Trim(RS!Discipline_Text & ""), _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Temp", "Temp"), _
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
   eClass.CurrentProcedure = "IceConfig.LoadDisciplineMap.ServiceDisciplines"
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

Public Function Validate(NewCode As String) As Boolean
   Dim RS As New ADODB.Recordset

   strSQL = "SELECT * " & _
            "FROM CRIR_Specialty " & _
            "WHERE Specialty_Code = '" & Left(NewCode, 3) & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount = 0 Then
      MsgBox "The first 3 digits of the local Discipline MUST equate to a valid national code", vbInformation, _
             "Discipline map validation"
      Validate = True
   Else
      plist("DMSPEC").value = Left(NewCode, 3)
      plist("NATDESC").value = RS!Specialty
      Validate = False
   End If
   RS.Close
   Set RS = Nothing
End Function

Private Sub Class_Terminate()
   plist.Caption = ""
End Sub


