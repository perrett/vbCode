VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadReadCodes"
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
Private blnShowReview As Boolean

Private Sub Class_Initialize()
   On Error GoTo procEH
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   mCtrl.TreeViewUpdate = True
   
   blnShowReview = (Read_Ini_Var("GENERAL", "ShowReviewLevel", iniFile) = 1)
   fView.FrameToShow = Fra_EDI
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , "Hidden"

'     Add control items as required
      .PropertyItems.Add "INDEX", "", plpsString, , "Hidden"
'      .PropertyItems.Add "RLC", "", plpsString, , "Hidden"
'      .PropertyItems.Add "RST", "", plpsString, , "Hidden"
'      .PropertyItems.Add "RUBRIC", "", plpsString, , "Hidden"

'     Add pages as required
      .Pages.Add "Main", "EDI_InvTest_Codes"
      
'     Set up the required property items
      .PropertyItems.Add "LR", "Local Rubric", plpsString, , "The local description code"
      .PropertyItems.Add "EC", "EDI_Read_Code", plpsCustom, , "The national read code", "RS!EDI_Read_Code"
      .PropertyItems.Add "ST", "EDI Sample type Code", plpsCustom, , "The national specimen code"
      .PropertyItems.Add "NATDESC", "EDI Sample Description", plpsString, , "The National Specmen Description"
      .PropertyItems.Add "LT", "Local Test Code", plpsString, , "The local test code"
      .PropertyItems.Add "SP", "Suppress Output", plpsBoolean, , "Suppress the output of any tests using this code"
      .PropertyItems.Add "AC", "Checked", plpsBoolean, , "Has this read code been checked?"
      
      If blnShowReview Then
         .PropertyItems.Add "REV", "Review Level", plpsNumber, 1, "Show in Ward Reporting"
      End If
      
   End With
   
   With plist
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 1

'     Set other Control items to the page key
      .PropertyItems("INDEX").PageKeys = "Control"
'      .PropertyItems("RLC").PageKeys = "Control"
'      .PropertyItems("RST").PageKeys = "Control"
'      .PropertyItems("RUBRIC").PageKeys = "Control"
      
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
      With .PropertyItems("LR")
         .Tag = "EDI_Local_Rubric"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("EC")
         .Tag = "EDI_Read_Code"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("ST")
         .Tag = "EDI_Sample_TypeCode"
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("NATDESC")
         .ReadOnly = True
         .DefaultExtension = "Item"
      End With
      With .PropertyItems("LT")
         .Tag = "EDI_Local_Test_Code"
         .DefaultExtension = "Item"
         .ReadOnly = True
      End With
      With .PropertyItems("SP")
         .DefaultExtension = "Item"
         .Tag = "EDI_OP_Suppress"
      End With
      With .PropertyItems("AC")
         .DefaultExtension = "Item"
         .Tag = "EDI_OP_Active"
      End With
      
      If blnShowReview Then
         With .PropertyItems("REV")
            .DefaultExtension = "Item"
            .Tag = "Review_Level"
            .Min = 0
            .max = 999
            .defaultValue = Null
            .PageKeys = "Map"
         End With
      End If
   End With

'  Validate fields and set max length for string values
   mCtrl.SetDbStructures plist
   plist.Pages("Main").Caption = "ICECONFIG_Amend_Invtest_Codes"
   fView.FrameToShow = Fra_EDI
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadReadCodes.Class.Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Headers()
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("N", _
                                              "NoReadCode", _
                                              "FirstView"), _
                             "Non-Read-Coded Rubrics", _
                            plist("ICON").Icon, _
                            plist("ICON").Icon)
   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", "Tmp"), _
                 "Please wait...", _
                plist("ICON").Icon, _
                plist("ICON").Icon
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("Y", "ReadCode", "ReadCodes"), _
                             "Read-Coded Rubrics", _
                            plist("ICON").Icon, _
                            plist("ICON").Icon)
   tv1.Nodes.Add nd(0), _
                 tvwChild, _
                 mCtrl.NewNodeKey("Tmp", "Tmp"), _
                 "Please wait...", _
                plist("ICON").Icon, _
                plist("ICON").Icon
End Sub

Public Property Get KeyControl() As ManageControls
   Set KeyControl = mCtrl
End Property

Public Sub FirstView(Optional NodeId As Node = Nothing)
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim sampTxt As String
   
'  SQL string to extract top level branches for node
   tv1.Visible = False
   tv1.Nodes.Clear
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("N", _
                                              "NoReadCode", _
                                              "None"), _
                             "Non-Read-Coded Rubrics", _
                            plist("ICON").Icon, _
                            plist("ICON").Icon)
   strSQL = "SELECT * FROM EDI_InvTest_Codes " & _
            "WHERE (EDI_Read_Code = '' OR EDI_Read_Code=null) " & _
            "ORDER BY EDI_Local_Rubric"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      Set nd(1) = tv1.Nodes.Add(nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!EDI_Index, _
                                                 Trim(RS!EDI_Local_Test_Code), _
                                                 "Detail", _
                                                 , _
                                                 ms_DELETE), _
                                RS!EDI_Local_Rubric & "(" & Replace(RS!EDI_Local_Test_Code, Chr(0), "") & ")", _
                                1, _
                                1)
      
'     Colour code the node?
      If RS!EDI_Op_Active Then
         nd(1).ForeColor = BPBLUE
      Else
         nd(1).ForeColor = BPRED
      End If
      
      tv1.Nodes.Add nd(1), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    2, _
                    2
      RS.MoveNext
   Loop
   RS.Close
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("Y", "ReadCode", "NONE"), _
                             "Read-Coded Rubrics", _
                             plist("ICON").Icon, _
                             plist("ICON").Icon)
   
   strSQL = "SELECT * FROM EDI_InvTest_Codes WHERE EDI_Read_Code <> '' " & _
            "AND EDI_Read_Code is not null " & _
            "ORDER BY EDI_Local_Rubric"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      Set nd(1) = tv1.Nodes.Add(nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!EDI_Index, _
                                                 Trim(RS!EDI_Local_Test_Code), _
                                                 "Detail", _
                                                 , _
                                                 ms_DELETE), _
                                RS!EDI_Local_Rubric & "(" & Replace(RS!EDI_Local_Test_Code, Chr(0), "") & ")", _
                                1, _
                                1)
      
'     Colour code the node?
      If RS!EDI_Op_Active Then
         nd(1).ForeColor = BPBLUE
      Else
         nd(1).ForeColor = BPRED
      End If
      
      tv1.Nodes.Add nd(1), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                    2, _
                    2
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadReadCodes.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
      
   vData = objTV.ReadNodeData(NodeId)
   
   strSQL = "DELETE FROM EDI_InvTest_Codes " & _
            "WHERE EDI_Index = " & vData(0) & _
               "AND Organisation='" & frmMain.OrgList.Text & "' "
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadReadCodes.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node, _
                   Refresh As Boolean)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   Dim RS As New ADODB.Recordset
   
   If Refresh Then
      vData = objTV.ReadNodeData(NodeId)
   
'  Set Control values
      plist("INDEX").value = vData(0)
'      plist("RLC").value = vData(0)
'      plist("RST").value = vData(1)
'      plist("RUBRIC").value = vData(4)
   End If
   
'  SQL string to be executed
   objTV.SQL = "SELECT * " & _
               "FROM EDI_InvTest_Codes " & _
               "WHERE Organisation LIKE '" & frmMain.OrgList.Text & "' " & _
                  "AND EDI_Index = " & plist("INDEX").value
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.Child.Index
      Next i
      objTV.UpdateTreeView plist, "Main", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Main", mCtrl, NodeId
   End If
      
'  Specific node actions eg set active colour
   Set nd(0) = objTV.NodeFromKey("SP", NodeId)
   If plist("SP").value = True Then
      NodeId.Image = 27
      nd(0).Image = 3
      NodeId.SelectedImage = 27
   Else
      NodeId.Image = 2
      NodeId.SelectedImage = 2
   End If

   strSQL = "SELECT * " & _
            "FROM CRIR_Sample_Type " & _
            "WHERE Sample_Code = '" & plist("ST").value & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.EOF = False Then
      plist("NATDESC").value = Trim(RS!Sample_Text & "")
   End If
   
   RS.Close
   Set RS = Nothing
      
   Set nd(0) = objTV.NodeFromKey("AC", NodeId)
   nd(0).EnsureVisible
   If plist("AC").value = True Then
      NodeId.ForeColor = BPBLUE
   Else
      NodeId.ForeColor = BPRED
   End If
   plist.Caption = " WHERE EDI_Index = " & plist("INDEX").value & _
                     " AND Organisation='" & frmMain.OrgList.Text & "' "

'   fView.Show Fra_EDI
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadReadCodes.Details"
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
                                  mCtrl.NewNodeKey("New", _
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
   eClass.CurrentProcedure = "IceConfig.LoadReadCodes.MenuAddEntry"
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
'   objTView.ActiveNode = newNode
'   fView.Show Fra_EDI
   
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadReadCodes.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function Refresh() As String
   Dim thisKey As String
   Dim strStart As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = True

'  Specific refresh activities
   If objTV.ActiveNode.Children = 0 Then
      objTV.ActiveNode = objTV.ActiveNode.Parent
   End If
   
   If Not objTV.ActiveNode.Parent Is Nothing Then
      If objTV.ActiveNode.Parent.Text = "Read-Coded Rubrics" Then
         strStart = "Non-Coded"
      Else
         strStart = "Coded"
      End If
   End If
   
   thisKey = objTV.ActiveNode.Key
   
   FirstView
   If strStart = "Coded" Then
      objTV.RefreshNode = tv1.Nodes(1)
   Else
      objTV.RefreshNode = tv1.Nodes(1).Next
   End If
   objTV.ActiveNode = objTV.RefreshNode
   fView.Show Fra_HELP
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
   
   fView.FrameToShow = Fra_EDI
   If TidyUp = False Then
'      tv1.Visible = False
      vData = objTV.ReadNodeData(NodeId)
      Select Case vData(2)
         Case "FirstView"
            FirstView NodeId
         
         Case "Detail"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId
            Details NodeId, Refresh
            mCtrl.TreeViewUpdate = False
         
         Case "NewMapping"
            NewMapping
         
'         Case "SubNodes"
'            objTV.RefreshNode = NodeId.Parent
'            SubNodes NodeId
            
         Case "Item"
            objTV.RefreshNode = NodeId.Parent
            Details NodeId.Parent, Refresh
            
'         Case "SubHeader"
'            mCtrl.TreeViewUpdate = True
'            objTV.RefreshNode = NodeId.Parent
'            SubHeader NodeId
'            mCtrl.TreeViewUpdate = False
'
'         Case "SubItem"
'            objTV.RefreshNode = NodeId.Parent.Parent
'            SubHeader NodeId.Parent
'
         Case Else
            fView.FrameToShow = Fra_HELP

      End Select
   End If
   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadReadCodes.RunWhat"
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

Public Function Update(PageId As String) As String
   On Error GoTo procEH
   Dim pl As PropertiesList
   Dim iceCmd As New ADODB.Command
   Dim newValue As String
   Dim intRet As Long
   Dim failInfo As String
   Dim strArray() As String
   Dim rLev As Integer
   
'  Run a stored procedure or some special SQL
   With iceCmd
      .ActiveConnection = iceCon
      .CommandText = "ICECONFIG_Amend_InvTest_Codes"
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
      .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , (objTV.newNode))
      .Parameters.Append .CreateParameter("Index", adInteger, adParamInput, , plist("INDEX"))
      .Parameters.Append .CreateParameter("Org", adVarChar, adParamInput, 6, frmMain.OrgList.Text)
      .Parameters.Append .CreateParameter("LocalCode", adVarChar, adParamInput, 10, plist("LT").value)
      .Parameters.Append .CreateParameter("Rubric", adVarChar, adParamInput, 30, plist("LR").value)
      .Parameters.Append .CreateParameter("ReadCode", adVarChar, adParamInput, 6, plist("EC").value)
      .Parameters.Append .CreateParameter("SpecCode", adVarChar, adParamInput, 15, plist("ST").value)
      .Parameters.Append .CreateParameter("OpUOM", adVarChar, adParamInput, 15, "<")
      .Parameters.Append .CreateParameter("OPSuppress", adBoolean, adParamInput, , 0)
      .Parameters.Append .CreateParameter("Active", adBoolean, adParamInput, , plist("AC").value)
      .Parameters.Append .CreateParameter("RevLev", adInteger, adParamInput)
      
      rLev = 0
      
      If blnShowReview Then
         With plist("REV")
            If .value > 0 Then
               If .value < 100 Or .value > 999 Then
                  MsgBox "Review level must be between 100 and 999" & vbCrLf & vbCrLf & _
                         "It has been reset to 0", vbExclamation, "Invalid Review_Level value"
               Else
                  rLev = .value
               End If
            End If
         End With
      End If
      .Parameters("RevLev").value = rLev
      
      .Execute
      intRet = .Parameters("Return")
'      failInfo = "ICECONFIG_Amend_InvTest_Codes (" & .Parameters("Failure") & ")"
      If intRet <> 0 Then
         eClass.FurtherInfo = intRet & " returned from Stored Procedure ICECONFIG_Amend_InvTest_Codes"
      End If
'            newValue = pl("LC").value & "|" & pl("ST").value
   End With
   fView.FrameToShow = Fra_HELP
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadReadCodes.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function



