VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadResultMapping"
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
Private fvSQL As String
Private ltSQL As String
Private fStatus As String

Private Sub Class_Initialize()
   On Error GoTo procEH
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   fView.FrameToShow = Fra_EDI
   mCtrl.TreeViewUpdate = True
   
   blnShowReview = (Read_Ini_Var("GENERAL", "ShowReviewLevel", iniFile) = 1)
   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "ICON", "Standard Icon to use", plpsString, , , 1
      .PropertyItems.Add "NATDESC", , plpsString, , "For compatabilty with the frmSampCodes code"
      .PropertyItems.Add "INDEX", "The read code ", plpsString, , "Hidden"
      .Pages.Add "Map", "EDI_InvTest_Codes"
      .PropertyItems.Add "ORG", "Not Shown", plpsString, , "Organisation"
      .PropertyItems.Add "LC", "Local Code", plpsString, , "The local code for this Test"
      .PropertyItems.Add "LD", "Local Description", plpsString, , "The local description for this Test"
      .PropertyItems.Add "RC", "Read Code", plpsCustom, , "The read code for this Test"
      .PropertyItems.Add "ST", "Specimen Code", plpsCustom, , "The sample code for this Test"
      .PropertyItems.Add "UM", "Measurement Unit", plpsString, , "The Unit of Measurement"
      .PropertyItems.Add "SP", "Suppress Output", plpsBoolean, , "Suppress output of any tests containing this code?"
'      .PropertyItems.Add "BINLC", "Binary Local Code", plpsString, , "Hidden"
      
      If blnShowReview Then
         .PropertyItems.Add "REV", "Review Level", plpsList, 1, "Show in Ward Reporting"
      End If
      
      .PropertyItems.Add "AC", "Active", plpsBoolean, , "Is this Result Mapping active?"
   End With
   
   With plist
      .PropertyItems("ICON").PageKeys = "Control"
      .PropertyItems("ICON").Icon = 1
      .PropertyItems("NATDESC").PageKeys = "Control"
      .PropertyItems("INDEX").PageKeys = "Control"
      
      With .PropertyItems("ORG")
         .Tag = "Organisation"
'         .value = frmMain.OrgList.Text
'         .defaultValue = frmMain.OrgList.Text
         .Visible = False
         .PageKeys = "Map"
      End With
      With .PropertyItems("LC")
         .DefaultExtension = "Item"
         .Tag = "EDI_Local_Test_Code"
         .defaultValue = ""
         .PageKeys = "Map"
      End With
      With .PropertyItems("LD")
         .DefaultExtension = "Item"
         .Tag = "EDI_Local_Rubric"
         .defaultValue = ""
         .PageKeys = "Map"
'         .ReadOnly = True
      End With
      With .PropertyItems("RC")
         .DefaultExtension = "Item"
         .Tag = "EDI_Read_Code"
         .defaultValue = ""
         .PageKeys = "Map"
      End With
      With .PropertyItems("ST")
         .DefaultExtension = "Item"
         .Tag = "EDI_Sample_TypeCode"
         .defaultValue = ""
         .PageKeys = "Map"
      End With
      With .PropertyItems("UM")
         .DefaultExtension = "Item"
         .Tag = "EDI_OP_UOM"
         .defaultValue = ""
         .PageKeys = "Map"
      End With
      With .PropertyItems("SP")
         .DefaultExtension = "Item"
         .OverlayIcon = 3
         .Tag = "EDI_OP_Suppress"
         .defaultValue = False
         .PageKeys = "Map"
      End With
      
      If blnShowReview Then
         With .PropertyItems("REV")
            .DefaultExtension = "Item"
            .Tag = "Review_Level"
            .defaultValue = 0
            .PageKeys = "Map"
            With .ListItems
               .Add "Hide in Ward Reporting", 0
               .Add "Show in Ward Reporting", 1
            End With
         End With
      End If
      
      With .PropertyItems("AC")
         .DefaultExtension = "Item"
         .Tag = "EDI_OP_Active"
         .defaultValue = True
         .PageKeys = "Map"
      End With
'      With .PropertyItems("BINLC")
'         .Tag = "EDI_Loc_Code_Binary"
'         .Visible = False
'         .PageKeys = "MAP"
'      End With
      
      .UsePageKeys = True
   End With
   
   fStatus = "Filter..."
   mCtrl.SetDbStructures plist
   plist.Pages("Map").Caption = "IceCONFIG_Amend_InvTest_Codes"
'   frmMain.OrgList.Visible = True
'   frmMain.labOrgList.Visible = True
   frmMain.OrgList.ListIndex = 0
'   frmMain.OrgList_Click
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadResultMapping.Class.Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let FirstViewSQL(strNewValue As String)
   fvSQL = strNewValue
End Property

Public Sub FirstView()
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim sampTxt As String
   Dim rcLoop As Long
   
   tv1.Visible = False
   tv1.Nodes.Clear
   
   strSQL = "SELECT * " & _
            "FROM Read_Version"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("Main", _
                                              "RootNode", _
                                              "None"), _
                             "Version " & RS!Version & " (Updated " & Format(RS!LastAmended, "dd/mm/yyyy") & ")", _
                             9, _
                             9)
   RS.Close
   
'   Set nd(1) = tv1.Nodes.Add(nd(0), _
                             tvwChild, _
                             mCtrl.NewNodeKey("New", _
                                              "NewNode", _
                                              "NewMapping"), _
                             "Add new Result Map...", _
                             37, _
                             37)
'   nd(1).Bold = True
   
   Set nd(1) = tv1.Nodes.Add(nd(0), _
                             tvwChild, _
                             mCtrl.NewNodeKey("Find", _
                                              "FindNode", _
                                              "Find"), _
                             fStatus, _
                             36, _
                             36)
   nd(1).Bold = True
   
   If fStatus = "Clear filter" Then
      nd(1).ForeColor = BPRED
   Else
      nd(1).ForeColor = BPBLUE
   End If
   
   frmWait.Label1.Caption = "Loading Result Map data..."
   frmWait.ProgressBar1.Visible = True
   frmWait.Show
   frmWait.Refresh
   
   strSQL = "SELECT EDI_LTS_Index, EDI_OrgCode, EDI_Msg_Type " & _
            "FROM EDI_Local_Trader_Settings"
            
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   Do Until RS.EOF
      Set nd(1) = tv1.Nodes.Add(nd(0), _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!EDI_LTS_Index, _
                                                 "LTS", _
                                                 "None"), _
                                RS!EDI_OrgCode & " (" & RS!EDI_Msg_Type & ")", _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      
      If fvSQL = "" Then
         strSQL = "SELECT DISTINCT Left(EDI_Local_Rubric,1) As Initial " & _
                  "FROM EDI_InvTest_Codes " & _
                  "WHERE Organisation = '" & frmMain.cboTrust.Text & "' "
'                     "AND EDI_LTS_Index = " & RS!EDI_LTS_Index & _
                  " ORDER BY Initial"
      Else
         strSQL = fvSQL
      End If
      strSQL = strSQL & "AND EDI_LTS_Index = " & RS!EDI_LTS_Index & _
                        " ORDER BY Initial"

      RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      Do Until RS2.EOF
         Set nd(2) = tv1.Nodes.Add(nd(1), _
                                   tvwChild, _
                                   mCtrl.NewNodeKey(RS2!Initial, _
                                                    RS!EDI_LTS_Index, _
                                                    "LocalTests"), _
                                   RS2!Initial & "...", _
                                   plist("ICON").Icon, _
                                   plist("ICON").Icon)
         
         tv1.Nodes.Add nd(2), _
                       tvwChild, _
                       mCtrl.NewNodeKey("Tmp", "Tmp"), _
                       "Please wait...", _
                      plist("ICON").Icon, _
                      plist("ICON").Icon
         
         RS2.MoveNext
      Loop
      
      RS2.Close
      RS.MoveNext
   Loop
   
   RS.Close
   
   Unload frmWait
   nd(0).Expanded = True
   Set RS = Nothing
   tv1.Visible = True
   Exit Sub

procEH:
   Unload frmWait
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadResultMapping.FirstView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub LocalTests(Optional pNode As MSComctlLib.Node = Nothing)
   On Error GoTo procEH
'   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim sampTxt As String
   Dim i As Integer
   Dim rcLoop As Long
   
   vData = objTV.ReadNodeData(pNode)
   
   For i = 1 To pNode.Children
      tv1.Nodes.Remove pNode.child.Index
   Next i
   
   If ltSQL = "" Then
      strSQL = "SELECT ec.*, Sample_Text " & _
               "FROM EDI_InvTest_Codes ec " & _
                  "INNER JOIN EDI_Local_Trader_Settings el " & _
                  "ON ec.EDI_LTS_Index = el.EDI_LTS_Index " & _
                  "LEFT JOIN CRIR_Sample_Type " & _
                  "ON EDI_Sample_TypeCode = Sample_Code " & _
               "WHERE ec.Organisation = '" & frmMain.cboTrust.Text & "' " & _
                  "AND ec.EDI_LTS_index = " & vData(1) & _
                 " AND Left(EDI_Local_Rubric,1) = '" & vData(0) & "' " & _
              "ORDER BY EDI_Local_Rubric"
   Else
      strSQL = ltSQL & _
                 " AND Left(EDI_Local_Rubric,1) = '" & vData(0) & "' " & _
                 "AND ec.EDI_LTS_Index = " & vData(1) & _
              " ORDER BY EDI_Local_Rubric"
      
   End If
   
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.RecordCount > 0 Then
      frmWait.ProgressBar1.max = RS.RecordCount
      frmWait.ProgressBar1.value = 0
   End If
   rcLoop = 0
   
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(pNode, _
                                tvwChild, _
                                mCtrl.NewNodeKey(RS!EDI_Index, _
                                                 Trim(Replace(RS!EDI_Local_Test_Code, Chr(0), "")), _
                                                 "Detail", _
                                                 Trim(RS!EDI_Sample_TypeCode) & Trim(RS!EDI_Local_Rubric), _
                                                 ms_DELETE), _
                                Trim(RS!EDI_Local_Rubric) & " - " & Trim(Replace(RS!EDI_Local_Test_Code, Chr(0), "")) & " (" & RS!Sample_Text & ")", _
                                plist("ICON").Icon, _
                                plist("ICON").Icon)
      If RS!EDI_OP_Suppress Then
         nd(0).Image = 3
         nd(0).SelectedImage = 3
      Else
         nd(0).Image = 1
         nd(0).SelectedImage = 1
      End If
      
      If RS!EDI_Op_Active Then
         If Trim(RS!EDI_Read_Code & "") = "" Then
            nd(0).ForeColor = BPBLUE
         Else
            nd(0).ForeColor = BPGREEN
         End If
      Else
         nd(0).ForeColor = BPRED
      End If
      
'      tv1.Nodes.Add nd(0), _
                    tvwChild, _
                    mCtrl.NewNodeKey("Tmp", "Tmp"), _
                    "Please wait...", _
                   plist("ICON").Icon, _
                   plist("ICON").Icon
      rcLoop = rcLoop + 1
      frmWait.ProgressBar1.value = rcLoop
      RS.MoveNext
   Loop
   RS.Close
   Unload frmWait
   Set RS = Nothing
   Exit Sub

procEH:
   Unload frmWait
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadResultMapping.LocalTests"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let LocalTestsSQL(strNewValue As String)
   ltSQL = strNewValue
End Property

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim strCode As String
   Dim strSpec As String
   Dim strSex As String
   Dim minAge As String
   Dim maxAge As String
   Dim tNode As Node
   
   strSQL = "DELETE FROM EDI_InvTest_Codes " & _
            "WHERE EDI_Index = " & objTV.NodeLevel(NodeId) & "; " & _
            "DELETE FROM EDI_InvTest_Ranges " & _
            "WHERE EDI_Index = " & objTV.NodeLevel(NodeId)
   iceCon.Execute strSQL
   tv1.Nodes.Remove NodeId.Index
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadResultMapping.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
   Dim nText As String
   Dim i As Integer
   Dim iconId As Long
   
   vData = objTV.ReadNodeData(NodeId)
   plist("INDEX").value = vData(0)
'   plist("SPEC").value = vData(1)
'   plist("RUBRIC").value = vData(4)
   
   objTV.SQL = "SELECT * " & _
               "FROM EDI_InvTest_Codes " & _
               "WHERE EDI_Index = " & plist("INDEX").value
   
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Map", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Map", mCtrl, NodeId
   End If
   
'  Specific node actions
      
   Set nd(0) = objTV.NodeFromKey("SP", NodeId)
   If plist("SP").value = True Then
      NodeId.Image = 3
      NodeId.SelectedImage = 3
   Else
      NodeId.Image = 1
      NodeId.SelectedImage = 1
   End If
   
   Set nd(0) = objTV.NodeFromKey("SP", NodeId)
'   nd(0).EnsureVisible
   If plist("AC").value = True Then
      If plist.PropertyItems("RC").value = "" Then
         NodeId.ForeColor = BPBLUE
      Else
         NodeId.ForeColor = BPGREEN
      End If
   Else
      NodeId.ForeColor = BPRED
   End If
   objTV.ResetActive NodeId
'   RS.Close
'   Set RS = Nothing
'   fView.Show Fra_EDI
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadResultMapping.Details"
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
   plist("INDEX").value = vData(0)
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(CStr(vData(0)), _
                                                   "NewNode", _
                                                   "RangeDetails", _
                                                   , _
                                                   , _
                                                   True), _
                                  "New", _
                                 plist("ICON").Icon, _
                                 plist("ICON").Icon)
   End If
   
   objTV.RefreshNode = objTV.ActiveNode
   objTV.ActiveNode = newNode
   
   With plist
'      With .PropertyItems("ORG")
'         .PageKeys = "Range"
'         .Visible = False
'         .value = frmMain.OrgList.Text
'      End With
'      .PropertyItems("RLC").value = vData(0)
'      .PropertyItems("RST").value = vData(1)
      .Pages("Range").Selected = True
   End With
      
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadResultMapping.MenuAddEntry"
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
'      .PropertyItems("ORG").PageKeys = "Map"
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
   
   plist("ORG").value = frmMain.OrgList.Text
   plist.Pages("Map").Selected = True
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
   eClass.CurrentProcedure = "LoadResultMapping.NewMapping"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub RangeNodes(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   vData = objTV.ReadNodeData(NodeId.Parent)
   
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   strSQL = "SELECT * " & _
            "FROM EDI_InvTest_Ranges " & _
            "WHERE EDI_Index = " & vData(0)
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(NodeId, _
                                tvwChild, _
                                mCtrl.NewNodeKey(Trim(RS!EDI_Range_Sex), _
                                                 Trim(RS!EDI_Range_MinAge) & "-" & Trim(RS!EDI_Range_MaxAge), _
                                                 "RangeHeader", _
                                                 , _
                                                 ms_DELETE), _
                                Trim(RS!EDI_Range_Sex) & " " & Trim(RS!EDI_Range_MinAge) & " to " & Trim(RS!EDI_Range_MaxAge), _
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
   eClass.CurrentProcedure = "LoadResultMapping.RangeDetails"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub RangeHeader(NodeId As Node)
   On Error GoTo procEH
   Dim strLoc As String
   Dim strSpec As String
   Dim strMin As String
   Dim strMax As String
   Dim pos As Integer
   Dim i As Integer
   Dim iconId As Long
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   plist("INDEX").value = vData(0)
   vData = objTV.ReadNodeData(NodeId)
   pos = InStr(1, vData(1), "-")
   plist("SEX").value = vData(0)
   plist("MIN").value = Left(vData(1), pos - 1)
   plist("MAX").value = Mid(vData(1), pos + 1)
   
   
   objTV.SQL = "SELECT * " & _
                "FROM EDI_InvTest_Ranges " & _
                "WHERE EDI_Index = '" & plist("INDEX").value & "' " & _
                   "AND EDI_Range_Sex = '" & plist("SEX").value & " ' " & _
                   "AND EDI_Range_MinAge = " & plist("MIN").value & _
                   " AND EDI_Range_MaxAge = " & plist("MAX").value
   If mCtrl.TreeViewUpdate Then
      For i = 1 To NodeId.Children
         tv1.Nodes.Remove NodeId.child.Index
      Next i
      objTV.UpdateTreeView plist, "Range", mCtrl, NodeId
   Else
      objTV.UpdatePropList plist, "Range", mCtrl, NodeId
   End If
   
   plist.Caption = " WHERE EDI_Index = '" & plist("INDEX").value & "' " & _
                     "AND EDI_Range_Sex = '" & plist("SEX").value & " ' " & _
                     "AND EDI_Range_MinAge = " & plist("MIN").value & _
                     " AND EDI_Range_MaxAge = " & plist("MAX").value

   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadResultMapping.RangeHeader"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Public Function Refresh() As String
   Dim tNode As Node
   
   mCtrl.TreeViewUpdate = True
'   With plist
'      .PropertyItems("ORG").PageKeys = "Map"
'      .PropertyItems("ORG").Visible = False
'   End With
   
   If objTV.newNode Then
      objTV.ActiveNode.Key = mCtrl.NewNodeKey("Tmp", "Tmp", , , , False)
   End If
   
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
   
   fView.FrameToShow = Fra_EDI
   blnShowBrowser = False
   
   If TidyUp = False Then
'      tv1.Visible = False
      vData = objTV.ReadNodeData(NodeId)
      
      Select Case vData(2)
         Case "LocalTests"
            objTV.RefreshNode = NodeId
            LocalTests NodeId
            fView.FrameToShow = Fra_HELP
            blnShowBrowser = True
         
         Case "Detail"
            mCtrl.TreeViewUpdate = False
            objTV.RefreshNode = NodeId
            Details NodeId
            mCtrl.TreeViewUpdate = False
         
         Case "NewMapping"
            NewMapping
         
         Case "RangeNodes"
            objTV.RefreshNode = NodeId.Parent
            RangeNodes NodeId
            
         Case "Item"
            objTV.RefreshNode = NodeId.Parent
            Details NodeId.Parent
            
         Case "RangeHeader"
            mCtrl.TreeViewUpdate = True
            objTV.RefreshNode = NodeId.Parent
            RangeHeader NodeId
            mCtrl.TreeViewUpdate = False
            
         Case "RangeItem"
            objTV.RefreshNode = NodeId.Parent.Parent
            RangeHeader NodeId.Parent
            
         Case "LocalTests"
            objTV.RefreshNode = NodeId
            LocalTests NodeId
            fView.FrameToShow = Fra_HELP
            
         Case "Find"
            If fStatus = "Filter..." Then
               frmResultMapFilter.Show 1
               fStatus = frmResultMapFilter.FilterStatus
               Unload frmResultMapFilter
            Else
               fvSQL = ""
               ltSQL = ""
               fStatus = "Filter..."
            End If
            
            FirstView
            
         Case Else
            fView.FrameToShow = Fra_HELP
            blnShowBrowser = True
            
      End Select
'      objTV.ActiveNode.EnsureVisible
   End If
   
   'fView.ShowReadCodes
   tv1.Visible = True
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadResultMapping.RunWhat"
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
   Dim sampTxt As String
   Dim RS As New ADODB.Recordset
   
   Set pl = frmMain.edipr
   Select Case PageId
      Case "Map"
         If pl("LC").value <> "" Then
            With iceCmd
               .ActiveConnection = iceCon
               .CommandText = "ICECONFIG_Amend_InvTest_Codes"
               .CommandType = adCmdStoredProc
               .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
               .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , (objTV.newNode))
               .Parameters.Append .CreateParameter("Index", adInteger, adParamInputOutput, , IIf(plist("INDEX").value = "", 0, plist("INDEX").value))
               .Parameters.Append .CreateParameter("Org", adVarChar, adParamInput, 6, frmMain.cboTrust.Text)
               .Parameters.Append .CreateParameter("LTS", adInteger, adParamInput, , frmMain.CurrentLTSIndex)
               .Parameters.Append .CreateParameter("LocalCode", adVarChar, adParamInput, 10, pl("LC").value)
               .Parameters.Append .CreateParameter("Rubric", adVarChar, adParamInput, 35, pl("LD").value)
               .Parameters.Append .CreateParameter("ReadCode", adVarChar, adParamInput, 6, pl("RC").value)
               .Parameters.Append .CreateParameter("SpecCode", adVarChar, adParamInput, 15, pl("ST").value)
               .Parameters.Append .CreateParameter("OpUOM", adVarChar, adParamInput, 15, pl("UM").value)
               .Parameters.Append .CreateParameter("OPSuppress", adBoolean, adParamInput, , pl("SP").value)
               .Parameters.Append .CreateParameter("Active", adBoolean, adParamInput, , pl("AC").value)
               
               If blnShowReview Then
                  .Parameters.Append .CreateParameter("RevLev", adInteger, adParamInput, , pl("REV").value)
               Else
                  .Parameters.Append .CreateParameter("RevLev", adInteger, adParamInput, , 0)
               End If
               
               strSQL = "SET ANSI_NULLS ON; " & _
                        "SET QUOTED_IDENTIFIER ON; " & _
                        "SET ARITHABORT ON; " & _
                        "SET NUMERIC_ROUNDABORT OFF"
               iceCon.Execute strSQL, , adCmdText
               

               .Execute
               intRet = .Parameters("Return")
               If intRet = 0 Then
                  pl("INDEX").value = .Parameters("Index").value
               Else
                  eClass.FurtherInfo = "SQL Error " & intRet & " when adding to/Amending table " & .Parameters("Failure")
               End If
   '            newValue = pl("LC").value & "|" & pl("ST").value
            End With
            
'            RS.Close
            Set RS = Nothing
            
            strSQL = "SELECT Sample_Text " & _
                     "FROM CRIR_Sample_Type " & _
                     "WHERE Sample_Code = '" & plist("ST").value & "'"
            RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
            
            objTV.RefreshNode.Text = pl("LD").value & " - " & Replace(pl("LC").value, Chr(0), "") & " (" & RS!Sample_Text & ")"
            objTV.RefreshNode.Key = mCtrl.NewNodeKey(pl("INDEX").value, _
                                                     pl("LC").value, _
                                                     "Detail", _
                                                     , _
                                                     ms_DELETE)
            RS.Close
         End If
         Set RS = Nothing
'         objTView.ActiveNode = objTView.TopLevelNode(objTView.ActiveNode)
         
      Case "Range"
         eClass.FurtherInfo = "SQL Error when Adding to/Amending table EDI_InvTest_Ranges"
'         strARray = Split(pl.Pages("Range").Caption, "|")
         
         With iceCmd
            .ActiveConnection = iceCon
            .CommandText = "ICECONFIG_Amend_InvTest_Ranges"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
            .Parameters.Append .CreateParameter("AddNew", adBoolean, adParamInput, , objTV.newNode)
            .Parameters.Append .CreateParameter("Index", adVarChar, adParamInput, 6, pl("INDEX").value)
            .Parameters.Append .CreateParameter("Sex", adVarChar, adParamInput, 1, pl("IN1").value)
            .Parameters.Append .CreateParameter("MinAge", adBigInt, adParamInput, , pl("MIN_AGE").value)
            .Parameters.Append .CreateParameter("MaxAge", adBigInt, adParamInput, , pl("MAX_AGE").value)
            .Parameters.Append .CreateParameter("RangeLo", adVarChar, adParamInput, 10, pl("NUMMIN").value)
            .Parameters.Append .CreateParameter("RangeHi", adVarChar, adParamInput, 10, pl("NUMMAX").value)
            .Parameters.Append .CreateParameter("UOM", adVarChar, adParamInput, 15, pl("IN6").value)
            .Parameters.Append .CreateParameter("Comment", adVarChar, adParamInput, 70, pl("IN7").value)
            .Execute
            intRet = .Parameters("Return")
            failInfo = "ICECONFIG_Amend_InvTest_Ranges"
         End With
         newValue = pl.Pages("Range").Caption
         
   End Select
   Update = newValue
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadResultMapping.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function
