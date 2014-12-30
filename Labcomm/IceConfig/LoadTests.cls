VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private tvNode(3) As MSComctlLib.Node
Private nd(3) As Node
Private strSQL As String
Private plist As PropertiesList
Private mCtrl As New ManageControls
'Private rCtrl As LoadRules
Private curTest As String
Private newNode As Node
Private provCount As Integer
Private vData As Variant
Private OLESourceNode As Node
Private OLERelative As Long
Private OLEMoveKey As String
Private OLEMoveText As String

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
'   objctrl.TreeviewOrigin = "T"
'   objctrl.PracticeId = "System"
'   objctrl.SetDbStructures "Request_Tests"
   objTV.Origin = "T_DRAGDROP"
   
   With plist
      .Ownerdraw = plodSimple
      
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Control"
      .PropertyItems.Add "INDEX", "dbTestIndex", plpsString, , "Hidden"
      
      .Pages.Add "Main", "Request_Tests"
      
      .PropertyItems.Add "TESTINDEX", "Test Index", plpsNumber, , "Hidden"
      .PropertyItems.Add "TESTID", "Test Code", plpsString, , "Mandatory - Hidden"
      .PropertyItems.Add "TEST_CODE", "Test Code", plpsString, "", "Mandatory - Test code as required by laboratory system", , , pldsNormal
      .PropertyItems.Add "TESTSPEC", "Discipline for this test", plpsList, , "Dependent on test provider specialty setting"
      .PropertyItems.Add "SCREEN_PANEL", "Screen Panel", plpsList, "", "Mandatory - Panel with which test is associated"
      .PropertyItems.Add "PANEL_PAGE", "Panel Page", plpsList, "", "Page on panel on which test appears"
      .PropertyItems.Add "SCREEN_POSN", "Screen Position", plpsList, 1, "Position in which test is displayed on screen. ", , , pldsNormal
      .PropertyItems.Add "SCREEN_CAPTION", "Test Caption", plpsString, , "Mandatory - Caption displayed on screen for user selection"
      .PropertyItems.Add "SCREEN_HELP", "Label Help", plpsString, , "Help message displayed on screen upon user selection of test"
      .PropertyItems.Add "SCREEN_COLOUR", "Caption Colour", plpsColor, , "Mandatory - Colour of test text on screen"
      .PropertyItems.Add "HELP_COLOUR", "Help Colour", plpsColor, , "Colour of background behind help message"
      .PropertyItems.Add "INFO", "Test Information", plpsString, , "Information about the test"
      .PropertyItems.Add "PROV_ID", "Test Provider ID", plpsCustom, , "Mandatory - Which lab handles this test"
      .PropertyItems.Add "READ_CODE", "Read Code", plpsCustom, , "READ Code assigned to this test"
      .PropertyItems.Add "ENABLED", "Enabled", plpsBoolean, True, "Can users request this test"
      .PropertyItems.Add "TUBE_CODE", "Sample Container", plpsColor, , "Sample container required for this test"
      .PropertyItems.Add "VOL", "Sample Volume Required", plpsNumber, 0, "Amount of sample required to perform test"
      .PropertyItems.Add "PAED_TUBE_CODE", "Paediatric Container", plpsColor, , "Sample container required for this test when carried out on paediatric patient"
      .PropertyItems.Add "PAED_VOL", "Paediatric Volume Required", plpsNumber, 0, "Amount of sample required to perform test on a paediatric patient"
      .PropertyItems.Add "WORKLIST", "Worklist Enabled", plpsBoolean, False, "Worklist enabled"
      .PropertyItems.Add "HIST_ST", "History Search Type", plpsList, , "Select type of search you wish to do to show previous results prior to requesting this test"
      .PropertyItems.Add "HIST_SS", "History Search String", plpsString, , "Enter string you wish to search for (only valid if above option not set to None)"
      .PropertyItems.Add "HIST_MSG", "History Message", plpsString, , "Enter message to display to users when performing search"
      .PropertyItems.Add "EIM", "Exclude if Male", plpsBoolean, , "Do not allow this test to be selected if the patient is male"
      .PropertyItems.Add "EIF", "Exclude if female", plpsBoolean, , "Do not allow this test to be selected if the patient is female"
      .PropertyItems.Add "MIN_AGE", "Minimum age", plpsNumber, , "The minimum age to which this test is applicable"
      .PropertyItems.Add "MAX_AGE", "Maximum age", plpsNumber, , "The maximum age to which this test is applicable"
      .PropertyItems.Add "VALID", "Report Validity Period", plpsNumber, , "The length of time the test results remain valid"
      .PropertyItems.Add "TEST_TYPE", "Type of test", plpsCustom, , "Double-Click to see and bloodbank details"
   End With
   
   With plist
      .PropertyItems("INDEX").PageKeys = "Control"
      With .PropertyItems("TESTID")
         .Tag = "Test_Code"
         .Visible = False
      End With
      With .PropertyItems("TEST_CODE")
'         .Tag = "Test_Code"
         .DialogTitle = "*"
      End With
      With .PropertyItems("SCREEN_PANEL")
         .Tag = "Screen_Panel"
         .PageKeys = "Main"
         strSQL = "SELECT PanelID, PanelName " & _
                  "FROM Request_Panels " & _
                  "WHERE PanelType = 1 " & _
                  "ORDER BY PanelID"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         If RS.RecordCount > 0 Then
            Do While Not RS.EOF
               .ListItems.Add RS!panelName, Format(RS!PanelId)
               RS.MoveNext
            Loop
         End If
         RS.Close
      End With
      With .PropertyItems("TESTINDEX")
         .Tag = "Test_Index"
         .Visible = False
      End With
      With .PropertyItems("PANEL_PAGE")
         .DialogTitle = "*"
         .Tag = "Screen_Panel_Page"
      End With
      With .PropertyItems("SCREEN_POSN")
         .DialogTitle = "*"
         .Tag = "Screen_Position"
      End With
      With .PropertyItems("SCREEN_CAPTION")
         .Tag = "Screen_Caption"
         .max = 35
      End With
      With .PropertyItems("SCREEN_HELP")
         .Tag = "Screen_Help"
      End With
      With .PropertyItems("SCREEN_COLOUR")
         .Tag = "Screen_Colour"
         .defaultValue = 1
      End With
      With .PropertyItems("HELP_COLOUR")
         .Tag = "Screen_Help_Backcolour"
      End With
      With .PropertyItems("INFO")
         .Tag = "Information"
      End With
      With .PropertyItems("PROV_ID")
         .Tag = "Provider_Id"
      End With
      With .PropertyItems("READ_CODE")
         .Tag = "Read_Code"
      End With
      With .PropertyItems("TUBE_CODE")
         .Tag = "Tube_Code"
         .defaultValue = 1
      End With
      With .PropertyItems("PAED_TUBE_CODE")
         .Tag = "PaedTube_Code"
         .defaultValue = 1
      End With
      With .PropertyItems("ENABLED")
         .Tag = "Enabled"
      End With
      With .PropertyItems("VOL")
         .Tag = "Test_Volume"
         .max = 5000
         .Min = 0
         .Increment = 1
      End With
      With .PropertyItems("PAED_VOL")
         .Tag = "PaedTube_Test_Volume"
         .max = 5000
         .Min = 0
         .Increment = 1
      End With
      With .PropertyItems("WORKLIST")
         .Tag = "Worklist_Enabled"
      End With
      With .PropertyItems("HIST_ST")
         .Tag = "ResHistory_SearchType"
         .ListItems.Add "None", "NO"
         .ListItems.Add "Results", "SR"
         .ListItems.Add "Investigations", "SI"
      End With
      With .PropertyItems("HIST_SS")
         .Tag = "ResHistory_SearchString"
      End With
      With .PropertyItems("HIST_MSG")
         .Tag = "ResHistory_Message"
      End With
      With .PropertyItems("EIM")
         .Tag = "Exclude_If_Male"
      End With
      With .PropertyItems("EIF")
         .Tag = "Exclude_If_Female"
      End With
      With .PropertyItems("MIN_AGE")
         .Tag = "Min_Age"
         .Min = 0
         .max = 100
      End With
      With .PropertyItems("MAX_AGE")
         .Tag = "Max_Age"
         .Min = 0
         .max = 150
      End With
      With .PropertyItems("VALID")
         .Tag = "Report_Validity_Period"
      End With
      With .PropertyItems("TEST_TYPE")
         .DialogTitle = "*"
         .Tag = "Test_Type"
'         With .ListItems
'            .Add "Standard", 0
'            .Add "Blood Bank", 1
'            .Add "Histology", 2
'         End With
      End With
      With .PropertyItems("TESTSPEC")
         .Tag = "Discipline_Index"
         With .ListItems
            strSQL = "SELECT * " & _
                     "FROM Service_Discipline_Map " & _
                     "WHERE Specialty_Code <> '' " & _
                     "ORDER BY Specialty_Code"
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            Do Until RS.EOF
               .Add RS!Discipline_Expansion, RS!Discipline_Index
               RS.MoveNext
            Loop
            RS.Close
         End With
      End With
   End With
   
   Set RS = Nothing
   mCtrl.SetDbStructures plist
   plist.ActivePage = "Main"
   fView.FrameToShow = Fra_EDI
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.Class_Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get ControlObject() As ManageControls
   Set ControlObject = mCtrl
End Property

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim vData As Variant
   Dim tNode As Node
   
   If NodeId = objTV.TopLevelNode(NodeId) Then
      strSQL = "SELECT Count(Test_Index) " & _
               "FROM Request_Tests " & _
               "WHERE Provider_Id = " & objTV.NodeLevel(NodeId)
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      If RS(0) = 0 Then
         strSQL = "DELETE FROM Service_Providers " & _
                  "WHERE Provider_Id = " & objTV.NodeLevel(NodeId)
      Else
         strSQL = ""
         MsgBox "Delete or remove all the tests under " & NodeId.Text & " before deleting this provider"
      End If
      RS.Close
      Set RS = Nothing
      
   Else
   
      Set tNode = NodeId.Parent.Parent
      vData = objTV.ReadNodeData(NodeId)
      If vData(1) = "TestNode" Then
         strSQL = "DELETE FROM Request_Tests " & _
                  "WHERE Test_Index = " & vData(0) & "; " & _
                  "DELETE FROM Request_Test_Prompts " & _
                  "WHERE Test_Index = " & vData(0) & "; " & _
                  "DELETE FROM Request_Included_Tests " & _
                  "WHERE Included_Test_index = " & vData(0) & _
                  " OR Test_Index = " & vData(0) & "; " & _
                  "DELETE FROM Request_Excluded_Tests " & _
                  "WHERE Excluded_Test_Index = " & vData(0) & _
                  " OR Test_Index = " & vData(0) & "; " & _
                  "DELETE FROM Request_Reflex_Tests " & _
                  "WHERE Reflex_Test_Index = " & vData(0) & _
                  " OR Test_Index = " & vData(0) & "; " & _
                  "DELETE FROM Request_Bloodbank_Procedures " & _
                  "WHERE Test_Index = " & vData(0) & "; " & _
                  "DELETE FROM Request_Histology_Input_Type " & _
                  "WHERE Test_Index = " & vData(0)
      Else
         Select Case vData(2)
            Case "INC"
               strSQL = "DELETE FROM Request_Included_Tests " & _
                        "WHERE Test_Index = " & vData(1) & _
                           "AND Included_Test_index = " & vData(0)
            
            Case "EXC"
               strSQL = "DELETE FROM Request_Excluded_Tests " & _
                        "WHERE Test_Index = " & vData(1) & _
                           "AND Excluded_Test_Index = " & vData(0)
                    
            Case "RFX"
               strSQL = "DELETE FROM Request_Reflex_Tests " & _
                        "WHERE Test_Index = " & vData(1) & _
                           "AND Reflex_Test_Index = " & vData(0)
                           
            Case "RuleDetails"
               strSQL = "DELETE FROM Request_Test_Prompts " & _
                        "WHERE Test_Index = " & vData(1) & _
                           " AND Prompt_Index = " & vData(0)
                           
         End Select
      End If
   End If
   
   If strSQL <> "" Then
      iceCon.Execute strSQL
      tv1.SelectedItem = NodeId.Parent
      frmMain.TreeView1.Nodes.Remove NodeId.Index
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.FurtherInfo = strSQL
   eClass.CurrentProcedure = "LoadTests.Delete"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub DeleteBloodBank(testId As String)
   On Error GoTo procEH
   strSQL = "DELETE FROM Request_BloodBank_Products " & _
            "WHERE Test_index = " & testId & "; " & _
            "DELETE FROM Request_BloodBank_Procedures " & _
            "WHERE Test_index = " & testId
   iceCon.Execute strSQL
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.FurtherInfo = strSQL
   eClass.CurrentProcedure = "LoadTests.DeleteBloodBank"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
Public Sub Excluded(Index As String, pNode As MSComctlLib.Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   If pNode.Children > 0 Then
      If pNode.child.Text = "Please wait..." Then
         tv1.Nodes.Remove pNode.child.Index
         strSQL = "SELECT Screen_Caption, Request_Excluded_Tests.Test_Index, Excluded_Test_Index " & _
                  "FROM Request_Tests " & _
                     "INNER JOIN Request_Excluded_Tests " & _
                     "ON Request_Tests.Test_Index = Request_Excluded_Tests.Excluded_Test_Index " & _
                  "WHERE Request_Excluded_Tests.Test_Index = " & Index
         
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         If RS.RecordCount = 0 Then
         Else
            Do While Not RS.EOF
               tv1.Nodes.Add pNode, _
                             tvwChild, _
                             mCtrl.NewNodeKey(RS!Excluded_Test_Index, _
                                                 RS!Test_Index, _
                                                 "EXC", _
                                                 , ms_DELETE), _
                             RS!Screen_Caption, _
                             1, _
                             1
               RS.MoveNext
            Loop
         End If
         RS.Close
      End If
   End If
   fView.Show Fra_INCEX, "1"
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.Included"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function GetVacantScreenPositions(PanelId As String, _
                                         PageId As String, _
                                         curPosn As String) As Boolean
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim seq As Integer
   Dim i As Integer
   Dim sp As Integer
   
   eClass.FurtherInfo = "PanelId: " & PanelId & " - Page Id: " & PageId
   strSQL = "SELECT * " & _
            "FROM Request_Panels_Pages " & _
            "WHERE PanelId = " & PanelId & _
               " AND PageName = '" & PageId & "'"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   With frmMain
      If RS.RecordCount > 0 Then
         seq = Val(RS!Sequence)
         RS.Close
         .edipr.Redraw = False
         .edipr("SCREEN_POSN").ListItems.Clear
         strSQL = "SELECT Screen_Position " & _
                  "FROM Request_Tests " & _
                  "WHERE Screen_Panel = " & PanelId & _
                     " AND Screen_Panel_Page = '" & PageId & "' " & _
                  "ORDER BY Screen_Position"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         .ProgressBar1.value = 0
         .ProgressBar1.max = 45
         With .edipr("SCREEN_POSN")
   '         .ListItems.Add .value, .value
            For i = 1 To 45
               If RS.EOF Then
                  .ListItems.Add "Available", Format(i)
               Else
                  sp = Val(RS!Screen_Position) Mod 45
                  If sp = 0 Then
                     sp = 45
                  End If
                  If i < sp Then
                     .ListItems.Add "Available", Format(i)
'                     RS.MoveNext
                  ElseIf i = sp Then
                     If i = Val(curPosn) Then
                        .ListItems.Add "Current", Format(i)
                     End If
                     RS.MoveNext
                  End If
               End If
            Next i
            If .ListItems.Count > 0 Then
               .value = .ListItems(1).value
               GetVacantScreenPositions = True
            Else
               GetVacantScreenPositions = False
            End If
         End With
         RS.Close
         .ProgressBar1.value = 0
         .edipr.Redraw = True
      Else
         strSQL = "INSERT INTO Request_Panels_Pages (" & _
                     "PanelId, PageName, Sequence) " & _
                  "VALUES (" & _
                     PanelId & ", '" & PageId & "', 1)"
         iceCon.Execute strSQL
'         MsgBox "No Page set up for " & pageId & " - one has been created", vbExclamation, "Ice Configuration"
      End If
   End With
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.GetVacantScreenPositions"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Sub Included(Index As String, pNode As MSComctlLib.Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   If pNode.Children > 0 Then
      If pNode.child.Text = "Please wait..." Then
         tv1.Nodes.Remove pNode.child.Index
         strSQL = "SELECT Screen_Caption, Request_Included_Tests.Test_Index, Included_Test_Index " & _
                  "FROM Request_Tests " & _
                     "INNER JOIN Request_Included_Tests " & _
                     "ON Request_Tests.Test_Index = Request_Included_Tests.Included_Test_Index " & _
                  "WHERE Request_Included_Tests.Test_Index = " & Index
         
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         If RS.RecordCount = 0 Then
         Else
            Do While Not RS.EOF
               tv1.Nodes.Add pNode, _
                             tvwChild, _
                             mCtrl.NewNodeKey(RS!Included_Test_Index, _
                                                 RS!Test_Index, _
                                                 "INC", _
                                                 , ms_DELETE), _
                             RS!Screen_Caption, _
                             1, _
                             1
               RS.MoveNext
            Loop
         End If
         RS.Close
      End If
   End If
   fView.Show Fra_INCEX, "0"
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.Included"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Reflex(Index As String, pNode As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   If pNode.Children > 0 Then
      If pNode.child.Text = "Please wait..." Then
         tv1.Nodes.Remove pNode.child.Index
         strSQL = "SELECT Screen_Caption, Request_Reflex_Tests.Test_Index, Reflex_Test_Index " & _
                  "FROM Request_Tests " & _
                     "INNER JOIN Request_Reflex_Tests " & _
                     "ON Request_Tests.Test_Index = Request_Reflex_Tests.Reflex_Test_Index " & _
                  "WHERE Request_Reflex_Tests.Test_Index = " & Index
         
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         If RS.RecordCount = 0 Then
         Else
            Do While Not RS.EOF
               tv1.Nodes.Add pNode, _
                             tvwChild, _
                             mCtrl.NewNodeKey(RS!Reflex_Test_Index, _
                                                 RS!Test_Index, _
                                                 "RFX", _
                                                 , _
                                                 ms_DELETE), _
                             RS!Screen_Caption, _
                             1, _
                             1
               RS.MoveNext
            Loop
         End If
         RS.Close
      End If
   End If
   fView.Show Fra_INCEX, "2"
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.Reflex"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Rules(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim pNode As Node
   Dim tNode As Node
   frmMain.AutoRedraw = False
   If NodeId.Children > 0 Then
      Set pNode = NodeId.child
      Do Until TypeName(pNode.Next) = "Nothing"
         Set tNode = pNode.Next
         tv1.Nodes.Remove pNode.Index
         Set pNode = tNode
      Loop
      tv1.Nodes.Remove pNode.Index
   End If
   
   strSQL = "SELECT Request_Tests.Test_Index, Request_Test_Prompts.Sequence AS Seq, Request_Prompt.* " & _
            "FROM Request_Test_Prompts " & _
               "INNER JOIN Request_Tests ON " & _
               "Request_Test_Prompts.Test_Index = Request_Tests.Test_Index " & _
                  "INNER JOIN Request_Prompt ON " & _
                  "Request_Test_Prompts.Prompt_Index = Request_Prompt.Prompt_Index " & _
            "WHERE Request_Tests.Test_Index = " & objTV.NodeLevel(NodeId) & _
            " ORDER BY Sequence"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set tvNode(0) = tv1.Nodes.Add(NodeId, _
                                    tvwChild, _
                                    mCtrl.NewNodeKey(RS!Prompt_Index, _
                                                     RS!Test_Index, _
                                                     "RuleDetails", _
                                                     "DRAG|RULES|CHILD", _
                                                     ms_DELETE), _
                                    RS!Prompt_Desc & " (" & RS!Prompt_Type & ")", _
                                    1, _
                                    1)
      
'     Indicate node is drag-drop enabled
      tvNode(0).Tag = "*"
      tvNode(0).Checked = True
      RS.MoveNext
   Loop
   RS.Close
   fView.Show Fra_INCEX, "3"
   Set RS = Nothing
   frmMain.AutoRedraw = True
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadRules.RuleList"
   eClass.Add Err.Number, Err.Description, Err.Source

End Sub

Public Sub Providers()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim lngMenu As ENUM_MenuStatus
   
'   Set nd(0) = tv1.Nodes.Add(, _
'                             , _
'                             mCtrl.NewNodeKey("New", _
'                                              "PRV", _
'                                              "NewMapping", _
'                                              , _
'                                              , _
'                                              True), _
'                             "Add new provider", _
'                             1, _
'                             1)
'   nd(0).Bold = True
   
   strSQL = "SELECT DISTINCT Request_Tests.Provider_ID, Service_Providers.Provider_Name, Service_Providers.Discipline_Index " & _
            "FROM Request_Tests " & _
               "INNER JOIN Service_Providers ON " & _
               "Request_Tests.Provider_ID = Service_Providers.Provider_ID"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   provCount = RS.RecordCount
   RS.Close
   
'   If RS.EOF Then
'      RS.Close
      strSQL = "SELECT * " & _
               "FROM Service_Providers"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
'      Set nd(0) = tv1.Nodes.Add(, _
'                                , _
'                                mCtrl.NewNodeKey(RS!Provider_ID, _
'                                                 "PRV", _
'                                                 "LoadTests", _
'                                                 Trim(RS!Discipline_Index & ""), _
'                                                 ms_BOTH), _
'                                RS!Provider_Name, _
'                                1, _
'                                1)
'      tv1.Nodes.Add nd(0), _
'                    tvwChild, _
'                    mCtrl.NewNodeKey("Temp", "Temp"), _
'                    "Please wait...", _
'                    1, _
'                    1
'   Else
      If RS.RecordCount > 1 Then
         lngMenu = ms_BOTH
      Else
         lngMenu = ms_ADD
      End If
      
      Do Until RS.EOF
         Set nd(0) = tv1.Nodes.Add(, _
                                   , _
                                   mCtrl.NewNodeKey(RS!Provider_ID, _
                                                    "PRV", _
                                                    "LoadTests", _
                                                    Trim(RS!Discipline_Index & ""), _
                                                    lngMenu), _
                                   RS!Provider_Name, _
                                   1, _
                                   1)
         tv1.Nodes.Add nd(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey("Temp", "Temp"), _
                       "Please wait...", _
                       1, _
                       1
   
         RS.MoveNext
      Loop
'   End If
   fView.Show Fra_HELP
   RS.Close
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.Providers"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get ProviderTotal() As Integer
   ProviderTotal = provCount
End Property

Public Sub LoadTests(NodeId As Node)
   On Error GoTo procEH
'   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim TestName As String
   Dim provID As String
   Dim tNode As Node
   Dim i As Integer
   
   fView.FrameToShow = Fra_HELP
   tv1.Visible = False
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   Load frmWait
   frmWait.Label1.Caption = "Please wait whilst tests details for this organisation are loaded..."
   frmWait.Show
   frmWait.Refresh

   provID = objTV.NodeLevel(NodeId)
   strSQL = "SELECT * " & _
            "FROM Request_Tests " & _
            "WHERE Test_Code LIKE '" & Format(frmMain.cboTrust.Text, "!@@@@@@") & "%' " & _
               "AND Provider_ID = " & provID & _
            " ORDER BY Screen_Caption"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   If RS.RecordCount > 0 Then
      frmMain.ProgressBar1.max = RS.RecordCount
      frmMain.ProgressBar1.value = 0
      frmMain.ProgressBar1.Visible = True
      Do While Not RS.EOF
         If Trim(RS!Screen_Caption & "") = "" Then
             TestName = Right$(RS!Test_Code, Len(RS!Test_Code) - 6)
         Else
             TestName = Trim(RS!Screen_Caption)
         End If
         
         Set tNode = tv1.Nodes.Add(NodeId, _
                                   tvwChild, _
                                   mCtrl.NewNodeKey(RS!Test_Index, _
                                                    "TestNode", _
                                                    "TestDetails", _
                                                    , _
                                                    ms_DELETE), _
                                   TestName, _
                                   1, _
                                   1)
         
         
         SetUpTestNode tNode
'         Set tNode = SetUpTestNode(NodeId, RS!Test_Index, RS!Screen_Caption)
         If Trim(RS!Screen_Panel_Page & "") = "" Then
            tNode.ForeColor = BPBLUE
         Else
            tNode.ForeColor = BPGREEN
         End If
         
         If RS!Enabled Then
'            tNode.Image = 1
'            tNode.SelectedImage = 1
'            tNode.ForeColor = BPBLUE
         Else
'            tNode.Image = 31
'            tNode.SelectedImage = 31
            tNode.ForeColor = BPRED
         End If
         RS.MoveNext
      Loop
   End If
   tv1.Visible = True
   Unload frmWait
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.LoadTests"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub MenuAddEntry()
   Dim testId As String
   Dim thisNode As Node
   
'   If objTV.ActiveNode.Children > 0 Then
'      testId = objTV.NodeLevel(objTV.TopLevelNode(objTV.ActiveNode, True))
'   End If
   If objTV.TopLevelNode <> objTV.ActiveNode Then
      testId = objTV.NodeLevel(objTV.TopLevelNode(objTV.ActiveNode, True))
   End If
   Set thisNode = objTV.ActiveNode ' tv1.SelectedItem
   If thisNode.Text = "Rules" Then
'      Set rCtrl = New LoadRules
      frmBehaviour.ControlObject = mCtrl
      frmBehaviour.Show 1
      WriteTestRules thisNode
      TestDetails thisNode, True
   ElseIf objTV.nodeKey(objTV.ActiveNode) = "PRV" Then
      NewTest
   Else
      frmTestMnt.NodeType = objTV.nodeKey(thisNode)
      frmTestMnt.CurrentTest = Val(testId)
'      frmTestMnt.CurrentTestCode = plist("TEST_CODE").value
      frmTestMnt.Show 1
      RefreshChildNodes thisNode  '  objtv.TopLevelNode(tv1.SelectedItem, True)
   End If
End Sub

Public Sub NewMapping()
   frmProviderEditor.Show 1
End Sub

Public Sub NewTest()
   Dim i As Integer
   Dim provID As String
   Dim tSpec As String
   Dim RS As New ADODB.Recordset
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode)
   provID = vData(0) ' objTV.NodeLevel(objTV.TopLevelNode(objTV.ActiveNode))
   tSpec = vData(4)
   
   For i = 1 To frmMain.edipr.PropertyItems.Count
      frmMain.edipr(i).value = frmMain.edipr(i).defaultValue
   Next i
   
   If newNode Is Nothing Then
      With plist("TESTSPEC")
'         If Val(tSpec) = 0 Then
'            .value = IIf(Val(Trim(RS!Discipline_Index & "")) = 0, "", Val(Trim(RS!Discipline_Index & "")))
'         Else
            .value = Val(tSpec)
'         End If
         
         If Val(tSpec) = 0 Then
            .Description = "Mandatory - No specialty specified for test provider"
            .ReadOnly = False
            .Style = plpsList
'            With .ListItems
'               strSQL = "SELECT * " & _
'                        "FROM Service_Discipline_Map " & _
'                        "WHERE Specialty_Code <> '' " & _
'                        "ORDER BY Specialty_Code"
'               RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'               Do Until RS.EOF
'                  .Add RS!Discipline_Expansion, RS!Discipline_Index
'                  RS.MoveNext
'               Loop
'               RS.Close
'            End With
         Else
            .Description = "Read only (Set by test provider specialty)"
            .ReadOnly = True
            .Style = plpsString
         End If
      End With
      
      strSQL = "SELECT max(Test_Index) " & _
               "FROM Request_Tests"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
      plist("INDEX").value = RS(0) + 1
      plist("TESTINDEX").value = RS(0) + 1
      
      RS.Close
      Set newNode = tv1.Nodes.Add(objTV.TopLevelNode, _
                                  tvwChild, _
                                  mCtrl.NewNodeKey(plist("INDEX").value, _
                                                   "TestNode", _
                                                   "TestDetails", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  1, _
                                  1)
      
      
      
'      Set newNode = SetUpTestNode(objTV.TopLevelNode(objTV.ActiveNode), "0", "New")
   End If
   objTV.ActiveNode = newNode
   
   With frmMain.edipr
      With .PropertyItems("TESTSPEC")
         If tSpec = "" Then
            .Description = "Mandatory - No specialty specified for test provider"
            .ReadOnly = False
            .Style = plpsList
         Else
            .Description = "Read only (Set by test provider specialty)"
            .ReadOnly = True
            .Style = plpsString
      With plist("TESTSPEC")
         If Val(tSpec) = 0 Then
            .Description = "Mandatory - No specialty specified for test provider"
            .ReadOnly = False
            .Style = plpsList
'            With .ListItems
'               strSQL = "SELECT * " & _
'                        "FROM Service_Discipline_Map " & _
'                        "WHERE Specialty_Code <> '' " & _
'                        "ORDER BY Specialty_Code"
'               RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'               Do Until RS.EOF
'                  .Add RS!Discipline_Expansion, RS!Discipline_Index
'                  RS.MoveNext
'               Loop
'               RS.Close
'            End With
         Else
            .Description = "Read only (Set by test provider specialty)"
            .ReadOnly = True
            .Style = plpsString
         End If
      End With
            .value = tSpec
         End If
      End With
      
      .PropertyItems("PROV_ID").DialogTitle = objTV.TopLevelNode(objTV.ActiveNode).Text
      .PropertyItems("PROV_ID").value = provID
      .PropertyItems("HELP_COLOUR").value = Val("&HFFFFFF&")
      .PropertyItems("HELP_COLOUR").DialogTitle = "0" & "-None Set"
      .PropertyItems("SCREEN_COLOUR").value = Val("&HFFFFFF&")
      .PropertyItems("SCREEN_COLOUR").DialogTitle = "0" + "-None Set"
'      .PropertyItems("PROV_ID").value = "0"
'      .PropertyItems("PROV_ID").DialogTitle = "None Set"
      .PropertyItems("TUBE_CODE").value = Val("&HFFFFFF&")
      .PropertyItems("TUBE_CODE").DialogTitle = "0" & "-No container required"
      .PropertyItems("PAED_TUBE_CODE").value = Val("&HFFFFFF&")
      .PropertyItems("PAED_TUBE_CODE").DialogTitle = "0" & "-No container required"
   End With
   fView.Show Fra_EDI, "Tests"
   Set RS = Nothing
End Sub

Public Function Refresh() As String
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim blnUpdate As Boolean
   
   blnUpdate = False
   For i = 0 To Forms.Count - 1
      If Forms(i) Is frmBloodBank Then
         blnUpdate = True
         Exit For
      End If
   Next i
   
   If objTV.newNode Then
'      strSQL = "SELECT max(Test_Index) " & _
'               "FROM Request_Tests"
'      RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'      Refresh = RS(0)
'      plist("INDEX").value = RS(0)
      objTV.ActiveNode.Text = frmMain.edipr("SCREEN_CAPTION").value
      objTV.ActiveNode.Key = mCtrl.NewNodeKey(plist("INDEX").value, _
                                              "TestNode", _
                                              "TestDetails", _
                                              , _
                                              ms_DELETE)
'      RS.Close
      SetUpTestNode objTV.ActiveNode
      Set newNode = Nothing
   End If
   
   If blnUpdate Then
      iceCon.Execute frmBloodBank.SQLData
      strSQL = "SELECT count(*) " & _
               "FROM Request_Bloodbank_Procedures " & _
               "WHERE Test_Index = " & Val(plist("INDEX").value)
      RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      If RS(0) = 0 Then
         RS.Close
         strSQL = "SELECT * " & _
                  "FROM Request_Histology_Input_Type " & _
                  "WHERE Test_Index = " & Val(plist("INDEX").value)
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         If RS.EOF Then
            plist("TEST_TYPE").value = 0
         End If
      End If
      Unload frmBloodBank
   End If
   
   Set RS = Nothing
   objTV.RefreshNode = objTV.ActiveNode
End Function

Public Sub RefreshChildNodes(FromNode As MSComctlLib.Node)
   Do Until FromNode.Children = 0
      tv1.Nodes.Remove FromNode.child.Index
   Loop
   tv1.Nodes.Add FromNode, tvwChild, "Temp", "Please wait...", 1, 1
   Select Case FromNode.Text
      Case "Excluded Tests"
         Excluded objTV.NodeLevel(FromNode), FromNode
      
      Case "Included Tests"
         Included objTV.NodeLevel(FromNode), FromNode
      
      Case "Reflex Tests"
         Reflex objTV.NodeLevel(FromNode), FromNode
         
      Case "Rules"
         Rules FromNode
         
   End Select
End Sub

Public Sub RunWhat(NodeId As MSComctlLib.Node, _
                   Optional Refresh As Boolean)
   On Error GoTo procEH
   Dim tNode As Node
   Dim vData As Variant
   
   If transCount > 0 Then
      iceCon.RollbackTrans
      transCount = 0
   End If
   
   Unload frmBloodBank
   If TidyUp = False Then
      vData = objTV.ReadNodeData(NodeId)
      fView.FrameToShow = Fra_EDI
      Select Case vData(2)
         Case "LoadTests"
            LoadTests NodeId
         
         Case "NewMapping"
            NewMapping
         
         Case "TestDetails"
            TestDetails NodeId, Refresh
         
'            If IsNumeric(Refresh) Then
'               TestDetails Refresh, "Yes"
'            Else
'               TestDetails CStr(vData(0)), Refresh
'            End If
            
         Case "NewTest"
            NewTest
         
         Case "Included"
            objTV.ActiveNode = NodeId.Parent
            TestDetails NodeId.Parent, Refresh
            objTV.ActiveNode = NodeId
            Included CStr(vData(0)), NodeId
            
         Case "Excluded"
            objTV.ActiveNode = NodeId.Parent
            TestDetails NodeId.Parent, Refresh
            objTV.ActiveNode = NodeId
            Excluded CStr(vData(0)), NodeId
            
         Case "Reflex"
            objTV.ActiveNode = NodeId.Parent
            TestDetails NodeId.Parent, Refresh
            objTV.ActiveNode = NodeId
            Reflex CStr(vData(0)), NodeId
            
         Case "Rules"
            objTV.ActiveNode = NodeId.Parent
            TestDetails NodeId.Parent, Refresh
            objTV.ActiveNode = NodeId
            Rules NodeId
            
         Case Else
            objTV.ActiveNode = NodeId.Parent.Parent
            TestDetails NodeId.Parent.Parent, Refresh
            objTV.ActiveNode = NodeId
            
      End Select
   End If
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "LoadTests.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function SetUpPages(PanelId As String)
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   strSQL = "SELECT PageName " & _
            "FROM Request_Panels_Pages " & _
            "WHERE PanelID = " & PanelId & _
            " ORDER BY PageName"
   
   With frmMain.edipr("PANEL_PAGE")
      .DialogTitle = "*"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount > 0 Then
         .ListItems.Clear
         .ListItems.Add 0, "<No Page>"
         i = 1
         Do While Not RS.EOF
            .ListItems.Add i, RS!PageName
            i = i + 1
            RS.MoveNext
         Loop
      Else
         .ListItems.Add "<No Page>", 0
         .value = ""
      End If
      RS.Close
   End With
   Set RS = Nothing
End Function

Private Function SetUpTestNode(NodeId As Node) ', _
                               TestIndex As String, _
                               TestName) As node
   
   Dim TestIndex As String
'   Set nd(0) = tv1.Nodes.Add(NodeId, _
'                             tvwChild, _
'                             mCtrl.NewNodeKey(TestIndex, _
'                                              "TestNode", _
'                                              "TestDetails", _
'                                              , _
'                                              ms_DELETE, _
'                                              IIf(TestName = "New", True, False)), _
'                             TestName, _
'                             1, _
'                             1)
   vData = objTV.ReadNodeData(NodeId)
   TestIndex = vData(0)
   Set nd(1) = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(TestIndex, _
                                              "INC", _
                                              "Included", _
                                              , _
                                              ms_ADD), _
                             "Included Tests", _
                             1, _
                             1)
   tv1.Nodes.Add nd(1), _
                 tvwChild, _
                 mCtrl.NewNodeKey("I", "Temp", "None"), _
                 "Please wait...", _
                 1, _
                 1
                                 
   Set nd(1) = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(TestIndex, _
                                              "EXC", _
                                              "Excluded", _
                                              , _
                                              ms_ADD), _
                             "Excluded Tests", _
                             1, _
                             1)
                                 
   tv1.Nodes.Add nd(1), _
                 tvwChild, _
                 mCtrl.NewNodeKey("I", "Temp", "None"), _
                 "Please wait...", _
                 1, _
                 1
   
   Set nd(1) = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(TestIndex, _
                                              "RFX", _
                                              "Reflex", _
                                              , _
                                              ms_ADD), _
                             "Reflex Tests", _
                             1, _
                             1)
   tv1.Nodes.Add nd(1), _
                 tvwChild, _
                 mCtrl.NewNodeKey("I", "Temp", "None"), _
                 "Please wait...", _
                 1, _
                 1
                                 
   Set nd(1) = tv1.Nodes.Add(NodeId, _
                             tvwChild, _
                             mCtrl.NewNodeKey(TestIndex, _
                                              "RLS", _
                                              "Rules", _
                                              "DROP|RULES|PARENT", _
                                              ms_ADD), _
                             "Rules", _
                             1, _
                             1)
   tv1.Nodes.Add nd(1), _
                 tvwChild, _
                 mCtrl.NewNodeKey("I", "Temp", "None"), _
                 "Please wait...", _
                 1, _
                 1
   Set SetUpTestNode = nd(1)
End Function

Public Function TV_DragDrop(NodeId As Node, _
                            Optional data As MSComctlLib.DataObject = Nothing) As Long
   Dim TVEffect As Long
   
   vData = objTV.ReadNodeData(NodeId)
   
   OLEMoveKey = objTV.ActiveNode.Key
   OLEMoveText = objTV.ActiveNode.Text
   
   TV_DragDrop = vbDropEffectNone
   
   If Left(vData(4), 4) = "DROP" Then
      If MsgBox("Are you sure you wish to move '" & objTV.ActiveNode.Text & _
                "' above '" & NodeId.child.FirstSibling.Text & "'?", _
                vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
         Set OLESourceNode = NodeId.child
         OLERelative = tvwFirst
         TVEffect = vbDropEffectMove
      End If
   Else
      
      If MsgBox("Are you sure you wish to move '" & objTV.ActiveNode.Text & _
                "' below '" & NodeId.Text & "'?", _
                vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
         Set OLESourceNode = NodeId
         OLERelative = tvwNext
         TVEffect = vbDropEffectMove
      End If
   End If
   
   If TVEffect = vbDropEffectMove Then
      tv1.Nodes.Remove objTV.ActiveNode.Index
      objTV.ActiveNode = Nothing
      tv1.Nodes.Add OLESourceNode, _
                    OLERelative, _
                    OLEMoveKey, _
                    OLEMoveText, _
                    1
   
      WriteTestRules OLESourceNode.Parent
   End If
   
   TV_DragDrop = vbDropEffectNone
End Function

Public Function TV_DragOver(NodeId As Node) As Long
   Dim dhData As Variant
   Dim lngEffect As Long
   
   dhData = objTV.ReadNodeData(NodeId)
   vData = objTV.ReadNodeData(objTV.ActiveNode)
   lngEffect = vbDropEffectNone
   If Left(vData(4), 4) = "DRAG" Then
      If Left(dhData(4), 4) = "DRAG" Then
         If dhData(1) = vData(1) Then
            lngEffect = vbDropEffectMove
         End If
      
      ElseIf Left(dhData(4), 4) = "DROP" Then
         If dhData(0) = vData(1) Then
            lngEffect = vbDropEffectMove
         End If
      End If
   End If
   TV_DragOver = lngEffect
End Function

Public Function TestDetails(NodeId As Node, _
                            Refresh As Boolean) 'TestIndex As String, _
                            Optional Refresh As String = "") As MSComctlLib.node
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim i As Integer
   Dim tCol As String
   Dim iceCmd As ADODB.Command
   Dim eList As Integer
   Dim tSpec As String
   
   vData = objTV.ReadNodeData(objTV.TopLevelNode(NodeId))
   tSpec = vData(4)
   vData = objTV.ReadNodeData(NodeId)
   plist("INDEX").value = vData(0)
   strSQL = "SELECT * " & _
            "FROM Request_Tests " & _
            "WHERE Test_Index = " & vData(0)
   eClass.FurtherInfo = strSQL
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      If Refresh = False Then
         objTV.ActiveNode.Text = RS!Screen_Caption
         objTV.ActiveNode.Key = mCtrl.NewNodeKey(CStr(vData(0)), "TestNode", "TestDetails", , ms_DELETE)
      End If
      With plist
         .Redraw = False
         .PropertyItems("TESTINDEX").value = RS!Test_Index
         .PropertyItems("TESTID").value = RS!Test_Code
         .PropertyItems("TEST_CODE").value = Mid(RS!Test_Code, 7)
         .PropertyItems("SCREEN_PANEL").value = Format(RS!Screen_Panel)
         
         SetUpPages RS!Screen_Panel
         
         .PropertyItems("PANEL_PAGE").value = Trim(RS!Screen_Panel_Page)
         
         If Trim(RS!Screen_Panel_Page & "") = "" Then
            .PropertyItems("SCREEN_POSN").ListItems.Clear
            .PropertyItems("SCREEN_POSN").value = ""
         Else
            GetVacantScreenPositions RS!Screen_Panel, _
                                     Trim(RS!Screen_Panel_Page), _
                                     RS!Screen_Position
         End If
         
         .PropertyItems("SCREEN_POSN").value = Format(RS!Screen_Position)
         
'         If Refresh Then
            If Trim(RS!Screen_Panel_Page & "") = "" Then
               NodeId.ForeColor = BPBLUE
            Else
               NodeId.ForeColor = BPGREEN
            End If
            If RS!Enabled = False Then
               NodeId.ForeColor = BPRED
            End If
'         End If
         
'            .propertyitems("DEPT").value = RS!Department
         .PropertyItems("SCREEN_CAPTION").value = RS!Screen_Caption
         .PropertyItems("SCREEN_HELP").value = RS!Screen_Help
         
         If Trim(RS!Screen_Colour & "") <> "" Then
            strSQL = "SELECT Colour_Code, Colour_Name " & _
                     "FROM Colours " & _
                     "WHERE Colour_Index = '" & RS!Screen_Colour & "'"
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            
            If RS2.RecordCount > 0 Then
               .PropertyItems("SCREEN_COLOUR").value = Val(RS2!Colour_Code)
               .PropertyItems("SCREEN_COLOUR").DialogTitle = Format(RS!Screen_Colour) & "-" & RS2!Colour_Name
            Else
               .PropertyItems("SCREEN_COLOUR").value = Val("&HFFFFFF&")
               .PropertyItems("SCREEN_COLOUR").DialogTitle = "0" + "-None Set"
            End If
            RS2.Close
         End If
         
         If Trim(RS!Screen_Help_Backcolour & "") <> "" Then
            strSQL = "SELECT Colour_Code, Colour_Name " & _
                     "FROM Colours " & _
                     "WHERE Colour_Index = '" & RS!Screen_Help_Backcolour & "'"
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            If RS2.RecordCount > 0 Then
               .PropertyItems("HELP_COLOUR").value = Val(RS2!Colour_Code)
               .PropertyItems("HELP_COLOUR").DialogTitle = Format(RS!Screen_Help_Backcolour) + "-" + RS2!Colour_Name
            Else
               .PropertyItems("HELP_COLOUR").value = Val("&HFFFFFF&")
               .PropertyItems("HELP_COLOUR").DialogTitle = "0-None Set"
            End If
            RS2.Close
         End If
         
         .PropertyItems("INFO").value = Trim(RS!Information)
         
         If Trim(RS!Provider_ID & "") <> "" Then
            strSQL = "SELECT Provider_Name " & _
                     "FROM Service_Providers " & _
                     "WHERE Provider_ID = " & RS!Provider_ID
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            If RS2.RecordCount > 0 Then
               .PropertyItems("PROV_ID").value = Val(RS!Provider_ID)
               .PropertyItems("PROV_ID").DialogTitle = RS2!Provider_Name & ""
            Else
               .PropertyItems("PROV_ID").value = "0"
               .PropertyItems("PROV_ID").DialogTitle = "None Set"
            End If
            RS2.Close
         End If
         
         .PropertyItems("READ_CODE").value = Trim(RS!Read_Code)
         
         If Trim(RS!Tube_Code & "") <> "" Then
            strSQL = "SELECT Name, Colour, Description " & _
                     "FROM Request_Tubes " & _
                     "WHERE Tube_Index = " & RS!Tube_Code
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            If RS2.RecordCount > 0 Then
               tCol = Trim(RS2!Colour & "")
               .PropertyItems("TUBE_CODE").DialogTitle = Format(RS!Tube_Code) + "-" + RS2!Name + " - " + Format(RS2!Description)
               RS2.Close
               If tCol <> "" Then
                  strSQL = "SELECT Colour_Code " & _
                           "FROM Colours " & _
                           "WHERE Colour_Index = '" & tCol & "'"
                  
                  RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
                  .PropertyItems("TUBE_CODE").value = Val(RS2!Colour_Code)
                  RS2.Close
               End If
            Else
               .PropertyItems("TUBE_CODE").value = Val("&HFFFFFF&")
               .PropertyItems("TUBE_CODE").DialogTitle = "0-No container required"
               RS2.Close
            End If
         End If
         
         If Trim(RS!PaedTube_Code & "") <> "" Then
            strSQL = "SELECT Name, Colour, Description " & _
                     "FROM Request_Tubes " & _
                     "WHERE Tube_Index = " & RS!PaedTube_Code
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            If RS2.RecordCount > 0 Then
               tCol = Trim(RS2!Colour & "")
               .PropertyItems("PAED_TUBE_CODE").DialogTitle = Format(RS!PaedTube_Code) + "-" + RS2!Name + " - " + Format(RS2!Description)
               RS2.Close
               If tCol <> "" Then
                   strSQL = "SELECT Colour_Code " & _
                            "FROM Colours " & _
                            "WHERE Colour_Index = '" & tCol & "'"
                   RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
                   .PropertyItems("PAED_TUBE_CODE").value = Val(RS2!Colour_Code)
                   RS2.Close
               End If
            Else
               .PropertyItems("PAED_TUBE_CODE").value = Val("&HFFFFFF&")
               .PropertyItems("PAED_TUBE_CODE").DialogTitle = "0-No container required"
               RS2.Close
            End If
         End If
         
         .PropertyItems("ENABLED").value = RS!Enabled
         .PropertyItems("VOL").value = IIf(IsNull(RS!Test_Volume), 0, RS!Test_Volume)
         .PropertyItems("PAED_VOL").value = IIf(IsNull(RS!PaedTube_Test_Volume), 0, RS!PaedTube_Test_Volume)
'            .propertyitems("DYNAMIC").value = RS!Dynamic
'            .propertyitems("SENSITIVE").value = RS!Sensitive
         .PropertyItems("WORKLIST").value = RS!Worklist_Enabled
         .PropertyItems("HIST_ST").value = Trim(RS!ResHistory_SearchType & "")
         .PropertyItems("HIST_SS").value = RS!ResHistory_SearchString
         .PropertyItems("HIST_MSG").value = RS!ResHistory_Message
         .PropertyItems("EIM").value = IIf(IsNull(RS!Exclude_If_Male), False, RS!Exclude_If_Male)
         .PropertyItems("EIF").value = IIf(IsNull(RS!Exclude_If_female), False, RS!Exclude_If_female)
         .PropertyItems("MIN_AGE").value = IIf(IsNull(RS!Min_Age), 0, RS!Min_Age)
         .PropertyItems("MAX_AGE").value = IIf(IsNull(RS!Max_Age), 0, RS!Max_Age)
         .PropertyItems("VALID").value = RS!Report_Validity_Period
         .PropertyItems("TEST_TYPE").value = Val(Trim(RS!Test_Type & ""))
         With .PropertyItems("TESTSPEC")
            If Val(tSpec) = 0 Then
'            If Val(Trim(RS!Discipline_Index & "")) = 0 Then
               .value = IIf(Val(Trim(RS!Discipline_Index & "")) = 0, "", Val(Trim(RS!Discipline_Index & "")))
            Else
               .value = Val(tSpec)
            End If
            
            If Val(tSpec) = 0 Then
               .Description = "Mandatory - No specialty specified for test provider"
               .ReadOnly = False
               .Style = plpsList
            Else
               .Description = "Read only (Set by test provider specialty)"
               .ReadOnly = True
               .Style = plpsString
            End If
         End With
         .Caption = "WHERE Test_Index = " & vData(0)
         .Redraw = True
      End With
   Else
   End If
   RS.Close
   Set RS = Nothing
   fView.Show Fra_EDI, "Tests"
   Exit Function
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.TestDetails"
   eClass.Add Err.Number, Err.Description, Err.Source
End Function

Public Function TidyUp() As Boolean
   TidyUp = False
   If Not newNode Is Nothing Then
'      If objTV.NodeLevel(newNode) = 0 Then
         tv1.Nodes.Remove newNode.Index
         tv1.SelectedItem = tv1.Nodes(1)
         TidyUp = True
'      End If
      Set newNode = Nothing
   End If
End Function

Public Sub WriteTestRules(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim TVN As MSComctlLib.Node
   Dim SeqNo As Integer
   Dim vData As Variant
   Dim i As Integer
   Dim exc1 As Integer
   Dim exc2 As Integer
   Dim blnReset As Boolean
   
   If NodeId.Children > 0 Then
      Set TVN = NodeId.child
      vData = objTV.ReadNodeData(TVN)
      eClass.FurtherInfo = strSQL
      If vData(1) <> "Temp" Then
         strSQL = "DELETE FROM Request_Test_Prompts " & _
                  "WHERE Test_Index = " & vData(1)
         iceCon.Execute strSQL
      End If
'      SeqNo = 1
'      strSQL = "SELECT Prompt_Index " & _
'               "FROM Request_Prompt " & _
'               "WHERE Prompt_Type = 'EIM' " & _
'                  "OR Prompt_Type = 'EIF'"
'      RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'      exc1 = RS!Prompt_Index
'      RS.MoveNext
'      exc2 = RS!Prompt_Index
'      RS.Close
      
      blnReset = False
      For i = 1 To NodeId.Children
         vData = objTV.ReadNodeData(TVN)
         If vData(1) <> "Temp" Then
'            If vData(0) = exc1 Or vData(0) = exc2 Then
'               strSQL = "INSERT INTO Request_Test_Prompts (" & _
'                           "Test_Index, Prompt_Index, Sequence) " & _
'                        "VALUES (" & _
'                           vData(1) & ", " & _
'                           vData(0) & ", " & _
'                           1 & ")"
'               If i <> 1 Then
'                  blnReset = True
'                  MsgBox "An exclude rule must be the first rule under a test," & _
'                         vbCrLf & "it's position has been reset to 1", vbInformation, "Rule Order Information"
'               End If
'            Else
               strSQL = "INSERT INTO Request_Test_Prompts (" & _
                           "Test_Index, Prompt_Index, Sequence) " & _
                        "VALUES (" & _
                           vData(1) & ", " & _
                           vData(0) & ", " & _
                           i + 1 & ")"
'            End If
            eClass.FurtherInfo = strSQL
            iceCon.Execute strSQL
         End If
         Set TVN = TVN.Next
'         SeqNo = SeqNo + 1
      Next i
      If blnReset Then
         Rules NodeId
      End If
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadTests.WriteTestRules"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub Class_Terminate()
   Unload frmBehaviour
   Unload frmBloodBank
End Sub


