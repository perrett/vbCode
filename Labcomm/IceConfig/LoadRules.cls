VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadRules"
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
   Dim RS As New ADODB.Recordset
   
   objTV.Origin = "RULES"
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr
   With plist
      .Redraw = False
      .Pages.Clear
      .PropertyItems.Clear
      .Pages.Add "Rules", "Request_Prompt"
      .PropertyItems.Add "STYLE", "Rule Style", plpsString, "", "Mandatory - The type of rule"
      .PropertyItems.Add "NAME", "Rule Name", plpsString, , "Mandatory - Name of this rule"
      .PropertyItems.Add "SAVETYPE", "Save Data As", plpsList, , "Mandatory - Select how the data is stored.  Clinical Detail is encoded in the PDF417 barcode, Test Information is not"
      .PropertyItems.Add "SAVEHEADER", "Save As Header", plpsString, , "Enter the tag you wish to use to identify this data"
      .PropertyItems.Add "QUESTION", "Question to Ask", plpsString, , "Mandatory - Enter the question you wish to ask the user on selection of the associated request"
      .PropertyItems.Add "YESTEXT", "Yes Value", plpsString, "Yes", "Enter the text you wish to save when the user clicks the Yes button"
      .PropertyItems.Add "NOTEXT", "No Value", plpsString, "No", "Enter the text you wish to save when the user clicks the No button"
      .PropertyItems.Add "YESACTION", "Yes Action", plpsCustom, , "Select the action you wish to take on clicking the Yes button"
      .PropertyItems.Add "NOACTION", "No Action", plpsCustom, , "Select the action you wish to take on clicking the No button"
      .PropertyItems.Add "CANCELTEXT", "Cancel Text", plpsString, , "If selected action cancels the test, message to display on the screen to inform user"
      .PropertyItems.Add "PROMPT", "User Prompt", plpsString, , "Mandatory - Enter the prompt you wish to display for data entry"
      .PropertyItems.Add "TYPE", "Type", plpsList, , "Mandatory - Select the type of data you wish the user to enter"
      .PropertyItems.Add "NUMMAX", "Numeric Upper Limit", plpsNumber, , "Enter the maximum numeric value a user can enter"
      .PropertyItems.Add "NUMMIN", "Numeric Lower Limit", plpsNumber, , "Enter the minimum numeric value a user can enter"
      .PropertyItems.Add "PICKLIST", "Picklist", plpsList, , "Select the picklist you wish the user to select from"
      .PropertyItems.Add "HELPTEXT", "Help Text", plpsString, , "Mandatory - Enter the help you wish to appear in a popup dialog box"
      .PropertyItems.Add "EXCLUSIONS", "Exclude if...", plpsList, , "Gender specific exclusions"
      .PropertyItems.Add "MIN_AGE", "Minimum Age", plpsNumber, , "The minimum age to which this rule applies"
      .PropertyItems.Add "MAX_AGE", "Maximum Age", plpsNumber, , "The Maximum age to which this rule applies"
      .PropertyItems.Add "URL", "Web Page help", plpsString, , "The URL of the relevant help page"
      With .PropertyItems("NAME")
         .Tag = "Prompt_Desc"
         .max = 50
         .defaultValue = ""
      End With
      .PropertyItems("STYLE").ReadOnly = True
      With .PropertyItems("SAVETYPE").ListItems
         .Add "Clinical Detail", "CD"
         .Add "Test Information", "TI"
      End With
      With .PropertyItems("SAVETYPE")
         .Tag = "Save_As_Type"
         .defaultValue = ""
      End With
      With .PropertyItems("SAVEHEADER")
         .Tag = "Save_As_String"
         .max = 25
      End With
      With .PropertyItems("QUESTION")
         .Tag = "Prompt_Text"
         .max = 140
      End With
      With .PropertyItems("YESTEXT")
         .Tag = "Yes_Text"
         .max = 25
      End With
      With .PropertyItems("NOTEXT")
         .Tag = "No_Text"
         .max = 25
      End With
      With .PropertyItems("YESACTION")
         .Tag = "Yes_Action_Type"
         .value = 0
         .defaultValue = "0 (No action selected)"
      End With
      With .PropertyItems("NOACTION")
         .Tag = "No_Action_Type"
         .value = 0
         .defaultValue = "0 (No action selected"
      End With
      With .PropertyItems("PROMPT")
         .Tag = "Prompt_Text"
         .max = 140
         .defaultValue = ""
      End With
      With .PropertyItems("TYPE")
         .Tag = "Prompt_Type"
         .defaultValue = ""
         With .ListItems
            .Add "Free Text", "FT"
            .Add "Numeric", "N"
            .Add "Date", "D"
            .Add "Time", "T"
            .Add "Date and Time", "DT"
            .Add "Picklist", "P"
         End With
      End With
      
      .PropertyItems("NUMMAX").Enabled = False
      .PropertyItems("NUMMIN").Enabled = False
      .PropertyItems("PICKLIST").Enabled = False
       
      strSQL = "SELECT * " & _
               "FROM Request_Picklist " & _
               "ORDER BY Picklist_Name"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      Do While Not RS.EOF
         .PropertyItems("PICKLIST").ListItems.Add RS!PickList_Name, RS!Picklist_Index
         RS.MoveNext
      Loop
      RS.Close

      With .PropertyItems("HELPTEXT")
         .Tag = "Dialog_Text"
         .max = 255
         .MultiLine = True
      End With
      With .PropertyItems("EXCLUSIONS")
         .ListItems.Add "Neither", 0
         .ListItems.Add "Male", 1
         .ListItems.Add "Female", 2
         .defaultValue = 0
      End With
      With .PropertyItems("MIN_AGE")
         .Min = 0
         .max = 150
         .defaultValue = 0
      End With
      With .PropertyItems("MAX_AGE")
         .Min = 0
         .max = 150
         .defaultValue = 0
      End With
      With .PropertyItems("URL")
         .Tag = "INFO_URL"
      End With
      .Redraw = True
   End With
   mCtrl.SetDbStructures plist
   plist.Pages("Rules").Caption = "|None"
   fView.FrameToShow = Fra_EDI
   Set RS = Nothing
End Sub

Public Property Get ControlObject() As ManageControls
   Set ControlObject = mCtrl
End Property

Public Sub Delete(NodeId As Node)
   On Error GoTo procEH
   Dim i As Integer
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim promptId As String
   Dim strRef As String
   Dim blnDelete As Boolean
   
   blnDelete = True
   promptId = objTV.NodeLevel(NodeId)
   strSQL = "SELECT * " & _
            "FROM Request_Prompt " & _
            "WHERE Yes_Action_Type = " & promptId & _
               " OR No_Action_Type = " & promptId
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      blnDelete = False
      strRef = "The following rules use this rule as a Yes_ or No_ Action:" & vbCrLf & vbCrLf
      Do Until RS.EOF
         strRef = strRef & RS!Prompt_Desc & vbCrLf
         RS.MoveNext
      Loop
      strRef = strRef & vbCrLf & "All references to this rule must be removed before the rule itself can be deleted"
   Else
      RS.Close
      strSQL = "SELECT * " & _
               "FROM Request_Test_Prompts " & _
               "WHERE Prompt_Index = " & promptId
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount > 0 Then
         blnDelete = False
         strRef = "This rule is still used by the following tests: " & vbCrLf & vbCrLf
         Do Until RS.EOF
            strSQL = "SELECT Screen_Caption " & _
                     "FROM Request_Tests " & _
                     "WHERE Test_Index = " & RS!Test_Index
            RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            strRef = strRef & RS2!Screen_Caption & vbCrLf
            RS.MoveNext
            RS2.Close
         Loop
         strRef = strRef & vbCrLf & "All references to this rule must be removed before the rule itself can be deleted"
      End If
   End If
   
   RS.Close
   
   If blnDelete Then
      strSQL = "DELETE FROM Request_Prompt " & _
               "WHERE Prompt_Index = " & promptId
      iceCon.Execute strSQL
      tv1.Nodes.Remove NodeId.Index
   Else
      MsgBox strRef, vbInformation, "Unable to delete rule"
   End If
   
   Set RS = Nothing
   Set RS2 = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadRules.Delete"
   eClass.FurtherInfo = "Node = " & NodeId.Text & " (Id = " & promptId & ")"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Details(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim i As Integer
   Dim strRType As String
   Dim ruleIndex As String
   Dim yAct As String
   Dim nAct As String

   objTV.ActiveNode = NodeId
   strRType = objTV.TopLevelNode(objTV.ActiveNode).Text
   ruleIndex = objTV.NodeLevel(NodeId)
      
   With plist
      .Redraw = False
      For i = 3 To .PropertyItems.Count
         .PropertyItems(i).Visible = False
      Next i
      
      strSQL = "SELECT * " & _
               "FROM Request_Prompt " & _
               "WHERE Prompt_Index = " & ruleIndex
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      objTV.ActiveNode.Text = RS!Prompt_Desc
      With frmMain.edipr
'            .PropertyItems("STYLE").value = Format(RS!Prompt_Type)
         .PropertyItems("STYLE").value = objTV.TopLevelNode(objTV.ActiveNode).Text
         Visibility RS!Prompt_Type
         Select Case RS!Prompt_Type
            Case "QUE"
               If Val(RS!Yes_Action_Type & "") > 0 Then
                  strSQL = "SELECT Prompt_Desc " & _
                           "FROM Request_Prompt " & _
                           "WHERE Prompt_Index = " & RS!Yes_Action_Type
                  RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
                  yAct = RS2!Prompt_Desc
                  RS2.Close
               Else
                  yAct = "No further action"
               End If
               If Val(RS!No_Action_Type & "") > 0 Then
                  strSQL = "SELECT Prompt_Desc " & _
                           "FROM Request_Prompt " & _
                           "WHERE Prompt_Index = " & RS!No_Action_Type
                  RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
                  nAct = RS2!Prompt_Desc
                  RS2.Close
               Else
                  nAct = "No further action"
               End If
               
               .PropertyItems("QUESTION").value = RS!Prompt_Text
               .PropertyItems("QUESTION").Visible = True
               .PropertyItems("YESTEXT").value = RS!Yes_Text
               .PropertyItems("YESTEXT").Visible = True
               .PropertyItems("NOTEXT").value = RS!No_Text
               .PropertyItems("NOTEXT").Visible = True
               .PropertyItems("YESACTION").value = Val(RS!Yes_Action_Type & "") & " (" & yAct & ")"
               .PropertyItems("YESACTION").Visible = True
               .PropertyItems("NOACTION").value = Val(Val(RS!No_Action_Type & "")) & " (" & nAct & ")"
               .PropertyItems("NOACTION").Visible = True
               .PropertyItems("CANCELTEXT").value = RS!Cancel_Text
               .PropertyItems("CANCELTEXT").Visible = True
               .PropertyItems("SAVETYPE").value = Format(RS!Save_As_Type)
               .PropertyItems("SAVETYPE").Visible = True
               .PropertyItems("SAVEHEADER").value = Trim(RS!Save_As_String)
               .PropertyItems("SAVEHEADER").Visible = True
               .PropertyItems("EXCLUSIONS").value = IIf(IsNull(RS!Gender_Exclude), 0, Format(RS!Gender_Exclude))
               .PropertyItems("EXCLUSIONS").Visible = True
               .PropertyItems("MIN_AGE").value = IIf(IsNull(RS!Min_Age), 0, RS!Min_Age)
               .PropertyItems("MIN_AGE").Visible = True
               .PropertyItems("MAX_AGE").value = IIf(IsNull(RS!Max_Age), 0, RS!Max_Age)
               .PropertyItems("MAX_AGE").Visible = True
               .PropertyItems("URL").value = Trim(RS!Info_URL & "")
               .PropertyItems("URL").Visible = True
               
            Case "DEN"
               .PropertyItems("PROMPT").value = RS!Prompt_Text
               .PropertyItems("PROMPT").Visible = True
               .PropertyItems("TYPE").value = Format(RS!DataEntry_Type)
               .PropertyItems("TYPE").Visible = True
               With .PropertyItems("NUMMAX")
                  .value = Format(RS!DataEntry_Upper_Val)
                  .Visible = True
                  .Enabled = False
               End With
               With .PropertyItems("NUMMIN")
                  .value = Format(RS!DataEntry_Lower_Val)
                  .Visible = True
                  .Enabled = False
               End With
               With .PropertyItems("PICKLIST")
                  .value = Format(RS!Picklist_Index)
                  .Visible = True
                  .Enabled = False
               End With
               
               If .PropertyItems("TYPE").value = "P" Then
                  .PropertyItems("PICKLIST").Enabled = True
               ElseIf .PropertyItems("TYPE").value = "N" Then
                  .PropertyItems("NUMMAX").Enabled = True
                  .PropertyItems("NUMMIN").Enabled = True
               End If
               
               .PropertyItems("SAVETYPE").value = Format(RS!Save_As_Type)
               .PropertyItems("SAVETYPE").Visible = True
               .PropertyItems("SAVEHEADER").value = RS!Save_As_String
               .PropertyItems("SAVEHEADER").Visible = True
               .PropertyItems("EXCLUSIONS").value = IIf(IsNull(RS!Gender_Exclude), 0, Format(RS!Gender_Exclude))
               .PropertyItems("EXCLUSIONS").Visible = True
               .PropertyItems("MIN_AGE").value = IIf(IsNull(RS!Min_Age), 0, RS!Min_Age)
               .PropertyItems("MIN_AGE").Visible = True
               .PropertyItems("MAX_AGE").value = IIf(IsNull(RS!Max_Age), 0, RS!Max_Age)
               .PropertyItems("MAX_AGE").Visible = True
               .PropertyItems("URL").value = Trim(RS!Info_URL & "")
               .PropertyItems("URL").Visible = True
            
            Case "HLP"
               .PropertyItems("HELPTEXT").value = RS!Dialog_Text
               .PropertyItems("HELPTEXT").Visible = True
               
         End Select
      End With
      .PropertyItems("NAME").value = RS!Prompt_Desc
      .Redraw = True
   End With
   RS.Close
'   fView.Show Fra_EDI
   Set RS = Nothing
   Exit Sub


procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadRules.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub RuleType(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   objTV.ActiveNode = NodeId
   For i = 1 To NodeId.Children
      tv1.Nodes.Remove NodeId.child.Index
   Next i
   
   vData = objTV.ReadNodeData(NodeId)
   strSQL = "SELECT * " & _
            "FROM Request_Prompt " & _
            "WHERE Prompt_Type = '" & vData(0) & "' " & _
            "ORDER BY Prompt_Desc"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   Do Until RS.EOF
      tv1.Nodes.Add NodeId, _
                    tvwChild, _
                    mCtrl.NewNodeKey(RS!Prompt_Index, _
                                     RS!Prompt_Type, _
                                     "Details", _
                                     , _
                                     ms_DELETE), _
                    RS!Prompt_Desc, _
                    1, _
                    1
      RS.MoveNext
   Loop
   RS.Close
'   fView.Show Fra_HELP, "1"
   Set RS = Nothing
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadRules.Details"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub MenuAddEntry()
   Dim i As Integer
   Dim thisNode As Node
   
   TidyUp
      
   Set newNode = tv1.Nodes.Add(objTV.ActiveNode, _
                               tvwChild, _
                               mCtrl.NewNodeKey("0", _
                                                "NewNode", _
                                                "RuleDetails", _
                                                , _
                                                ms_DELETE, _
                                                True), _
                               "New", _
                               1, _
                               1)
   objTV.ActiveNode = newNode
   
   With frmMain.edipr
      .Redraw = False
      .PropertyItems("STYLE").defaultValue = objTV.TopLevelNode(objTV.ActiveNode).Text
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
         If i > 2 Then
            .PropertyItems(i).Visible = False
         End If
      Next i
      .Redraw = True
      Visibility objTV.NodeLevel(objTV.TopLevelNode(objTV.ActiveNode))
   End With
   fView.Show Fra_EDI
End Sub

Public Sub NewRule()
   Dim i As Integer
   
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(, _
                                  , _
                                  mCtrl.NewNodeKey("0", _
                                                   -1, _
                                                   "Details", _
                                                   , _
                                                   ms_DELETE, _
                                                   True), _
                                  "New", _
                                  1, _
                                  1)
   End If
'   objTView.ActiveNode = newNode
   objTV.ActiveNode = newNode
   
   With plist
      .Redraw = False
      For i = 1 To .PropertyItems.Count
         .PropertyItems(i).value = .PropertyItems(i).defaultValue
         If i > 2 Then
            .PropertyItems(i).Visible = False
         End If
      Next i
      .Redraw = True
   End With
   fView.Show Fra_EDI
End Sub

Public Function Refresh() As String
   Details objTV.ActiveNode
   objTV.RefreshNode = objTV.ActiveNode
End Function

Public Sub RunWhat(NodeId As Node, Optional Refresh As String = "")
   On Error GoTo procEH
   
   If TidyUp = False Then
      vData = objTV.ReadNodeData(NodeId)
      fView.FrameToShow = Fra_EDI
      Select Case vData(2)
         Case "Details"
            Details NodeId
            
         Case "New"
            NewRule
            
         Case "RuleType"
            fView.FrameToShow = Fra_HELP
            RuleType NodeId
            
      End Select
   End If
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadRules.RunWhat"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub ListRules()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   
   strSQL = "SELECT * " & _
            "FROM Request_Prompt_Types " & _
            "ORDER BY Prompt_Type"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Prompt_Type, _
                                                 "PTP", _
                                                 "RuleType", _
                                                 , _
                                                 ms_ADD), _
                                RS!Prompt_Description, _
                                2, _
                                2)
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
   eClass.CurrentProcedure = "IceConfig.LoadRulesMap.ListRuless"
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

Public Sub TypeEntry(Required As String)
   With plist
      Select Case Required
         Case "N"
            .PropertyItems("NUMMAX").Enabled = True
            .PropertyItems("NUMMAX").value = 0
            .PropertyItems("NUMMIN").Enabled = True
            .PropertyItems("NUMMIN").value = 0
            .PropertyItems("PICKLIST").Enabled = False
            .PropertyItems("PICKLIST").value = ""
         
         Case "P"
            .PropertyItems("PICKLIST").Enabled = True
            .PropertyItems("NUMMAX").Enabled = False
            .PropertyItems("NUMMAX").value = ""
            .PropertyItems("NUMMIN").Enabled = False
            .PropertyItems("NUMMIN").value = ""
            
         Case Else
            .PropertyItems("NUMMAX").Enabled = False
            .PropertyItems("NUMMAX").value = ""
            .PropertyItems("NUMMIN").Enabled = False
            .PropertyItems("NUMMIN").value = ""
            .PropertyItems("PICKLIST").Enabled = False
            .PropertyItems("PICKLIST").value = ""
         
      End Select
   End With
End Sub

Public Sub Update(NodeId As Node)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim fieldSQL  As String
   Dim valueSQL As String
   Dim yesAct As String
   Dim noAct As String
   Dim numMax As String
   Dim numMin As String
   Dim pl As String
   Dim NPI As String
   Dim strPType As String
   Dim promptId As String
   
   promptId = objTV.NodeLevel(objTV.ActiveNode)
   strPType = objTV.NodeLevel(objTV.TopLevelNode(objTV.ActiveNode))
   
   With frmMain.edipr
      If objTV.newNode(NodeId) Then ' promptId = "NewNode" Then
         strSQL = "SELECT Max(Prompt_Index) 'MPI' " & _
                  "FROM Request_Prompt"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         
         If Format(RS!MPI) = "" Then
            NPI = 1
         Else
            NPI = RS!MPI + 1
         End If
           
         RS.Close
         
         fieldSQL = "INSERT INTO Request_Prompt (Prompt_Index, Prompt_Desc, Prompt_Type "
         valueSQL = ") Values (" & Format(NPI) & ",'" & .PropertyItems("NAME").value & "', '" & _
                    strPType & "' "
         
         Select Case .PropertyItems("STYLE").value
            Case "QUE"
               fieldSQL = fieldSQL & ", Prompt_Text, Yes_Text, No_Text, Yes_Action_Type, " & _
                                 "No_Action_Type, Cancel_Text, Save_As_Type, Save_As_String, " & _
                                 "Gender_Exclude, Min_Age, Max_Age, INFO_URL "
               
               If Format(.PropertyItems("YESACTION").value & "") = "" Then
                  .PropertyItems("YESACTION").value = "0 (No further action)"
               End If
               
               If Format(.PropertyItems("NOACTION").value & "") = "" Then
                  .PropertyItems("NOACTION").value = "0 (No further action)"
               End If
               
               yesAct = Left(.PropertyItems("YESACTION").value, InStr(1, .PropertyItems("YESACTION"), " (") - 1)
               noAct = Left(.PropertyItems("NOACTION").value, InStr(1, .PropertyItems("NOACTION"), " (") - 1)
'                                 Format (.PropertyItems("YESACTION").value & "") & ", " & _
'                                 Format(.PropertyItems("NOACTION").value & "") & ", '" & _

               valueSQL = valueSQL & ", '" & _
                                 .PropertyItems("QUESTION").value & "', '" & _
                                 .PropertyItems("YESTEXT").value & "', '" & _
                                 .PropertyItems("NOTEXT").value & "', " & _
                                 yesAct & ", " & _
                                 noAct & ", '" & _
                                 .PropertyItems("CANCELTEXT").value & "', '" & _
                                 .PropertyItems("SAVETYPE").value & "', '" & _
                                 .PropertyItems("SAVEHEADER").value & "', " & _
                                 .PropertyItems("EXCLUSIONS").value & ", " & _
                                 .PropertyItems("MIN_AGE").value & ", " & _
                                 .PropertyItems("MAX_AGE").value & ", '" & _
                                 .PropertyItems("URL").value & "')"
            
            Case "DEN"
               fieldSQL = fieldSQL & ", Prompt_Text, DataEntry_Type, DataEntry_Upper_Val, " & _
                                      "DataEntry_Lower_Val, PickList_Index, Save_As_Type, " & _
                                 "Save_As_String, Gender_Exclude, Min_Age, Max_Age, Info_URL "
               If Format(.PropertyItems("NUMMAX").value & "") = "" Then
                  numMax = 0
               Else
                  numMax = Format(.PropertyItems("NUMMAX").value)
               End If
                   
               If Format(.PropertyItems("NUMMIN").value & "") = "" Then
                  numMin = 0
               Else
                  numMin = Format(.PropertyItems("NUMMIN").value)
               End If
               
               If Format(.PropertyItems("PICKLIST").value & "") = "" Then
                  pl = "NULL"
               Else
                  pl = Format(.PropertyItems("PICKLIST").value)
               End If
                   
               valueSQL = valueSQL & ", '" & _
                                    .PropertyItems("PROMPT").value & "', '" & _
                                    .PropertyItems("TYPE").value & "', " & _
                                    numMax & ", " & _
                                    numMin & ", " & _
                                    pl & ", '" & _
                                    .PropertyItems("SAVETYPE").value & "', '" & _
                                    .PropertyItems("SAVEHEADER").value & "', " & _
                                    .PropertyItems("EXCLUSIONS").value & ", " & _
                                    .PropertyItems("MIN_AGE").value & ", " & _
                                    .PropertyItems("MAX_AGE").value & ", '" & _
                                    .PropertyItems("URL").value & "')"
               
            Case "HLP"
               fieldSQL = fieldSQL & ", Dialog_Text"
               valueSQL = valueSQL & ", '" & _
                                     .PropertyItems("HELPTEXT").value + "')"
           
           Case Else
               valueSQL = valueSQL & ")"
               
         End Select
         strSQL = fieldSQL & valueSQL
         
         objTV.ActiveNode.Text = .PropertyItems("NAME").value
         objTV.ActiveNode.Key = mCtrl.NewNodeKey(NPI, _
                                                 .PropertyItems("STYLE").value, _
                                                 "Details", _
                                                 , _
                                                 ms_DELETE)
      Else
         
         strSQL = "UPDATE Request_Prompt SET " & _
                     "Prompt_Desc = '" & .PropertyItems("NAME").value & "', "
         
         Select Case strPType
            
            Case "QUE"
               If Format(.PropertyItems("YESACTION").value & "") = "" Then
                  .PropertyItems("YESACTION").value = ""
               End If
               
               If Format(.PropertyItems("NOACTION").value & "") = "" Then
                  .PropertyItems("NOACTION").value = ""
               End If
               
               yesAct = Left(.PropertyItems("YESACTION").value, InStr(1, .PropertyItems("YESACTION"), " (") - 1)
               noAct = Left(.PropertyItems("NOACTION").value, InStr(1, .PropertyItems("NOACTION"), " (") - 1)
               
'                        "Yes_Action_Type = " & .PropertyItems("YESACTION").value & ", " & _
'                        "No_Action_Type = " & .PropertyItems("NOACTION").value & ", " & _

               
               strSQL = strSQL & _
                        "Prompt_Type = 'QUE', " & _
                        "Prompt_Text = '" & .PropertyItems("QUESTION").value & "', " & _
                        "Yes_Text = '" & .PropertyItems("YESTEXT").value & "', " & _
                        "No_Text = '" & .PropertyItems("NOTEXT").value & "', " & _
                        "Yes_Action_Type = " & yesAct & ", " & _
                        "No_Action_Type = " & noAct & ", " & _
                        "Cancel_Text = '" & .PropertyItems("CANCELTEXT").value & "', " & _
                        "Save_As_Type='" & .PropertyItems("SAVETYPE").value & "', " & _
                        "Save_As_String='" & .PropertyItems("SAVEHEADER").value & "', " & _
                        "Gender_Exclude = " & .PropertyItems("EXCLUSIONS").value & ", " & _
                        "Min_Age = " & .PropertyItems("MIN_AGE").value & ", " & _
                        "Max_Age = " & .PropertyItems("MAX_AGE").value & ", " & _
                        "INFO_URL = '" & .PropertyItems("URL").value & "' "
            Case "DEN"
               If Format(.PropertyItems("NUMMAX").value & "") = "" Then
                  numMax = 0
               Else
                  numMax = Format(.PropertyItems("NUMMAX").value)
               End If
               
               If Format(.PropertyItems("NUMMIN").value & "") = "" Then
                  numMin = 0
               Else
                  numMin = Format(.PropertyItems("NUMMIN").value)
               End If
               
               If Val(numMin) > Val(numMax) Then
                  MsgBox "The numeric range you have entered is invalid.  Please check and try again.", vbInformation + vbOKOnly, "Save Rule"
                  Exit Sub
               End If
               
               If Format(.PropertyItems("PICKLIST").value & "") = "" Then
                  pl = "NULL"
               Else
                  pl = Format(.PropertyItems("PICKLIST").value)
               End If
               
               strSQL = strSQL & _
                        "Prompt_Type = 'DEN', " & _
                        "Prompt_Text = '" & .PropertyItems("PROMPT").value & "', " & _
                        "DataEntry_Type = '" & .PropertyItems("TYPE").value & "', " & _
                        "DataEntry_Upper_Val = " & numMax & ", " & _
                        "DataEntry_Lower_Val = " & numMin & ", " & _
                        "PickList_Index = " & pl & ", " & _
                        "Save_As_Type = '" & .PropertyItems("SAVETYPE").value & "', " & _
                        "Save_As_String = '" & .PropertyItems("SAVEHEADER").value & "', " & _
                        "Gender_Exclude = " & .PropertyItems("EXCLUSIONS").value & ", " & _
                        "Min_Age = " & .PropertyItems("MIN_AGE").value & ", " & _
                        "Max_Age = " & .PropertyItems("MAX_AGE").value & ", " & _
                        "INFO_URL = '" & .PropertyItems("URL").value & "' "
                
            Case "MCD"
               strSQL = strSQL & "Prompt_Type = 'MCD' "
               
            Case "EIF"
               strSQL = strSQL & "Prompt_Type = 'EIF'"
                   
            Case "EIM"
               strSQL = strSQL & "Prompt_Type = 'EIM' "
           
            Case "CNL"
               strSQL = strSQL & "Prompt_Type = 'CNL' "

            Case "HLP"
               strSQL = strSQL & "Prompt_Type = 'HLP', " & _
                                 "Dialog_Text = '" & .PropertyItems("HELPTEXT").value & "' "
                        
         End Select
         strSQL = strSQL & "WHERE Prompt_Index = " & promptId
      End If
   End With
   iceCon.Execute strSQL
   Set newNode = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadRules.Update"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Visibility(qType As String)
   Dim i As Integer
   
   With frmMain.edipr
      .Redraw = False
      For i = 3 To .PropertyItems.Count - 2
         .PropertyItems(i).Visible = False
         .PropertyItems(i).PageKeys = ""
      Next i
      .PropertyItems("STYLE").value = qType
      Select Case qType
         Case "QUE"
            .PropertyItems("QUESTION").Visible = True
            .PropertyItems("QUESTION").PageKeys = "Rules"
            .PropertyItems("YESTEXT").Visible = True
            .PropertyItems("YESTEXT").PageKeys = "Rules"
            .PropertyItems("NOTEXT").Visible = True
            .PropertyItems("NOTEXT").PageKeys = "Rules"
            .PropertyItems("YESACTION").Visible = True
            .PropertyItems("YESACTION").PageKeys = "Rules"
            .PropertyItems("NOACTION").Visible = True
            .PropertyItems("NOACTION").PageKeys = "Rules"
            .PropertyItems("CANCELTEXT").Visible = True
            .PropertyItems("CANCELTEXT").PageKeys = "Rules"
            .PropertyItems("SAVETYPE").Visible = True
            .PropertyItems("SAVETYPE").PageKeys = "Rules"
            .PropertyItems("SAVEHEADER").Visible = True
            .PropertyItems("SAVEHEADER").PageKeys = "Rules"
            .PropertyItems("EXCLUSIONS").Visible = True
            .PropertyItems("MIN_AGE").Visible = True
            .PropertyItems("MAX_AGE").Visible = True
            .PropertyItems("URL").Visible = True
            
         Case "DEN"
            .PropertyItems("PROMPT").Visible = True
            .PropertyItems("PROMPT").PageKeys = "Rules"
            .PropertyItems("TYPE").Visible = True
            .PropertyItems("TYPE").PageKeys = "Rules"
            With .PropertyItems("NUMMAX")
               .Visible = True
               .Enabled = False
            End With
            With .PropertyItems("NUMMIN")
               .Visible = True
               .Enabled = False
            End With
            With .PropertyItems("PICKLIST")
               .Visible = True
               .Enabled = False
            End With
            
            If .PropertyItems("TYPE").value = "P" Then
               .PropertyItems("PICKLIST").Enabled = True
            ElseIf .PropertyItems("TYPE").value = "N" Then
               .PropertyItems("NUMMAX").Enabled = True
               .PropertyItems("NUMMIN").Enabled = True
            End If
            
            .PropertyItems("SAVETYPE").Visible = True
            .PropertyItems("SAVETYPE").PageKeys = "Rules"
            .PropertyItems("SAVEHEADER").Visible = True
            .PropertyItems("SAVEHEADER").PageKeys = "Rules"
            .PropertyItems("EXCLUSIONS").Visible = True
            .PropertyItems("MIN_AGE").Visible = True
            .PropertyItems("MAX_AGE").Visible = True
            .PropertyItems("URL").Visible = True
         
         Case "HLP"
            .PropertyItems("HELPTEXT").Visible = True
            .PropertyItems("HELPTEXT").PageKeys = "Rules"
            
      End Select
      .Redraw = True
   End With
End Sub

