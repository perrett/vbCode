VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tv1 As TreeView
Private plist As PropertiesList
Private nd(3) As Node
Private strSQL As String
Private newNode As Node
Private mCtrl As New ManageControls

Private Sub Class_Initialize()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   
   Set tv1 = frmMain.TreeView1
   Set plist = frmMain.edipr

   With plist
      .PropertyItems.Clear
      .Pages.Clear
      .Pages.Add "Main", "Request_Profiles"
      .PropertyItems.Add "WARD", "Ward", plpsList, , "Select the ward to which this test profile is available"
      .PropertyItems.Add "PR_ENABLE", "Enabled", plpsBoolean, False, "Is this profile enabled"
      .PropertyItems.Add "PROFILE_COLOUR", "Profile Colour", plpsColor, , "Select the colour of the text on the profile button"
      .PropertyItems.Add "POSITION", "Position", plpsNumber, 1, "Enter the position that the profile appears in"
      .PropertyItems.Add "CAPTION", "Caption", plpsString, , "Enter the caption that is displayed to the user"
      .PropertyItems.Add "HELP", "Help Text", plpsString, , "Enter the help detail that is displayed on screen when the user selects the profile"
      .PropertyItems.Add "HELP_COLOUR", "Help Colour", plpsColor, , "Select the background colour of the help string"
      .PropertyItems.Add "TEXT_STRING", "Text String", plpsString, , "Enter the text that appears on the printed request to show this profile has been selected"
   End With
      
   With frmMain
      With .edipr("WARD")
         .Tag = "Profile_Location_Code"
         strSQL = "SELECT Clinic_Name, Localation_Code " & _
                  "FROM Location " & _
                  "WHERE Localation_Code LIKE '" & Trim(Left(frmMain.OrgList.Text, 6)) & "%' " & _
                  "ORDER BY Clinic_Name"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         Do While Not RS.EOF
            .ListItems.Add IIf(IsNull(RS!Clinic_Name), "<None>", RS!Clinic_Name), RS!Localation_Code
            RS.MoveNext
         Loop
         RS.Close
         Set RS = Nothing
      End With
      With .edipr("PR_ENABLE")
         .Tag = "Enabled"
      End With
      With .edipr("PROFILE_COLOUR")
         .Tag = "Profile_Colour"
         .defaultValue = 1
      End With
      With .edipr("POSITION")
         .Tag = "Profile_Position"
         .max = 10
         .Min = 1
         .Increment = 1
         .defaultValue = 0
      End With
      With .edipr("CAPTION")
         .Tag = "Profile_Caption"
      End With
      With .edipr("HELP")
         .Tag = "Profile_Help"
      End With
      With .edipr("HELP_COLOUR")
         .Tag = "Profile_Help_BackColour"
         .defaultValue = 1
      End With
      With .edipr("TEXT_STRING")
         .Tag = "Profile_TextString"
      End With
   End With
   mCtrl.SetDbStructures plist
   fView.FrameToShow = Fra_EDI
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.IceConfig.Class_Initialize"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub Delete(NodeId As Node)
   Dim strSQL As String
   Dim vData As Variant
   
   vData = objTV.ReadNodeData(NodeId)
   If vData(1) = "PRF" Then
      strSQL = "DELETE FROM Request_Profiles " & _
               "WHERE Profile_Index = " & vData(0)
   Else
      strSQL = "DELETE FROM Request_Profile_Tests " & _
               "WHERE Profile_Index = " & vData(0) & _
                  " AND Profile_Test_index = " & vData(1)
   End If
      
   iceCon.Execute strSQL
   frmMain.TreeView1.Nodes.Remove NodeId.Index
End Sub

Private Sub Details(Index As String)
'      strSQL = "SELECT Screen_Caption, Request_Profile_Tests.Profile_Test_Index " & _
'               "FROM Request_Tests, Request_Profile_Tests " & _
'               "WHERE Request_Tests.Test_Index = Request_Profile_Tests.Profile_Test_Index " & _
'                  "And Request_Profile_Tests.Profile_Index = " & Format(RS!Profile_Index) & _
'               " ORDER BY Screen_Caption"
'      RS2.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'      Do While Not RS2.EOF
'         tv1.Nodes.Add nd(0), _
'                       tvwChild, _
'                       mctrl.newnodekey(Format(RS!Profile_Index), "SCREEN_CAPTION", "None", , ms_DELETE), _
'                       RS2!Screen_Caption, _
'                       1, _
'                       1
'         RS2.MoveNext
'      Loop
'      RS2.Close
'
End Sub

Public Sub LoadOrgProfiles(OrgID As String)
   On Local Error GoTo procEH
   Dim strWard As String
   Dim sPos As Integer
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   
'   objctrl.TreeviewOrigin = "O"
'   objctrl.PracticeId = strDataGroup
      
   Set nd(0) = tv1.Nodes.Add(, _
                             , _
                             mCtrl.NewNodeKey("New", "New", "NewProfile"), _
                             "New Profile...", _
                             1, _
                             1)
   nd(0).Bold = True
   
   strSQL = "SELECT Profile_Caption, Profile_Index, Profile_Location_Code " & _
            "FROM Request_Profiles " & _
            "WHERE Profile_Location_Code LIKE '" & OrgID & "%' " & _
            "ORDER BY Profile_Location_Code, Profile_Caption"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do While Not RS.EOF
      sPos = InStr(1, RS!Profile_Location_Code, " ") + 1
      strWard = Mid(RS!Profile_Location_Code, sPos)
      Set nd(0) = tv1.Nodes.Add(, _
                                , _
                                mCtrl.NewNodeKey(RS!Profile_Index, _
                                                    "PRF", _
                                                    "Profile", _
                                                    , _
                                                    ms_BOTH), _
                                strWard, _
                                1, _
                                1)

      strSQL = "SELECT Screen_Caption, Request_Profile_Tests.Profile_Test_Index " & _
               "FROM Request_Tests, Request_Profile_Tests " & _
               "WHERE Request_Tests.Test_Index = Request_Profile_Tests.Profile_Test_Index " & _
                  "And Request_Profile_Tests.Profile_Index = " & Format(RS!Profile_Index) & _
               " ORDER BY Screen_Caption"
      RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      Do While Not RS2.EOF
         tv1.Nodes.Add nd(0), _
                       tvwChild, _
                       mCtrl.NewNodeKey(Format(RS!Profile_Index), RS2!Profile_Test_Index, "PRF", , ms_DELETE), _
                       RS2!Screen_Caption, _
                       1, _
                       1
         RS2.MoveNext
      Loop
      RS2.Close
      RS.MoveNext
   Loop
   RS.Close
   Exit Sub

procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadProfiles.LoadProfiles"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub MenuAddEntry()
   On Error GoTo procEH
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim testId As String
   Dim thisNode As Node
   Dim i As Integer
   
   
   testId = objTV.NodeLevel(objTV.TopLevelNode(tv1.SelectedItem))
   Set thisNode = tv1.SelectedItem
   frmTestMnt.NodeType = objTV.nodeKey(thisNode)
   frmTestMnt.CurrentTest = testId
   frmTestMnt.Show 1
      
   For i = 1 To thisNode.Children
      tv1.Nodes.Remove thisNode.child.Index
   Next i
   
   strSQL = "SELECT Screen_Caption, Request_Profile_Tests.Profile_Test_Index " & _
            "FROM Request_Tests, Request_Profile_Tests " & _
            "WHERE Request_Tests.Test_Index = Request_Profile_Tests.Profile_Test_Index " & _
               "And Request_Profile_Tests.Profile_Index = " & testId & _
            " ORDER BY Screen_Caption"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do While Not RS.EOF
      tv1.Nodes.Add objTV.TopLevelNode(thisNode), _
                    tvwChild, _
                    mCtrl.NewNodeKey(testId, RS!Profile_Test_Index, "PRF", , ms_DELETE), _
                    RS!Screen_Caption, _
                    1, _
                    1
      RS.MoveNext
   Loop
   RS.Close
   
   
'   loadCtrl.RefreshChildNodes thisNode '  objtv.TopLevelNode(tv1.SelectedItem, True)
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadProfiles.MenuAddEntry"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub NewProfile()
   Dim i As Integer
   
   For i = 1 To plist.PropertyItems.Count
      plist(i).value = plist(i).defaultValue
   Next i
   If newNode Is Nothing Then
      Set newNode = tv1.Nodes.Add(, _
                                   , _
                                   mCtrl.NewNodeKey("New", "NewNode", "Profile", , , True), _
                                   "New", _
                                   1, _
                                   1)
   End If
   objTV.ActiveNode = newNode
   fView.Show Fra_EDI, "Profiles"
End Sub

Private Sub Profile(Index As String, _
                    Optional blnReWrite As Boolean = True)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim strWard As String
   Dim sPos As Integer
   Dim i As Integer
    
   strSQL = "SELECT * " & _
            "FROM Request_Profiles " & _
            "WHERE Profile_Index=" & Index
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      
   With frmMain
      .edipr("PR_ENABLE").value = RS!Enabled
      .edipr("WARD").value = RS!Profile_Location_Code & ""
      .edipr("POSITION").value = Val(RS!Profile_Position)
      .edipr("CAPTION").value = RS!Profile_Caption & ""
      If Format(RS!Profile_Colour) & "" <> "" Then
         strSQL = "SELECT Colour_Code, Colour_Name " & _
                  "FROM Colours " & _
                  "WHERE Colour_Index = " & RS!Profile_Colour
         RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         If RS2.RecordCount > 0 Then
            .edipr("PROFILE_COLOUR").value = Val(RS2!Colour_Code)
            .edipr("PROFILE_COLOUR").DialogTitle = Format(RS!Profile_Colour) & "-" & RS2!Colour_Name
         Else
            .edipr("PROFILE_COLOUR").value = Val("&HFFFFFF&")
            .edipr("PROFILE_COLOUR").DialogTitle = "0" + "-None Set"
         End If
         RS2.Close
         
'         If RS2.BOF = False And RS2.EOF = False Then
'            .ediPr("PROFILE_COLOUR").value = Val(RS2!Colour_Code)
'            .ediPr("PROFILE_COLOUR").Description = Format(RS!Profile_Colour) + "-" + RS2!Colour_Name
'         Else
'            .ediPr("PROFILE_COLOUR").value = Format(RS!Profile_Colour)
'         End If
'         RS2.Close
      End If
      .edipr("HELP").value = RS!Profile_Help
      If Format(RS!Profile_Help_Backcolour) & "" <> "" Then
         strSQL = "SELECT Colour_Code, Colour_Name " & _
                  "FROM Colours " & _
                  "WHERE Colour_Index = " & RS!Profile_Help_Backcolour
         RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         If RS2.RecordCount > 0 Then
            .edipr("HELP_COLOUR").value = Val(RS2!Colour_Code)
            .edipr("HELP_COLOUR").DialogTitle = Format(RS!Profile_Help_Backcolour) & "-" & RS2!Colour_Name
         Else
            .edipr("HELP_COLOUR").value = Val("&HFFFFFF&")
            .edipr("HELP_COLOUR").DialogTitle = "0" + "-None Set"
         End If
         RS2.Close
         
'         If RS2.EOF = False And RS2.BOF = False Then
'            .ediPr("HELP_COLOUR").value = Val(RS2!Colour_Code)
'            .ediPr("HELP_COLOUR").Description = Format(RS!Profile_Help_Backcolour) + "-" + RS2!Colour_Name
'         Else
'            .ediPr("HELP_COLOUR").value = RS!Profile_Help_Backcolour
'         End If
'         RS2.Close
      End If
      .edipr("TEXT_STRING").value = RS!Profile_TextString & ""
      .edipr.Caption = "WHERE Profile_Index = " & Index
      If blnReWrite Then
         sPos = InStr(1, RS!Profile_Location_Code, " ") + 1
         strWard = Mid(RS!Profile_Location_Code, sPos)
         objTV.ActiveNode.Text = strWard
         objTV.ActiveNode.Key = mCtrl.NewNodeKey(RS!Profile_Index, _
                                                      "PRF", _
                                                      "Profile", _
                                                      , _
                                                      ms_BOTH)
      End If
      RS.Close
   End With
   
   fView.Show Fra_EDI, "Profiles"
   Set RS2 = Nothing
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.LoadProfiles.Profile"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Function Refresh() As String
   Dim RS As New ADODB.Recordset
   
   If Left(objTV.ActiveNode.Text, 3) = "New" Then
      strSQL = "SELECT max(Profile_Index) " & _
               "FROM Request_Profiles"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      Refresh = RS(0)
      objTV.ActiveNode.Text = plist("WARD").value
      objTV.ActiveNode.Key = mCtrl.NewNodeKey(RS(0), _
                                                    "PRF", _
                                                    "Profile", _
                                                    , _
                                                    ms_BOTH)
      RS.Close
   Else
      Refresh = objTV.NodeLevel(objTV.ActiveNode)
   End If
   Set RS = Nothing
   objTV.ActiveNode = objTV.TopLevelNode(objTV.ActiveNode)
   objTV.RefreshNode = objTV.ActiveNode
End Function

Public Sub RunWhat(NodeId As MSComctlLib.Node, _
                   Optional Refresh As String = "")
   On Error GoTo procEH
   Dim vData As Variant
   
   vData = objTV.ReadNodeData(NodeId)
   If TidyUp = False Then
      Select Case vData(2)
         Case "Profile"
            If IsNumeric(Refresh) Then
               Profile Refresh
            Else
               Profile CStr(vData(0))
            End If
            
         Case "NewProfile"
            NewProfile
            
         Case "PRF"
            Profile CStr(vData(0)), False
            
      End Select
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
