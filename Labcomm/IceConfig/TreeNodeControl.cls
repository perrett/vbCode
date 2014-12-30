VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeNodeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private nKey As String
Private strOrigin As String
Private aNode As Node
Private rNode As Node
Private pNode As Node
Private menuFlag As Long

Public Property Let ActiveNode(objNewValue As Node)
   Set aNode = objNewValue
End Property

Public Property Get ActiveNode() As Node
   Set ActiveNode = aNode
End Property

Public Function NodeFromKey(keyId As String, _
                            Optional StartNode As Node = Nothing) As Node
   Dim tNode As Node
   Dim rNode As Node
   
   If StartNode Is Nothing Then
      Set tNode = aNode
   Else
      Set tNode = StartNode
   End If
   
   Set rNode = Nothing
   Set tNode = tNode.child
   Do Until tNode Is Nothing
      If nodeKey(tNode) = keyId Then
         Set rNode = tNode
         Exit Do
      End If
      Set tNode = tNode.Next
   Loop
   Set NodeFromKey = rNode
End Function

Public Function NodeFromText(Text As String, _
                             Optional StartNode As Node = Nothing) As Node
   Dim tNode As Node
   Dim rNode As Node
   
   If StartNode Is Nothing Then
      Set tNode = frmMain.TreeView1.Nodes(1)
   Else
      Set tNode = StartNode.child
   End If
   
   Set rNode = Nothing
   Do Until tNode Is Nothing
      If tNode.Text = Text Then
         Set rNode = tNode
         Exit Do
      End If
      Set tNode = tNode.Next
   Loop
   Set NodeFromText = rNode
End Function

Public Property Let MenuStatus(eNewValue As ENUM_MenuStatus)
   menuFlag = eNewValue
   Select Case menuFlag
      Case -1
         frmMain.itemAdd.Enabled = True
         frmMain.itemDelete.Enabled = True
      Case 1
         frmMain.itemAdd.Enabled = True
         frmMain.itemDelete.Enabled = False
      Case 2
         frmMain.itemAdd.Enabled = False
         frmMain.itemDelete.Enabled = True
   End Select
End Property

Public Property Get MenuStatus() As ENUM_MenuStatus
   MenuStatus = menuFlag
End Property

Public Function newNode(Optional NodeId As Node = Nothing, _
                        Optional setValue As String = "") As Boolean
   Dim nKey As String
   Dim sPos As Integer
   Dim ePos As Integer
   Dim tNode As Node
   
   If NodeId Is Nothing Then
      Set tNode = aNode
      nKey = aNode.Key
   Else
      Set tNode = NodeId
      nKey = NodeId.Key
   End If
   sPos = InStr(nKey, "!*") + 2
   ePos = InStr(sPos, nKey, "*!") - sPos
   If ePos > 0 Then
      If setValue = "" Then
         newNode = (Mid(nKey, sPos, ePos) = "Y")
      Else
         tNode.Key = Left(nKey, sPos) & setValue & Mid(nKey, sPos + 1)
         newNode = False
      End If
   Else
      newNode = False
   End If
End Function

Public Property Get nodeKey(NodeId As Node) As String
   Dim sPos As Integer
   Dim ePos As Integer
   
   nKey = NodeId.Key
   sPos = InStr(nKey, "!<") + 2
   ePos = InStr(sPos, nKey, ">!") - sPos
   nodeKey = Mid(nKey, sPos, ePos)
End Property

Public Property Get NodeLevel(NodeId As Node) As String
   Dim sPos As Integer
   Dim ePos As Integer
   Dim nKey As String

   nKey = NodeId.Key
   sPos = InStr(nKey, "!#") + 2
   ePos = InStr(sPos, nKey, "#!") - sPos
   NodeLevel = Mid(nKey, sPos, ePos)
End Property

Public Function NodeOrigin(Optional NodeId As Node = Nothing) As String
   Dim sPos As Integer
   Dim ePos As Integer
   
   If NodeId Is Nothing Then
      NodeOrigin = strOrigin
   Else
      nKey = NodeId.Key
      sPos = InStr(nKey, "!@") + 2
      ePos = InStr(sPos, nKey, "@!") - sPos
      
      strOrigin = Mid(nKey, sPos, ePos)
      NodeOrigin = strOrigin
   End If
End Function

Public Function NodeRunWhat(Optional NodeId As Node = Nothing)
   Dim tNode As Node
   
   If NodeId Is Nothing Then
      Set tNode = aNode
   Else
      Set tNode = NodeId
   End If
   
End Function

Public Function OldValue(PropertyItemKey As String, _
                         Optional ParentNodeId As Node = Nothing) As String
   Dim tNode As Node
   
   If ParentNodeId Is Nothing Then
      Set tNode = aNode
   Else
      Set tNode = ParentNodeId
   End If
   
   OldValue = ""
   If Not (tNode.child Is Nothing) Then
      Set tNode = tNode.child
      Do Until tNode.Next Is Nothing
         If nodeKey(tNode) = PropertyItemKey Then
            OldValue = NodeLevel(tNode)
            Exit Do
         End If
         Set tNode = tNode.Next
      Loop
      Set tNode = Nothing
   End If
End Function

Public Property Let Origin(strNewValue As String)
   strOrigin = strNewValue
End Property

Public Property Let ParentNode(NodeId As Node)
   Set pNode = NodeId
End Property

Public Property Get ParentNode() As Node
   Set ParentNode = pNode
End Property

Public Function ReadNodeData(Optional NodeId As Node = Nothing) As Variant
   On Error GoTo NotPresent
   Dim nKey As String
   Dim sPos As Integer
   Dim ePos As Integer
   Dim keyId As String
   Dim nLevel As String
   Dim Extra As String
   Dim callMethod As String
   Dim mStat As String
   
   If NodeId Is Nothing Then
      nKey = aNode.Key
   Else
      nKey = NodeId.Key
   End If
   
   sPos = InStr(nKey, "!#") + 2
   ePos = InStr(sPos, nKey, "#!") - sPos
   nLevel = Mid(nKey, sPos, ePos)
   
   sPos = InStr(nKey, "!$") + 2
   ePos = InStr(sPos, nKey, "$!") - sPos
   callMethod = Mid(nKey, sPos, ePos)
   
   sPos = InStr(nKey, "!<") + 2
   ePos = InStr(sPos, nKey, ">!") - sPos
   keyId = Mid(nKey, sPos, ePos)
   
   sPos = InStr(nKey, "!+") + 2
   ePos = InStr(sPos, nKey, "+!") - sPos
   Extra = Mid(nKey, sPos, ePos)
   
   ePos = InStr(sPos, nKey, ">!") + 2
   mStat = Mid(nKey, ePos)
   
NotPresent: '  Do not fail is item is not present
   ReadNodeData = Array(nLevel, keyId, callMethod, mStat, Extra)
End Function

Public Property Let RefreshNode(NodeId As Node)
   Set rNode = NodeId
End Property

Public Property Get RefreshNode() As Node
   Set RefreshNode = rNode
End Property

Public Sub ResetActive(NodeId As Node)
   Dim i As Integer
   
   For i = 1 To NodeId.Children
      If NodeId.child.Text = aNode.Text Then
         Set aNode = NodeId.child
         NodeId.child.EnsureVisible
         Exit For
      End If
   Next i
End Sub

Public Property Let SQL(strNewValue As String)
   strSQL = strNewValue
End Property

Public Function TopLevelNode(Optional NodeId As Node = Nothing, _
                             Optional ReturnChild As Boolean = False) As MSComctlLib.Node
   Dim pNode As MSComctlLib.Node
   Dim tNode As MSComctlLib.Node
   
   If NodeId Is Nothing Then
      Set NodeId = aNode
   End If
   
   Set pNode = NodeId
   
   If NodeId.Children > 0 Then
      Set tNode = NodeId.child
   End If
   
   Do Until TypeName(pNode.Parent) = "Nothing"
      Set tNode = pNode
      Set pNode = pNode.Parent
   Loop
   
   If ReturnChild Then
      Set TopLevelNode = tNode
   Else
      Set TopLevelNode = pNode
   End If
End Function

Public Sub UpdatePropList(pl As PropertiesList, _
                          ActivePage As String, _
                          mCtrl As ManageControls, _
                          Optional NodeId As Node = Nothing)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim vData As Variant
   
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   With pl
      .ActivePage = ActivePage
      For i = 1 To pl.PropertyItems.Count
         With .PropertyItems(i)
            If .PageKeys = ActivePage And _
               .Visible And _
               .Tag <> "MENU" And _
               .Tag <> "" Then
               vData = RS(.Tag)
               
               If TypeName(vData) = "String" Then
                  vData = Trim(vData)
               End If
                  
               If TypeName(vData) = "Null" Then
                  If .Style = plpsBoolean Then
                     vData = False
                  ElseIf .Style = plpsNumber Then
                     vData = 0
                  Else
                     vData = ""
                  End If
               End If
               
               Select Case .Style
                     Case 5
                        .value = CLng(vData)
                     
                     Case 8
                        .value = CStr(vData)
                     
                     Case 11
                        .value = CBool(vData)
                     
                     Case 104
                        .value = CLng(vData)
                     
                     Case Else
                        .value = vData
                        
               End Select
            
            End If
         End With
      Next i
   End With
   RS.Close
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "TreeNodeControl.UpdatePropList"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Sub UpdateTreeView(pl As PropertiesList, _
                          ActivePage As String, _
                          mCtrl As ManageControls, _
                          Optional NodeId As Node = Nothing)
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim vData As Variant
   Dim tNode As Node
   Dim subNode As Node
   Dim i As Integer
   Dim iconId As Long
   Dim tv1 As TreeView
   
   Set tv1 = frmMain.TreeView1
   
   If NodeId Is Nothing Then
      Set tNode = aNode
   Else
      Set tNode = NodeId
   End If
   
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   
   With pl
      .ActivePage = ActivePage
      For i = 1 To pl.PropertyItems.Count
         With .PropertyItems(i)
'            Debug.Print "Key: " & .Key & " Tag: " & .Tag & " Page: " & .PageKeys
            If .PageKeys = ActivePage And _
               .Visible Then
               iconId = pl("ICON").Icon
               
               If .Tag = "MENU" Then
                  Set subNode = tv1.Nodes.Add(tNode, _
                                              tvwChild, _
                                              mCtrl.NewNodeKey("Dummy", _
                                                               .Key, _
                                                               .DefaultExtension, _
                                                               , _
                                                               .Flags), _
                                              .Caption, _
                                              iconId, _
                                              iconId)
                  tv1.Nodes.Add subNode, _
                                tvwChild, _
                                mCtrl.NewNodeKey("Tmp", "Tmp"), _
                                "Please wait...", _
                                iconId, _
                                iconId
               ElseIf .Tag <> "" Then
                  
                  If RS.EOF Then
                     vData = ""
                  Else
                     vData = RS(.Tag)
                  End If
               
                  If TypeName(vData) = "String" Then
                     vData = Trim(vData)
                  End If
                  
                  If TypeName(vData) = "Null" Then
                     If .Style = plpsBoolean Then
                        vData = False
                     ElseIf .Style = plpsNumber Then
                        vData = 0
                     Else
                        vData = ""
                     End If
                  End If
                  
                  Select Case .Style
                     Case 5
                        .value = CLng(vData)
                     
                     Case 8
                        .value = CStr(vData)
                     
                     Case 11
                        If vData = "" Then
                           .value = False
                        Else
                           .value = CBool(vData)
                        End If
                     Case 104
                        .value = CLng(vData)
                     
                     Case Else
                        .value = vData
                        
                  End Select
'                  .value = vData
               
                  If .OverlayIcon > 0 Then
                     If .value = True Then
                        iconId = .OverlayIcon
                     End If
                  End If
               
                  Set subNode = tv1.Nodes.Add(tNode, _
                                              tvwChild, _
                                              mCtrl.NewNodeKey(CStr(vData), _
                                                               .Key, _
                                                               .DefaultExtension, _
                                                               , _
                                                               .Flags), _
                                              .Caption & " - " & .value, _
                                              iconId, _
                                              iconId)
                  
               End If
            End If
         End With
      Next i
   End With
   RS.Close
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "TreeNodeControl.UpdateTreeView"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub
