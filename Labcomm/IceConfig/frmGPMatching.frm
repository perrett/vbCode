VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGPMatching 
   Caption         =   "Practice System National Code (Org)"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   360
      Left            =   1860
      TabIndex        =   6
      Top             =   4185
      Width           =   3195
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laboratory Identification"
      Height          =   2910
      Left            =   3525
      TabIndex        =   9
      Top             =   1020
      Width           =   3270
      Begin VB.CheckBox chkSMTP 
         Alignment       =   1  'Right Justify
         Caption         =   "SMTP Active"
         Height          =   255
         Left            =   225
         TabIndex        =   14
         Top             =   1635
         Width           =   2265
      End
      Begin VB.CheckBox chkActive 
         Alignment       =   1  'Right Justify
         Caption         =   "Match Active:"
         Height          =   210
         Left            =   225
         TabIndex        =   13
         Top             =   2025
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.TextBox txtOPName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   4
         Top             =   1170
         Width           =   1605
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   330
         Left            =   825
         TabIndex        =   5
         Top             =   2430
         Width           =   2025
      End
      Begin VB.TextBox txtKey3 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   3
         Top             =   765
         Width           =   1575
      End
      Begin VB.TextBox txtKey1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1515
         TabIndex        =   2
         Top             =   375
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Output Name:"
         Height          =   270
         Left            =   225
         TabIndex        =   12
         Top             =   1229
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "GP Code:"
         Height          =   270
         Left            =   225
         TabIndex        =   11
         Top             =   817
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Practice Code:"
         Height          =   285
         Left            =   225
         TabIndex        =   10
         Top             =   390
         Width           =   1080
      End
   End
   Begin VB.ComboBox cboLTS 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Width           =   3330
   End
   Begin MSComctlLib.TreeView tvGP 
      Height          =   3360
      Left            =   135
      TabIndex        =   0
      Top             =   525
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   5927
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.Label Label6 
      Caption         =   "Local Trader"
      Height          =   285
      Left            =   4515
      TabIndex        =   8
      Top             =   225
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "GP Details"
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   255
      Width           =   1110
   End
   Begin VB.Menu mnuNode 
      Caption         =   "Node"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show All"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmGPMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private natCode As String
Private gpId As Long
Private gpName As String
Private GPNatCode As String
Private strSQL As String
Private mCtrl As New ManageControls
Private vData As Variant
Private sqlLK1_3 As String
Private sqlLK2 As String
Private tNode As Node
Private newNode As Node
Private callId As String
Private newCallId As String
Private blnAmended As Boolean

Public Property Let IndividualId(lngNewValue As Long)
   gpId = lngNewValue
End Property

Public Property Let GPDetails(strNewValue As String)
   Dim strArray() As String
   
   strArray = Split(strNewValue, "|")
   gpId = Val(strArray(0))
   gpName = strArray(1)
   GPNatCode = strArray(2)
   natCode = strArray(3)
End Property

Private Sub PrepareTreeView()
   Dim RS As New ADODB.Recordset
   Dim gpNode As Node
   Dim lastKey3 As String
   Dim key3Node As Node
   Dim lastKey1 As String
   Dim key1Node As Node
   Dim outputAs As String
   
   tvGP.Nodes.Clear
   Set gpNode = tvGP.Nodes.Add(, _
                               , _
                               mCtrl.NewNodeKey("GP", GPNatCode, "GP"), _
                               GPNatCode & " (" & gpName & ")")
   gpNode.BackColor = &H80FFFF

   strSQL = "SELECT MatchIndex, EDI_Local_Key1, EDI_Local_Key3, EDI_Op_Name, EDI_SMTP_Active, Matching_Active " & _
            "FROM EDI_Matching em " & _
               "LEFT JOIN EDI_Local_Trader_Settings ET " & _
               "ON EDI_Local_Key2 = EDI_LTS_Index " & _
            "WHERE individual_Index  = " & gpId & _
               " AND EDI_Local_Key2 = '" & cboLTS.ItemData(cboLTS.ListIndex) & "' " & _
            "ORDER BY EDI_Local_Key1, EDI_Local_Key3"
   RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
   
   Do Until RS.EOF
      outputAs = ""
      
      If RS!EDI_Local_Key1 <> lastKey1 Then
         lastKey1 = RS!EDI_Local_Key1
         Set key1Node = tvGP.Nodes.Add(gpNode, _
                                       tvwChild, _
                                       mCtrl.NewNodeKey(lastKey1, natCode, "KEY1"), _
                                       lastKey1)
         key1Node.Bold = True
         key1Node.Expanded = True
      End If
      
      If RS!EDI_OP_Name <> "" Then
         If RS!EDI_OP_Name <> gpName Then
            outputAs = " (" & RS!EDI_OP_Name & ")"
         End If
      End If
      
      Set key3Node = tvGP.Nodes.Add(key1Node, _
                                    tvwChild, _
                                    mCtrl.NewNodeKey(RS!MatchIndex, _
                                                     RS!EDI_Local_Key3, _
                                                     "KEY3", _
                                                     RS!EDI_OP_Name & ""), _
                                    RS!EDI_Local_Key3 & outputAs)
      key3Node.Tag = Abs(CInt(IIf(IsNull(RS!EDI_SMTP_Active), False, RS!EDI_SMTP_Active)))
      If RS!Matching_Active Then
         key3Node.ForeColor = BPGREEN
      Else
         key3Node.ForeColor = BPRED
      End If
      
      RS.MoveNext
   Loop
   
   gpNode.Expanded = True
   tvGP_NodeClick gpNode
'   key1Node.Expanded = True
   RS.Close
   Set RS = Nothing
End Sub

Private Sub cboLTS_Click()
   PrepareTreeView
End Sub

Private Sub cmdApply_Click()
   On Error GoTo procEH
   Dim rNum As Integer
   Dim rNode As Node
   Dim strArray() As String
   Dim strDuplicates As String
   Dim strMsg As String
   Dim RS As ADODB.Recordset
   Dim iceCmd As New ADODB.Command
   Dim i As Integer
   Dim sqlRes As Long
   
   If Not newNode Is Nothing Then
      vData = objTV.ReadNodeData(newNode)
   End If
   
   If txtKey1.Text = "" Then
      MsgBox "A practice key is required", vbExclamation, "Validation"
      blnAmended = False
   Else
      blnAmended = (txtKey1.Text <> vData(0))
      
      If txtKey3.Text = "" And callId = "KEY3" Then
         MsgBox "A gp key is required", vbExclamation, "Validation"
      Else
         blnAmended = blnAmended Or (txtKey3.Text <> vData(1))
      End If
      
      If tNode.ForeColor = BPGREEN Then
         blnAmended = blnAmended Or (chkActive.value = 0)
      Else
         blnAmended = blnAmended Or (chkActive.value = 1)
      End If
      
      blnAmended = blnAmended Or (txtOPName.Text <> vData(4))
   End If
   
   strSQL = "SELECT DISTINCT EDI_Org_Natcode " & _
            "FROM EDI_Matching em " & _
               "INNER JOIN EDI_Recipient_Individuals er " & _
               "ON em.Individual_Index = er.Individual_Index " & _
            "WHERE EDI_Local_Key1 = '" & txtKey1.Text & "' " & _
               "AND EDI_Local_Key2 = '" & cboLTS.ItemData(cboLTS.ListIndex) & "' " & _
               "AND EDI_Org_NatCode <> '" & natCode & "'"
   Set RS = iceCon.Execute(strSQL)
   
   Do Until RS.EOF
      strDuplicates = strDuplicates & RS!EDI_Org_NatCode
      RS.MoveNext
   Loop
      
   If strDuplicates <> "" Then
      MsgBox "Key1 already used for practice " & strDuplicates & " with this datastream", vbInformation, "Duplicate Key"
      strDuplicates = ""
      Set newNode = Nothing
   Else
      If newNode Is Nothing Then
         If blnAmended Then
            With iceCmd
               .ActiveConnection = iceCon
               .CommandType = adCmdStoredProc
               .CommandText = "ICECONFIG_EDI_MatchingUpdate"
               .Parameters.Append .CreateParameter("mIndex", adInteger, adParamInput, , 0)
               .Parameters.Append .CreateParameter("gpId", adInteger, adParamInput, , gpId)
               .Parameters.Append .CreateParameter("Key2", adInteger, adParamInput, , cboLTS.ItemData(cboLTS.ListIndex))
               .Parameters.Append .CreateParameter("Key1", adVarChar, adParamInput, 30, "")
               .Parameters.Append .CreateParameter("Key3", adVarChar, adParamInput, 30, "")
               .Parameters.Append .CreateParameter("gpName", adVarChar, adParamInput, 35, gpName)
               .Parameters.Append .CreateParameter("Act", adBoolean, adParamInput, , (chkActive.value = 1))
               .Parameters.Append .CreateParameter("SMTP", adBoolean, adParamInput, , (chkSMTP.value = 1))
               .Parameters.Append .CreateParameter("rData", adVarChar, adParamOutput, 250, "")
               .Parameters.Append .CreateParameter("res", adInteger, adParamReturnValue, , sqlRes)
            End With
            
            If callId = "KEY1" Then
               Set rNode = tNode.child
               With iceCmd
                  Do Until rNode Is Nothing
                     vData = objTV.ReadNodeData(rNode)
                     .Parameters("Key1").value = txtKey1.Text
                     .Execute
                     
                     If .Parameters("res").value <> 0 Then
                        strDuplicates = strDuplicates & .Parameters("rData")
                     End If
      '               strDuplicates = strDuplicates & .Parameters("data").value
                     
                     Set rNode = rNode.Next
                  Loop
               End With
               
               If strDuplicates <> "" Then
                  strArray = Split(strDuplicates, "|")
                  strMsg = "Amending this Local Practice Code would create the following duplicate entries: " & vbCrLf
                  With iceCmd
                     For i = 0 To UBound(strArray) Step 2
                        
                        strMsg = strMsg & "Practice: " & strArray(i) & " GP: " & strArray(i + 1) & _
                                 "(" & strArray(i + 2) & ")" & vbCrLf
                     Next i
                  End With
                  
                  MsgBox strMsg, vbExclamation, "Duplicate record(s) found"
                  blnAmended = False
               End If
               
            Else
               
               With iceCmd
                  .Parameters("mIndex").value = vData(0)
                  .Parameters("Key3").value = txtKey3.Text
                  .Parameters("gpName").value = txtOPName.Text
   '               Dim i As Integer
                  For i = 0 To .Parameters.Count - 1
                     Debug.Print .Parameters(i).Name & " - " & .Parameters(i).value
                  Next i
                  .Execute
                  
                  If .Parameters("res").value <> 0 Then
                     strArray = Split(.Parameters("rData").value, "|")
                     MsgBox "A match has already been set up for these details." & vbCrLf & _
                            "See " & strArray(0) & ": " & strArray(2) & " (" & strArray(1) & ")", _
                            vbExclamation, "Duplicate Record Found"
                     blnAmended = False
                  End If
               End With
            End If
         End If
      Else
         If objTV.newNode Then
            loadCtrl.Update "IN"
         End If
         
         With iceCmd
            .ActiveConnection = iceCon
            .CommandType = adCmdStoredProc
            .CommandText = "ICECONFIG_EDI_MatchNew"
            .Parameters.Append .CreateParameter("gpId", adInteger, adParamInput, , gpId)
            .Parameters.Append .CreateParameter("Key1", adVarChar, adParamInput, 30, txtKey1.Text)
            .Parameters.Append .CreateParameter("Key2", adInteger, adParamInput, , cboLTS.ItemData(cboLTS.ListIndex))
            .Parameters.Append .CreateParameter("Key3", adVarChar, adParamInput, 30, txtKey3.Text)
            .Parameters.Append .CreateParameter("OpName", adVarChar, adParamInput, 35, txtOPName.Text)
            .Parameters.Append .CreateParameter("ACR", adBoolean, adParamInput, , True)
            .Parameters.Append .CreateParameter("SMTP", adBoolean, adParamInput, , (chkSMTP.value = 1))
            .Parameters.Append .CreateParameter("rData", adVarChar, adParamOutput, 250, "")
            .Parameters.Append .CreateParameter("res", adInteger, adParamReturnValue, , sqlRes)
            .Execute
            
            If .Parameters("res").value <> 0 Then
               strArray = Split(.Parameters("rData").value, "|")
               MsgBox "A match has already been set up for these details." & vbCrLf & _
                      "See " & strArray(0) & ": " & strArray(2) & " (" & strArray(1) & ")", _
                      vbExclamation, "Duplicate Record Found"
               blnAmended = False
            End If
         End With
         
         If objTV.newNode Then
            loadCtrl.Refresh
            loadCtrl.SubNodes objTV.RefreshNode
         End If
         
         Set newNode = Nothing
      End If
   End If
   
   PrepareTreeView
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmGPMatching.cmdApply_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub cmdCancel_Click()
   Set newNode = Nothing
   Unload Me
End Sub

Private Sub CmdOk_Click()
   Set newNode = Nothing
   Unload Me
End Sub

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   
   cboLTS.Clear
   strSQL = "SELECT EDI_OrgCode, EDI_Msg_Type, EDI_LTS_Index " & _
            "FROM EDI_Local_Trader_Settings " & _
            "WHERE Organisation = '" & frmMain.cboTrust.Text & "'"
   RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
   Do Until RS.EOF
      cboLTS.AddItem RS!EDI_OrgCode & " - " & RS!EDI_Msg_Type
      cboLTS.ItemData(cboLTS.ListCount - 1) = RS!EDI_LTS_Index
      RS.MoveNext
   Loop
   
   RS.Close
   cboLTS.ListIndex = 0
   PrepareTreeView
   frmGPMatching.Caption = "Practice System: " & natCode & " (" & frmMain.ediPr("NA").value & ")"
   Set RS = Nothing
End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   Effect = vbDropEffectNone
End Sub

Private Sub mnuAdd_Click()
   Dim nodeKey As String
   
   txtKey3.Text = GPNatCode
   txtOPName.Text = gpName
   
   If callId = "GP" Then
      txtKey1.Text = natCode
      nodeKey = "newKEY1"
      newCallId = "KEY1"
   
   Else
      newCallId = "KEY3"
      nodeKey = "newKEY3"
   End If
   
   Set newNode = tvGP.Nodes.Add(tNode, _
                                tvwChild, _
                                mCtrl.NewNodeKey(natCode, nodeKey, newCallId), _
                                nodeKey)
   If newCallId = "KEY1" Then
      EditKey1 True
   Else
      EditKey3 True
   End If
   
   tvGP.SelectedItem = newNode
End Sub

Private Sub tvAl_NodeClick(ByVal Node As MSComctlLib.Node)
   vData = objTV.ReadNodeData(Node)
   txtKey1.Text = vData(0)
   txtKey3.Text = vData(1)
   sqlLK1_3 = "AND EDI_Local_Key1 = '" & vData(0) & "' " & _
              "AND EDI_Local_Key3 = '" & vData(1) & "' "
End Sub

Private Sub mnuDelete_Click()
   Dim pNode As Node
   Dim RS As New ADODB.Recordset
   
   If MsgBox("This will remove the matching details - continue?", vbExclamation Or vbYesNo, _
             "Confirm delete") = vbYes Then
      If callId = "KEY1" Then
         strSQL = "SELECT Count(*) " & _
                  "FROM EDI_Matching " & _
                  "WHERE EDI_Local_Key1 = '" & vData(0) & "' " & _
                     "AND EDI_Local_Key2 = '" & cboLTS.ItemData(cboLTS.ListIndex) & "' "
         Set RS = iceCon.Execute(strSQL)
         If RS(0) = 0 Then
            strSQL = "DELETE FROM EDI_Matching " & _
                     "WHERE EDI_Local_Key1 = '" & vData(0) & "' " & _
                        "AND EDI_Local_Key2 = '" & cboLTS.ItemData(cboLTS.ListIndex) & "' "
         Else
            MsgBox "Key references " & RS(0) & " individuals" & vbCrLf & _
                   "Please remove these before deleting this Key", vbInformation, "Key in use - Not Removed"
            strSQL = ""
         End If
         
         RS.Close
      Else
         strSQL = "DELETE FROM EDI_Matching " & _
                  "WHERE MatchIndex= " & vData(0)
      End If
      
      If strSQL <> "" Then
         iceCon.Execute strSQL
         Set pNode = objTV.TopLevelNode(tNode)
         tvGP.Nodes.Remove tNode.Index
         Set tNode = pNode
      End If
      
      Set RS = Nothing
   End If
End Sub

Private Sub tvGP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Set tNode = tvGP.HitTest(x, y)
   If Not tNode Is Nothing Then
      tvGP.SelectedItem = tNode
      tvGP_NodeClick tNode
      If Not newNode Is Nothing Then
         tvGP.Nodes.Remove newNode.Index
         Set newNode = Nothing
      End If
      
      vData = objTV.ReadNodeData(tNode)
      callId = vData(2)
      
      If Button = vbRightButton Then
         Select Case callId
            Case "GP"
               mnuAdd.Caption = "New Key 1"
               mnuAdd.Visible = True
               mnuDelete.Visible = False
               PopupMenu mnuNode
               
            Case "KEY1"
               mnuAdd.Caption = "New Key 3"
               mnuAdd.Visible = True
               mnuDelete.Caption = "Delete " & tNode.Text
               mnuDelete.Visible = True
               PopupMenu mnuNode
            
            Case "KEY3"
               mnuDelete.Caption = "Delete " & tNode.Text
               mnuDelete.Visible = True
               mnuAdd.Visible = False
               PopupMenu mnuNode
            
         End Select
      End If
   End If
End Sub

Private Sub tvGP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim tNode As Node
   Dim vData As Variant
   
   Set tNode = tvGP.HitTest(x, y)
   If Not tNode Is Nothing Then
      vData = objTV.ReadNodeData(tNode)
      If vData(2) = "GP" Then
         tvGP.ToolTipText = "Right click to add a new Local Practice Key"
      ElseIf vData(2) = "KEY1" Then
         tvGP.ToolTipText = "Right click to add a new Local GP Key or delete this Practice Key"
      Else
         tvGP.ToolTipText = "Right click to delete this local GP Key"
      End If
   End If
End Sub

Private Sub tvGP_NodeClick(ByVal Node As MSComctlLib.Node)
   vData = objTV.ReadNodeData(Node)
   
   If vData(2) = "KEY3" Then
      txtKey1.Text = Node.Parent.Text
      txtKey3.Text = vData(1)
      EditKey3 True
      
'      With txtKey1
'         .Text = vData(0)
'         .BackColor = &H8000000F
'         .Enabled = False
'      End With
'
'      txtKey3.Text = vData(1)
'      txtKey3.BackColor = &H80000005
      chkSMTP.value = Node.Tag
      txtOPName.Text = vData(4)
      txtOPName.BackColor = &H80000005
      chkActive.value = IIf(Node.ForeColor = BPGREEN, 1, 0)
   
   ElseIf vData(2) = "KEY1" Then
      EditKey1 True
      txtKey1.Text = vData(0)
'      With txtKey1
'         .Text = vData(0)
'         .BackColor = &H80000005
'         .Enabled = True
'      End With
'
'      txtKey3.Text = ""
'      txtKey3.BackColor = &H8000000F
'      txtOPName.Text = ""
'      txtOPName.BackColor = &H8000000F
   Else
      With txtKey1
         .Text = ""
         .BackColor = &H8000000F
         .Enabled = False
      End With
      
      txtKey3.Text = ""
      txtKey3.BackColor = &H8000000F
      
      txtOPName.Text = ""
      txtOPName.BackColor = &H8000000F
      
      chkActive.Visible = False
      chkSMTP.Visible = False
   End If
   
'   txtKey3.Enabled = (txtKey3.Text <> "")
'   txtOPName.Enabled = (txtKey3.Text <> "")
End Sub

Private Sub EditKey1(SelectField As Boolean)
   With txtKey1
      .Enabled = True
      .BackColor = &H80000005
      
      If SelectField Then
         .SelStart = 0
         .SelLength = Len(.Text)
      End If
      
      .SetFocus
   End With
   
   If newNode Is Nothing Then
      With txtKey3
         .Text = ""
         .Enabled = False
         .BackColor = &H8000000F
      End With
      
      With txtOPName
         .Text = ""
         .Enabled = False
         .BackColor = &H8000000F
      End With
   Else
      With txtKey3
         .Enabled = True
         .BackColor = &H80000005
      End With
      
      With txtOPName
         .Enabled = True
         .BackColor = &H80000005
      End With
   End If
   
   chkActive.Visible = False
   chkSMTP.Visible = False
End Sub

Private Sub EditKey3(SelectField As Boolean)
   With txtKey1
      .Enabled = False
      .BackColor = &H8000000F
   End With
   
   With txtKey3
      .Enabled = True
      .BackColor = &H80000005
      
      If SelectField Then
         .SelStart = 0
         .SelLength = Len(.Text)
      End If
      
      .SetFocus
   End With
   
   With txtOPName
      .Text = gpName
      .Enabled = True
      .BackColor = &H80000005
   End With
   chkActive.Visible = True
   chkSMTP.Visible = True
End Sub
