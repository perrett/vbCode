VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Begin VB.Form frmBehaviour 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Behaviour Editor"
   ClientHeight    =   7650
   ClientLeft      =   3675
   ClientTop       =   2775
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   6015
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10610
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      DisableIcons    =   -1  'True
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6675
      Left            =   90
      TabIndex        =   4
      Top             =   465
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11774
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Entry"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fixed Rules"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Questions"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   7260
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   780
      TabIndex        =   1
      Top             =   7260
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double-click the behaviour you wish to use."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmBehaviour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tv1 As TreeView
Private mCtrl As ManageControls
Private blnReturnData As Boolean
Private propIndex As String
Private strSQL As String
Private propRow As Integer

Private Sub Command1_Click()
   vbalGrid1_DblClick vbalGrid1.SelectedRow, vbalGrid1.SelectedCol
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Public Property Let ControlObject(objNewValue As ManageControls)
   Set mCtrl = objNewValue
End Property

Private Sub Form_Load()
   Set tv1 = frmMain.TreeView1
   If TypeName(loadCtrl) = "LoadRules" Then
      blnReturnData = True
   Else
      blnReturnData = False
   End If
   TabStrip1_Click
'   LoadRules
End Sub

Public Property Let InitialValue(intNewValue As Integer)
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   propRow = intNewValue
   If propRow > 0 Then
      strSQL = "SELECT Prompt_Desc, Prompt_Type " & _
               "FROM Request_Prompt " & _
               "WHERE Prompt_Index = " & propRow
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      Select Case RS!Prompt_Type
         Case "DEN"
            TabStrip1.Tabs(1).Selected = True
         
         Case "HLP"
            TabStrip1.Tabs(3).Selected = True
            
         Case "QUE"
            TabStrip1.Tabs(4).Selected = True
            
         Case Else
            TabStrip1.Tabs(2).Selected = True
         
      End Select
      RS.Close
   Else
      TabStrip1.Tabs(2).Selected = True
   End If
   TabStrip1_Click
   For i = 1 To vbalGrid1.rows
      If vbalGrid1.Cell(i, 2).Text = propRow Then
         Exit For
      End If
   Next i
   vbalGrid1.SelectedRow = i
End Property

Public Property Let ReturnIndex(strNewValue As String)
   propIndex = strNewValue
End Property

Private Sub TabStrip1_Click()
   Select Case TabStrip1.SelectedItem
      Case "Data Entry"
         LoadRules "Prompt_Type = 'DEN'"
         
      Case "Fixed Rules"
         If blnReturnData Then
            LoadRules "Prompt_Type = 'ACP' OR Prompt_Type = 'CNL' " & _
                      "OR Prompt_Type = 'MCD'"
         Else
            LoadRules "Prompt_Type = 'EIF' OR Prompt_Type = 'EIM' " & _
                      "OR Prompt_Type = 'MCD'"
         End If
         
      Case "Help"
         LoadRules "Prompt_Type = 'HLP'"
      
      Case "Questions"
         LoadRules "Prompt_Type = 'QUE'"
         
   End Select
End Sub

Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   Dim TVN As Node
   Dim tIndex As String
   
   If lRow > 0 Then
      If blnReturnData Then
         frmMain.edipr(propIndex).value = vbalGrid1.Cell(vbalGrid1.SelectedRow, 2).Text & " (" & _
                                          vbalGrid1.Cell(vbalGrid1.SelectedRow, 1).Text & ")"
      Else
         tIndex = objTV.NodeLevel(tv1.SelectedItem)
         On Error GoTo DupNode
         If tv1.SelectedItem.Children Then
            tv1.Nodes.Add tv1.SelectedItem.child.LastSibling, _
                          tvwLast, _
                          mCtrl.NewNodeKey(vbalGrid1.Cell(vbalGrid1.SelectedRow, 2).Text, _
                                           tIndex, _
                                           "RuleDetails", _
                                           , _
                                           ms_DELETE), _
                          vbalGrid1.Cell(vbalGrid1.SelectedRow, 1).Text, _
                          1, _
                          1
         Else
            tv1.Nodes.Add tv1.SelectedItem, _
                          tvwChild, _
                          mCtrl.NewNodeKey(vbalGrid1.Cell(vbalGrid1.SelectedRow, 2).Text, _
                                           tIndex, _
                                           "RuleDetails", _
                                           , _
                                           ms_DELETE), _
                          vbalGrid1.Cell(vbalGrid1.SelectedRow, 1).Text, _
                          1, _
                          1
         End If
      End If
      On Error GoTo 0
   End If
   Unload Me
   Exit Sub
DupNode:
   MsgBox "This rule is already present on the selected test and cannot be added again", vbOKOnly + vbInformation, "Cannot Add Rule"
   On Error GoTo 0
   Exit Sub
End Sub

Private Sub LoadRules(PromptCondition As String)
   Dim RS As New ADODB.Recordset
   Dim noBehaviours As Integer
   
   With vbalGrid1
      .Clear True
      .AddColumn "COL1", "Behaviour", ecgHdrTextALignCentre
      .AddColumn "COL2", "ID", ecgHdrTextALignCentre, , , False
      If TabStrip1.SelectedItem = "Fixed Rules" And blnReturnData Then
         .CellDetails 1, 1, "No further action", DT_LEFT + DT_NOCLIP
         .CellDetails 1, 2, 0, DT_CENTER
         noBehaviours = 1
      Else
         noBehaviours = 0
      End If
      strSQL = "SELECT Prompt_Desc, Prompt_Index, Prompt_Type " & _
               "FROM Request_Prompt " & _
               "WHERE (" & PromptCondition & ") " & _
               "AND Prompt_Index NOT IN (" & _
                  "SELECT Prompt_Index " & _
                  "FROM Request_Test_Prompts " & _
                  "WHERE Test_Index = " & objTV.NodeLevel(objTV.ActiveNode) & ") " & _
               "ORDER BY Prompt_Type, Prompt_Desc"
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount > 0 Then
         .rows = RS.RecordCount
'         nobehaviours = 1
         Do While Not RS.EOF
            noBehaviours = noBehaviours + 1
            .CellDetails noBehaviours, 1, RS!Prompt_Desc, DT_LEFT + DT_NOCLIP
            .CellDetails noBehaviours, 2, RS!Prompt_Index, DT_CENTER
            RS.MoveNext
         Loop
      End If
      RS.Close
      .AutoWidthColumn "COL1"
   End With
End Sub
