VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalgrid6.ocx"
Begin VB.Form frmIncExcRfx 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Test"
   ClientHeight    =   5250
   ClientLeft      =   4695
   ClientTop       =   3690
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   4095
      Left            =   150
      TabIndex        =   1
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7223
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double-click on the test you wish to add."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmIncExcRfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "SELECT Screen_Caption, Test_Index " & _
            "FROM Request_Tests " & _
            "ORDER BY Screen_Caption"
   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      vbalGrid1.AddColumn "COL1", "Test"
      vbalGrid1.AddColumn "COL2", "Index", , , , False
      vbalGrid1.Rows = RS.RecordCount
      NoTests = 0
      Do While Not RS.EOF
         NoTests = NoTests + 1
         vbalGrid1.CellDetails NoTests, 1, RS!Screen_Caption & ""
         vbalGrid1.CellDetails NoTests, 2, Format(RS!Test_Index)
         RS.MoveNext
      Loop
      vbalGrid1.AutoWidthColumn "COL1"
   End If
   RS.Close
End Sub

Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)

   Dim strArray() As String
   Dim tempKey As String
   
    If lRow < 1 Then Exit Sub
    Dim TVN As node
    Set TVN = frmMain.TreeView1.SelectedItem
    If TVN.Children > 0 Then
        Set TVN = TVN.Child
        If UCase(TVN.Text) = UCase(vbalGrid1.Cell(lRow, 1).Text) Then
            MsgBox "This test has already been added and cannot be added again", vbOKOnly, "Add Test"
            Exit Sub
        End If
        For i = 1 To frmMain.TreeView1.SelectedItem.Children - 1
            Set TVN = TVN.Next
            If UCase(TVN.Text) = UCase(vbalGrid1.Cell(lRow, 1).Text) Then
                MsgBox "This test has already been added and cannot be added again", vbOKOnly, "Add Test"
                Exit Sub
            End If
        Next i
    End If
    strArray = Split(objTView.NodeLevel(frmMain.TreeView1.SelectedItem.Key), ":")
    NPI = strArray(2)
    tempKey = objTView.AddNode(frmMain.AddMode.Caption, vbalGrid1.Cell(lRow, 2).Text, "", 1)
    If strArray(3) = "O" Then
'        NPI = Mid$(frmMain.TreeView1.SelectedItem.Key, 2, Len(frmMain.TreeView1.SelectedItem.Key) - 1)
        TempStr = "Insert Into Request_Profile_Tests (Profile_Index,Profile_Test_Index) Values ("
        TempStr = TempStr + NPI + "," + vbalGrid1.Cell(lRow, 2).Text + ")"
        Debug.Print TempStr
        ICECon.Execute TempStr
        frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, tempKey, vbalGrid1.Cell(lRow, 1).Text, 1, 1
    Else
'        NPI = Mid$(frmMain.TreeView1.SelectedItem.Parent.Key, 2, Len(frmMain.TreeView1.SelectedItem.Parent.Key) - 1)
        If vbalGrid1.Cell(lRow, 1).Text = frmMain.TreeView1.SelectedItem.Parent.Text Then
            MsgBox "You cannot add a reference to yourself", vbInformation + vbOKOnly, "Test Configuration"
            Exit Sub
        End If
        If CheckDuplicates(frmMain.AddMode.Caption, vbalGrid1.Cell(lRow, 2).Text, NPI) Then Exit Sub
        tempKey = objTView.AddNode(frmMain.AddMode.Caption, vbalGrid1.Cell(lRow, 2).Text, "", 1)
        Select Case frmMain.AddMode.Caption
            Case "I"
                TempStr = "Insert Into Request_Included_Tests (Test_Index,Included_Test_Index) Values (" & NPI & "," & vbalGrid1.Cell(lRow, 2).Text & ")"
                If frmMain.TreeView1.SelectedItem.Children > 0 Then
                    frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem.Child, tvwLast, tempKey, vbalGrid1.Cell(lRow, 1).Text, 1, 1
                Else
                    frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, tempKey, vbalGrid1.Cell(lRow, 1).Text, 1, 1
                End If
            Case "X"
                TempStr = "Insert Into Request_Excluded_Tests (Test_Index,Excluded_Test_Index) Values (" & NPI & "," & vbalGrid1.Cell(lRow, 2).Text & ")"
                If frmMain.TreeView1.SelectedItem.Children > 0 Then
                    frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem.Child, tvwLast, tempKey, vbalGrid1.Cell(lRow, 1).Text, 1, 1
                Else
                    frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, tempKey, vbalGrid1.Cell(lRow, 1).Text, 1, 1
                End If
            Case "R"
                TempStr = "Insert Into Request_Reflex_Tests (Test_Index,Reflex_Test_Index) Values (" & NPI & "," & vbalGrid1.Cell(lRow, 2).Text & ")"
                If frmMain.TreeView1.SelectedItem.Children > 0 Then
                    frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem.Child, tvwLast, tempKey, vbalGrid1.Cell(lRow, 1).Text, 1, 1
                Else
                    frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, tempKey, vbalGrid1.Cell(lRow, 1).Text, 1, 1
                End If
        End Select
        ICECon.Execute TempStr
    End If
    Unload Me
End Sub

Private Function CheckDuplicates(AddMode As String, NTI As String, NPI) As Boolean
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    CheckDuplicates = False
    Select Case AddMode
        Case "I"
            Expanded = "Included"
        Case "X"
            Expanded = "Excluded"
        Case "R"
            Expanded = "Reflex"
    End Select
    If AddMode <> "I" Then
        RS.Open "Select * From Request_Included_Tests Where Test_Index=" & NPI & " And Included_Test_Index=" & NTI, ICECon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            CheckDuplicates = True
            MsgBox "This test is already setup as an Included test and cannot be setup as a " & Expanded & " test.", vbInformation + vbOKOnly, "Add " & Expanded & " Test"
        End If
        RS.Close
    End If
    If AddMode <> "X" Then
    RS.Open "Select * From Request_Excluded_Tests Where Test_Index=" & NPI & " And Excluded_Test_Index=" & NTI, ICECon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            CheckDuplicates = True
            MsgBox "This test is already setup as an Excluded test and cannot be setup as a " & Expanded & " test.", vbInformation + vbOKOnly, "Add " & Expanded & " Test"
        End If
        RS.Close
    End If
    If AddMode <> "R" Then
    RS.Open "Select * From Request_Reflex_Tests Where Test_Index=" & NPI & " And Reflex_Test_Index=" & NTI, ICECon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            CheckDuplicates = True
            MsgBox "This test is already setup as a Reflex test and cannot be setup as a " & Expanded & " test.", vbInformation + vbOKOnly, "Add " & Expanded & " Test"
        End If
        RS.Close
    End If
End Function
