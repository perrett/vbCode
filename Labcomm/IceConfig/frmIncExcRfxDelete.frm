VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalgrid6.ocx"
Begin VB.Form frmIncExcRfxDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Test"
   ClientHeight    =   5220
   ClientLeft      =   4845
   ClientTop       =   3900
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   4920
      Width           =   1455
   End
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7011
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
      Caption         =   "Double-click the test you wish to delete.  When you have finshed click OK."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmIncExcRfxDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim RS As New ADODB.Recordset
    Dim strarray() As String
    
    If lRow < 1 Then Exit Sub
    NPI = objTView.NodeKey(frmMain.TreeView1.SelectedItem.Key)
    strarray = Split(objTView.NodeLevel(frmMain.TreeView1.SelectedItem.Key), ":")
'    NPI = Mid$(frmMain.TreeView1.SelectedItem.Key, 2, Len(frmMain.TreeView1.SelectedItem.Key) - 1)
    If frmMain.TreeView1.SelectedItem.Children > 0 Then
        Dim TVN As Node
        For i = 1 To frmMain.TreeView1.SelectedItem.Children
            frmMain.TreeView1.Nodes.Remove (frmMain.TreeView1.SelectedItem.Child.Index)
        Next i
    End If
    
    If strarray(3) = "O" Then
'    If Left(frmMain.TreeView1.SelectedItem.Key, 1) = "O" Then
        ICECon.Execute "Delete From Request_Profile_Tests Where Profile_Index=" & NPI & " And Profile_Test_Index=" & vbalGrid1.Cell(lRow, 2).Text
        RS.Open "Select Screen_Caption,Request_Profile_Tests.Profile_Index,Profile_Test_Index From Request_Profile_Tests,Request_Tests Where Request_Tests.Test_Index=Request_Profile_Tests.Profile_Test_Index And Request_Profile_Tests.Profile_Index=" & NPI, ICECon, adOpenKeyset, adLockReadOnly
        Do While Not RS.EOF
            frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, "O" + NPI + "-" + Format(RS!Profile_Test_Index), RS!Screen_Caption, 1, 1
            RS.MoveNext
        Loop
        RS.Close
    Else
        Select Case frmMain.AddMode.Caption
            Case "I"
                ICECon.Execute "Delete From Request_Included_Tests Where Test_Index=" & NPI & " And Included_Test_Index=" & vbalGrid1.Cell(lRow, 2).Text
                RS.Open "Select Request_Included_Tests.Test_Index,Included_Test_Index,Screen_Caption From Request_Included_Tests,Request_Tests Where Request_Included_Tests.Included_Test_Index=Request_Tests.Test_Index And Request_Included_Tests.Test_Index=" & NPI, ICECon, adOpenKeyset, adLockReadOnly
                If RS.RecordCount > 0 Then
                    Do While Not RS.EOF
                        tempKey = objTView.AddNode(strarray(3), Format(RS!Included_Test_Index), RS!Screen_Caption, 1)
                        frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, tempKey, RS!Screen_Caption, 1, 1
                        RS.MoveNext
                    Loop
                End If
                RS.Close
            Case "X"
                ICECon.Execute "Delete From Request_Excluded_Tests Where Test_Index=" & NPI & " And Excluded_Test_Index=" & vbalGrid1.Cell(lRow, 2).Text
                RS.Open "Select Request_Excluded_Tests.Test_Index,Excluded_Test_Index,Screen_Caption From Request_Excluded_Tests,Request_Tests Where Request_Excluded_Tests.Excluded_Test_Index=Request_Tests.Test_Index And Request_Excluded_Tests.Test_Index=" & NPI, ICECon, adOpenKeyset, adLockReadOnly
                If RS.RecordCount > 0 Then
                    Do While Not RS.EOF
                        tempKey = objTView.AddNode(strarray(3), Format(RS!Excluded_Test_Index), RS!Screen_Caption, 1)
                        frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, "E" + NPI + "-" + Format(RS!Excluded_Test_Index), RS!Screen_Caption, 1, 1
                        RS.MoveNext
                    Loop
                End If
                RS.Close
            Case "R"
                ICECon.Execute "Delete From Request_Reflex_Tests Where Test_Index=" & NPI & " And Reflex_Test_Index=" & vbalGrid1.Cell(lRow, 2).Text
                RS.Open "Select Request_Reflex_Tests.Test_Index,Reflex_Test_Index,Screen_Caption From Request_Reflex_Tests,Request_Tests Where Request_Reflex_Tests.Reflex_Test_Index=Request_Tests.Test_Index And Request_Reflex_Tests.Test_Index=" & NPI, ICECon, adOpenKeyset, adLockReadOnly
                If RS.RecordCount > 0 Then
                    Do While Not RS.EOF
                        tempKey = objTView.AddNode(strarray(3), Format(RS!Reflex_Test_Index), RS!Screen_Caption, 1)
                        frmMain.TreeView1.Nodes.Add frmMain.TreeView1.SelectedItem, tvwChild, tempKey, RS!Screen_Caption, 1, 1
                        RS.MoveNext
                    Loop
                End If
                RS.Close
        End Select
    End If
    vbalGrid1.RemoveRow lRow
End Sub
