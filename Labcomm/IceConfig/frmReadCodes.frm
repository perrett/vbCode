VERSION 5.00
Begin VB.Form frmReadCodes 
   Caption         =   "Read Codes"
   ClientHeight    =   3735
   ClientLeft      =   4335
   ClientTop       =   3720
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3735
   ScaleWidth      =   8400
   Begin VB.TextBox txtComment 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "frmReadCodes.frx":0000
      Top             =   3240
      Width           =   7935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test Details"
      Height          =   1695
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   2175
      Begin VB.Label lblDetails 
         Caption         =   "Release Date"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblDetails 
         Caption         =   "Status"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblDetails 
         Caption         =   "Read ratio"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblDetails 
         Caption         =   "Battery/Header"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ListBox lstHide 
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   6030
      TabIndex        =   7
      Top             =   2745
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   3600
      TabIndex        =   6
      Top             =   2745
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "Version"
      Height          =   1665
      Left            =   5640
      TabIndex        =   2
      Top             =   255
      Width           =   2385
      Begin VB.OptionButton optRC 
         Caption         =   "Version 3 Read Codes"
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optRC 
         Caption         =   "Version 2 Read Codes"
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1905
      End
      Begin VB.OptionButton optRC 
         Caption         =   "4-Byte Read Codes"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0E0FF&
      Height          =   2205
      Left            =   360
      TabIndex        =   1
      Top             =   900
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2385
   End
   Begin VB.Label Label2 
      Caption         =   "Descriptive Text"
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Read Code selected"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   1545
   End
End
Attribute VB_Name = "frmReadCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private keyCount As Integer
Private bText As String
Private pos As Integer
Private returnKey As String
Private searchKey As String
Private RS As New ADODB.Recordset
Private assignedRC As String
Private rcIndex As Integer
Private strSQL As String
Private frmCaption As String

Private Sub cmdCancel_Click()
   Set RS = Nothing
   Unload Me
End Sub

Private Sub CmdOk_Click()
   If Text2.Text = "" And Text1.Text <> "" Then
      MsgBox "Please select the read code description required", vbExclamation, _
             "No read code selected"
   Else
      frmMain.edipr(frmMain.edipr.Tag).value = Text2.Text
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Dim RS2 As New ADODB.Recordset
   
   strSQL = "SELECT * " & _
            "FROM Read_Version"
   RS2.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
   frmReadCodes.Caption = "Version: " & RS2!Version & " - Amended: " & Format(RS2!LastAmended, "dd/mm/yyyy")
   RS2.Close
   Set RS2 = Nothing
   optRC(1).value = True
   Initialize
End Sub

Private Sub Initialize()
   Text2.Text = assignedRC
   RS.MoveFirst
   RS.Find searchKey & " =  '" & Text2.Text & "'"
   If RS.EOF = False Then
      Do Until RS(searchKey) = Text2.Text Or RS.EOF
         RS.MoveNext
         RS.Find searchKey & " =  '" & Text2.Text & "'"
      Loop
   End If
   If RS.EOF = False Then
      Text1.Text = RS(returnKey)
      Text1_KeyUp 0, 0
   End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   pos = List1.ListIndex
   If pos >= 0 Then
      Text1.Text = List1.List(pos)
      rcIndex = List1.ItemData(pos)
      ShowReadCode searchKey
   End If
End Sub

Private Sub optRC_Click(Index As Integer)
   Dim i As Integer
   
   rcIndex = -1
   For i = 0 To List1.ListCount - 1
      If List1.Selected(i) = True Then
         rcIndex = List1.ItemData(i)
         Exit For
      End If
   Next i
   
   If rcIndex = -1 Then
      If List1.ListCount = 1 Then
         rcIndex = List1.ItemData(0)
      End If
   End If
   
   Select Case Index
      Case 0
         searchKey = ReadDb("Read_4BRubric")
         returnKey = "Read_4BRubric"
'         frmReadCodes.Caption = "4-Byte Read Codes" & frmCaption
      
      Case 1
         searchKey = ReadDb("Read_V2Rubric")
         returnKey = "Read_V2Rubric"
'         frmReadCodes.Caption = "Version 2 Read Codes" & frmCaption
      
      Case 2
         searchKey = ReadDb("Read_V3Rubric")
         returnKey = "Read_V3Rubric"
'         frmReadCodes.Caption = "Version 3 Read Codes" & frmCaption
   
   End Select
   List1.Visible = False
   List1.Clear
   For i = 0 To lstHide.ListCount - 1
      List1.AddItem lstHide.List(i)
      List1.ItemData(i) = lstHide.ItemData(i)
   Next i
   List1.Visible = True
      
'   List1.Selected(pos) = True
   ShowReadCode searchKey
   Text1_KeyUp 0, 0
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Dim i As Integer
   
   keyCount = Len(Text1.Text)
   
   If keyCount = 0 Then
      Text2.Text = ""
   End If 'Else
      List1.Clear
      For i = 0 To lstHide.ListCount
         If InStr(1, lstHide.List(i), Text1.Text, 1) > 0 Then
            List1.AddItem lstHide.List(i)
            List1.ItemData(List1.ListCount - 1) = lstHide.ItemData(i)
         End If
      Next i
   'End If
   If KeyCode = 8 Or KeyCode = 46 Then
      Text2.Text = ""
   Else
      If List1.ListCount > 0 Then
         List1.Selected(0) = True
         rcIndex = List1.ItemData(0)
         ShowReadCode searchKey
      End If
   End If
'      Else
'         Text2.Text = ""
'      End If
End Sub

Private Function ReadDb(OrderKey As String) As String
   
   Dim strSQL As String
   Dim i As Integer
   
   strSQL = "SELECT * " & _
            "FROM Read_Codes " & _
            "WHERE Read_Release_Date <= '" & Format(Now(), "yyyymmdd") & "' " & _
            "ORDER BY " & OrderKey
   On Local Error Resume Next
   RS.Close
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   lstHide.Clear
   Do Until RS.EOF
      Select Case OrderKey
         Case "Read_4BRubric"
            lstHide.AddItem RS!Read_4BRubric
         
         Case "Read_V2Rubric"
            lstHide.AddItem RS!Read_V2Rubric
         
         Case "Read_V3Rubric"
            lstHide.AddItem RS!Read_V3Rubric
         
      End Select
      lstHide.ItemData(i) = RS!Read_Index
      i = i + 1
      RS.MoveNext
   Loop
   Select Case OrderKey
      Case "Read_4BRubric"
         ReadDb = "Read_4BRC"
      
      Case "Read_V2Rubric"
         ReadDb = "Read_V2RC"
         
      Case "Read_V3Rubric"
         ReadDb = "Read_V3RC"
   End Select
   
End Function

Private Sub ShowReadCode(ReturnData As String)
   If rcIndex > -1 Then
      RS.MoveFirst
      RS.Find "Read_Index = " & rcIndex
      Text2.Text = RS(searchKey)
      
      If RS!Read_Battery = "T" Then
         lblDetails(0).Caption = "Battery Header"
      ElseIf RS!Read_Test = "T" Then
         lblDetails(0).Caption = "Test Only"
      Else
         lblDetails(0).Caption = "No restriction"
      End If
      
      If RS!Read_Status = "D" Then
         lblDetails(1).Caption = "Status: Deleted"
         lblDetails(1).ForeColor = BPRED
         txtComment.Text = RS!Read_Comments
         txtComment.Visible = True
      Else
         lblDetails(1).Caption = "Status: Current"
         lblDetails(1).ForeColor = BPGREEN
         txtComment.Text = Trim(RS!Read_Comments & "")
         
         If txtComment.Text = "" Then
            txtComment.Visible = False
         Else
            txtComment.Visible = True
         End If
      End If
         
      If Trim(RS!Read_Release_Date & "") = "" Then
         lblDetails(2).Caption = "Released: Not Specified"
      Else
         lblDetails(2).Caption = "Released: " & RS!Read_Release_Date
      End If
      
      If RS!Read_Ratio = "T" Then
         lblDetails(3).Caption = "Read Ratio"
         lblDetails(3).ForeColor = BPBLUE
      Else
         lblDetails(3).Caption = ""
      End If
      
'      Text1.Text = RS(returnKey)
   Else
      Text2.Text = ""
   End If
End Sub

Public Property Let CurrentReadCode(ByVal strNewValue As String)
   assignedRC = strNewValue
End Property
