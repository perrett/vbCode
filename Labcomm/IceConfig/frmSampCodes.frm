VERSION 5.00
Begin VB.Form frmSampCodes 
   Caption         =   "<No sample selected>"
   ClientHeight    =   3735
   ClientLeft      =   6225
   ClientTop       =   3975
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4320
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3120
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   1515
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   945
   End
   Begin VB.TextBox txtSpec 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2715
   End
   Begin VB.ListBox lstSpec 
      BackColor       =   &H00C0E0FF&
      Height          =   1620
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Label Label3 
      Caption         =   "Sample List"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Sample Text"
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Sample Code"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmSampCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RS As New ADODB.Recordset
Private keyCount As Integer
Private bText As String
Private pos As Integer
Private tableId As String
Private natCodeField As String
Private natDescField As String
Private retIndex As String
Private iVal As String
Private mCtrl As ManageControls

Public Property Let DbTable(strNewValue As String)
   tableId = strNewValue
End Property

Public Property Let NationalCodeField(strNewValue As String)
   natCodeField = strNewValue
End Property

Public Property Let NationalDescriptionField(strNewValue As String)
   natDescField = strNewValue
End Property

Public Property Let ReturnDataTo(strNewValue As String)
   retIndex = strNewValue
End Property

Public Property Let InitialValue(strNewValue As String)
   iVal = strNewValue
End Property

Private Sub Form_Load()

   Dim strSQL As String
   Dim hLite As Integer
   Dim i As Integer
   
   strSQL = "SELECT * FROM " & tableId & " ORDER BY " & natDescField
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      lstSpec.AddItem RS(1)   '  Description Field
      If Trim(RS(0)) = iVal Then
         txtSpec.Text = RS(1)
         hLite = i
      End If
      RS.MoveNext
      i = i + 1
   Loop
   txtCode.Text = iVal
   lstSpec.Selected(hLite) = True
'   ShowSpecCode
End Sub

Private Sub cmdCancel_Click()
   RS.Close
   Set RS = Nothing
   Unload Me
'   frmSampCodes.Visible = False
End Sub

Private Sub CmdOk_Click()
   Dim NodeId As Node
   Dim i As Integer
   Dim locTable As String
   Dim nType As String
   
   If retIndex <> "ST" Then
      Set mCtrl = loadCtrl.KeyControl
   End If
   txtSpec.Text = lstSpec.List(lstSpec.ListIndex)
   frmMain.edipr(retIndex).value = txtCode.Text
'   frmMain.edipr("NATDESC").value = txtSpec.Text
'   If retIndex = "ST" Then
'      frmMain.ediPr("SD").value = txtSpec.Text
'   End If
'   frmSampleData.txtNatCode.Text = txtCode.Text
   RS.Close
   Set RS = Nothing
   
   Set NodeId = Nothing
   
'   For i = 1 To nodeId.Children
'      frmMain.TreeView1.Nodes.Remove nodeId.Child.Index
'   Next i

   Select Case tableId
      Case "CRIR_Sample_Type"
         locTable = "EDI_Local_Sample_Types"
         nType = "SAMP"
         
      Case "CRIR_Sample_AnatOrigin"
         locTable = "EDI_Local_Sample_AnatOrigin"
         nType = "ANAT"
         
      Case "CRIR_Sample_CollectionType"
         locTable = "EDI_Local_Sample_CollectionTypes"
         nType = "COLL"
         
   End Select
'   frmMain.TreeView1.Nodes.Add nodeId, _
'                           tvwChild, _
'                           mCtrl.NewNodeKey(txtCode.Text, _
'                                            "SAMP", _
'                                            "SubHeader", _
'                                            , _
'                                            ms_DELETE), _
'                           txtCode.Text & " - " & txtSpec.Text, _
'                           frmMain.edipr("ICON").Icon, _
'                           frmMain.edipr("ICON").Icon
   
   If retIndex <> "ST" Then
      loadCtrl.EvNoEdit = True
'      loadCtrl.EvaluateRecords NodeId, locTable, txtCode.Text
   End If
   Set mCtrl = Nothing
   Unload Me
End Sub

Private Sub lstSpec_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   pos = lstSpec.ListIndex
   txtSpec.Text = lstSpec.List(pos)
   ShowSpecCode

End Sub

Private Sub txtSpec_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Dim i As Integer
   
   keyCount = Len(txtSpec.Text)
   If keyCount = 0 Then
      frmSampCodes.Caption = "<No Sample Code Selected>"
      lstSpec.Selected(pos) = False
      txtCode.Text = ""
   Else
      For i = 0 To lstSpec.ListCount
         If UCase(Left(lstSpec.List(i), keyCount)) = UCase(txtSpec.Text) Then
            lstSpec.Selected(i) = True
            pos = i
            ShowSpecCode
            Exit For
         End If
      Next i
   End If
   
End Sub

Private Sub ShowSpecCode()
   
   If pos > -1 Then
      frmSampCodes.Caption = lstSpec.List(pos)
      RS.MoveFirst
      RS.Find natDescField & " = '" & lstSpec.List(pos) & "'"
      txtCode.Text = RS(0) '  Code
   Else
      txtCode.Text = ""
   End If
End Sub

