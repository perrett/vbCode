VERSION 5.00
Begin VB.Form frmGP 
   Caption         =   "National GP Codes"
   ClientHeight    =   6075
   ClientLeft      =   5640
   ClientTop       =   3270
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGP 
      Caption         =   "GP Details"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   5295
      Begin VB.TextBox txtGPName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtGPNatCode 
         Height          =   285
         Left            =   3120
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "National Code"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Practice"
      Height          =   3375
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtPrNatCode 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "National Code"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtPrAdr 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmGP.frx":0000
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtPrName 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   645
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmGP.frx":0011
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "National Code"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GP Selection"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
      Begin VB.OptionButton optGPSel 
         Caption         =   "Practice only"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optGPSel 
         Caption         =   "All GP's"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optGPSel 
         Caption         =   "Locally Defined"
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optGPSel 
         Caption         =   "System Defined"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ListBox lstGP 
      BackColor       =   &H00C0E0FF&
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      Caption         =   "Processing request: Please wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   600
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prId As String
Dim gpId As String
Dim IceNum As Long
Dim blnFirstPass As Boolean

Public Property Let PracticeId(strNEwValue As String)
   prId = strNEwValue
End Property

Public Property Let GPNatCode(strNEwValue As String)
   gpId = strNEwValue
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub CmdOk_Click()
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   
   If UCase(Left(txtGPNatCode.Text, 3)) = "ICE" Then
      strSQL = "{CALL ICEIMP_ICE_Nums}"
      RS.Open strSQL, ICECon, adOpenKeyset, adLockOptimistic
      RS("ICE_Clinician_Num") = IceNum + 1
      RS.Update
      RS.Close
   End If
   Set RS = Nothing
   frmMain.ediPr("IN+IN4").value = txtGPName.Text
   frmMain.ediPr("IN+IN5").value = txtGPNatCode.Text
   Unload Me
End Sub

Public Sub Form_Init()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   strSQL = "{CALL ICEIMP_ICE_Nums}"
   RS.Open strSQL, ICECon, adOpenKeyset, adLockPessimistic
   If RS.BOF And RS.EOF Then
      RS.Close
      strSQL = "{CALL ICEIMP_Ins_ICE_Temp_Nums (1,0,0,0,0)}"
      ICECon.Execute strSQL
      strSQL = "{CALL ICEIMP_ICE_Nums}"
      RS.Open strSQL, ICECon, adOpenKeyset, adLockPessimistic
   End If
   IceNum = RS("ICE_Clinician_Num")
   RS.Close
   If IceNum = 0 Then
      IceNum = 1
   End If
   optGPSel(0).Caption = "For practice " & prId
   
   strSQL = "SELECT * " & _
            "FROM National_GPs " & _
            "WHERE Clinician_National_Code = '" & gpId & "' " & _
               "AND Clinician_Name = '" & txtGPName.Text & "'"
   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount = 0 Then
      If UCase(Left(gpId, 3)) = "ICE" Then
         optGPSel(2).value = True
      Else
         optGPSel(3).value = True
      End If
   Else
      If RS!Practice_Code = prId Then
         optGPSel(0).value = True
      Else
         optGPSel(1).value = True
      End If
   End If
   RS.Close
   Set RS = Nothing
End Sub
   
Private Sub Form_Load()
'   optGPSel(0).value = True
End Sub

Private Sub lstGP_Click()
   Dim pos As Integer
   pos = InStr(lstGP.List(lstGP.ListIndex), " - ") - 1
   txtGPNatCode.Text = Left(lstGP.List(lstGP.ListIndex), pos)
   txtGPName.Text = Mid(lstGP.List(lstGP.ListIndex), pos + 4)
End Sub

Private Sub optGPSel_Click(Index As Integer)
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim lIndex As Integer
   
   strSQL = "SELECT * " & _
            "FROM National_GPs "
   lstGP.Enabled = True
   txtGPName.BackColor = &HC0FFFF
   txtGPNatCode.BackColor = &HC0FFFF
   lstGP.BackColor = &HC0E0FF
   fraGP.Enabled = False
   Select Case Index
      Case 0   '  Practice only
         strSQL = "SELECT * " & _
                  "FROM National_GPs " & _
                  "WHERE Practice_Code = '" & prId & "' " & _
                  "ORDER BY Clinician_Name"
         RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
         i = 0
         lstGP.Clear
         Do Until RS.EOF
            lstGP.AddItem RS!Clinician_National_Code & " - " & RS!Clinician_Name
            RS.MoveNext
         Loop
         RS.Close
         
      Case 1   '  All GP's
         frmGP.MousePointer = vbHourglass
'         frmGP.Caption = "Please wait..."
         strSQL = "SELECT * " & _
                  "FROM National_GPs " & _
                  "ORDER BY Clinician_Name"
         lstGP.Visible = False
         lblWait.Visible = True
         DoEvents
         RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
         i = 0
         lstGP.Clear
         Do Until RS.EOF
            lstGP.AddItem RS!Clinician_National_Code & " - " & RS!Clinician_Name
            RS.MoveNext
         Loop
         RS.Close
         lblWait.Visible = False
         lstGP.Visible = True
'         frmGP.Caption = "National GP Codes"
         frmGP.MousePointer = vbNormal
      
      Case 2   '  ICE No required
         lstGP.Enabled = False
         txtGPName.BackColor = &HFFFFFF
         txtGPNatCode.Locked = True
'         txtGPNatCode.BackColor = &HFFFFFF
         lstGP.BackColor = &H8000000F
         If UCase(Left(gpId, 3)) = "ICE" Then
            txtGPNatCode.Text = gpId
         Else
            txtGPNatCode.Text = "ICE" & IceNum
         End If
         fraGP.Enabled = True
         
      Case 3   '  A manually entered code
         lstGP.Enabled = False
         txtGPName.BackColor = &HFFFFFF
         txtGPNatCode.Locked = False
         txtGPNatCode.BackColor = &HFFFFFF
         lstGP.BackColor = &H8000000F
         txtGPNatCode.Text = gpId
         fraGP.Enabled = True
         
   End Select
   Set RS = Nothing
   blnFirstPass = False
End Sub

'Private Sub txtGPName_KeyDown(KeyCode As Integer, Shift As Integer)
'   If Len(txtGPName.Text) > 30 Then
'      MsgBox "A maximum of 30 chars only", vbExclamation, "Field too long"
'      txtGPName.Text = Left(txtGPName.Text, 30)
'   End If
'End Sub

Private Sub txtGPName_Validate(Cancel As Boolean)
   If Len(txtGPName.Text) > 30 Then
      MsgBox "30 characters is the maximum permissable length for this field", vbExclamation, "Field too long"
      txtGPName.SelStart = 30
      txtGPName.SelLength = Len(txtGPName.Text) - 30
      Cancel = True
   Else
      Cancel = False
   End If
End Sub

Private Sub txtGPNatCode_Validate(Cancel As Boolean)
   If Len(txtGPNatCode.Text) > 10 Then
      MsgBox "10 characters is the maximum permissable length for this field", vbExclamation, "Field too long"
      txtGPNatCode.SelStart = 10
      txtGPNatCode.SelLength = Len(txtGPNatCode.Text) - 10
      Cancel = True
   Else
      Cancel = False
   End If

End Sub
