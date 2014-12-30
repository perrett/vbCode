VERSION 5.00
Begin VB.Form frmGP 
   Caption         =   "National GP Codes"
   ClientHeight    =   6075
   ClientLeft      =   5640
   ClientTop       =   3270
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGP 
      Caption         =   "GP Details"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   5295
      Begin VB.CommandButton cmdGPFind 
         Caption         =   "Find..."
         Height          =   225
         Left            =   915
         TabIndex        =   21
         Top             =   900
         Width           =   1125
      End
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
      Left            =   180
      TabIndex        =   3
      Top             =   2010
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
      Left            =   3495
      TabIndex        =   2
      Top             =   5325
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   615
      TabIndex        =   1
      Top             =   5325
      Width           =   1455
   End
   Begin VB.ListBox lstGP 
      BackColor       =   &H00C0E0FF&
      Height          =   1620
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   2535
   End
   Begin VB.Label lblWait 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter part of the GP's name  in GP Details/Name then click Find... to display matching GP's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1395
      Left            =   195
      TabIndex        =   17
      Top             =   210
      Visible         =   0   'False
      Width           =   2490
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
Dim gpRS As New ADODB.Recordset

Public Property Let PracticeId(strNewValue As String)
   prId = strNewValue
End Property

Public Property Let GPNatCode(strNewValue As String)
   gpId = strNewValue
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdGPFind_Click()
   Dim gpLen As Integer
   
'   optGPSel(1).value = True
   gpRS.Filter = ""
   gpLen = Len(txtGPName.Text)
   If gpLen > 0 Then
      gpRS.Filter = "Clinician_Name LIKE '%" & txtGPName.Text & "%'"
      lstGP.Visible = False
      lstGP.Clear
      Do Until gpRS.EOF
         lstGP.AddItem gpRS!Clinician_National_Code & " - " & gpRS!Clinician_Name
         gpRS.MoveNext
      Loop
      lstGP.Visible = True
   End If
End Sub

Private Sub CmdOk_Click()
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim blnAllow As Boolean
   
   blnAllow = False
   
   If txtGPNatCode.Text = "<System defined>" Then
      If txtGPNatCode.Enabled = False Then
         iceCon.BeginTrans
         strSQL = "{CALL ICEIMP_ICE_Nums}"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockOptimistic
         
         If RS.BOF And RS.EOF Then
'           No ice record present
            RS.Close
            strSQL = "{CALL ICEIMP_Ins_ICE_Temp_Nums (1,0,0,0,0)}"
            
            iceCon.Execute strSQL
            strSQL = "{CALL ICEIMP_ICE_Nums}"
            RS.Open strSQL, iceCon, adOpenKeyset, adLockPessimistic
         End If
         
         IceNum = RS("ICE_Clinician_Num") + 1
         RS("ICE_Clinician_Num") = IceNum
         RS.Update
         iceCon.CommitTrans
         txtGPNatCode.Text = "ICE" & IceNum
         RS.Close
         blnAllow = True
      End If
   
   ElseIf UCase(Left(txtGPNatCode.Text, 1)) <> "G" Then
      If MsgBox("This is not a valid GP National Code. Are you sure you wish to use " & _
                txtGPNatCode.Text & " as a National Identifier", vbExclamation + vbYesNo, _
                "Data Validation Warning") = vbYes Then
         blnAllow = True
      End If
   Else
      blnAllow = True
   End If
   
   If blnAllow Then
      If objTV.newNode Then
         strSQL = "SELECT IDENT_CURRENT('Edi_Recipient_Individuals') + 1 As NextId"
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         frmMain.edipr("SUBID").value = RS!NextId
         RS.Close
      End If
      frmMain.edipr("IN+IN1").value = Replace(txtGPName.Text, "'", "`")
      frmMain.edipr("IN+IN2").value = txtGPNatCode.Text
      Unload Me
   Else
      MsgBox "The GP National Code is in an invalid format - please re-enter." & vbCrLf & _
             "If you are unsure, please use a system defined (ICE) code", vbExclamation, "Invalid GP National Code"
   End If
   gpRS.Close
   
   Set RS = Nothing
   Set gpRS = Nothing
   Unload Me
End Sub

Public Sub Form_Init()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
'   strSQL = "{CALL ICEIMP_ICE_Nums}"
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockPessimistic
'   If RS.BOF And RS.EOF Then
'      RS.Close
'      strSQL = "{CALL ICEIMP_Ins_ICE_Temp_Nums (1,0,0,0,0)}"
'      ICECon.Execute strSQL
'      strSQL = "{CALL ICEIMP_ICE_Nums}"
'      RS.Open strSQL, ICECon, adOpenKeyset, adLockPessimistic
'   End If
'   IceNum = RS("ICE_Clinician_Num")
'   RS.Close
'   If IceNum = 0 Then
'      IceNum = 1
'   End If
'   IceNum = -1
   If gpId <> "" Then
      optGPSel(0).Caption = "For practice " & prId
      
      gpRS.Filter = ""
'      gpRS.MoveFirst
      gpRS.Filter = "Clinician_National_Code = '" & gpId & "' "
   
'   strSQL = "SELECT * " & _
            "FROM National_GPs " & _
            "WHERE Clinician_National_Code = '" & gpId & "' "
'               "AND Clinician_Name = '" & txtGPName.Text & "'"
'   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If gpRS.RecordCount = 0 Then
         If UCase(Left(gpId, 3)) = "ICE" Then
            IceNum = Val(Mid(gpId, 4))
            optGPSel(2).value = True
         Else
            optGPSel(3).value = True
         End If
      Else
         If gpRS!Practice_Code = prId Then
            optGPSel(0).value = True
         Else
            optGPSel(1).value = True
            txtGPName = gpRS!Clinician_Name
            txtGPNatCode.Text = gpId
            cmdGPFind_Click
         End If
      End If
   Else
      optGPSel(0).value = True
   End If
End Sub
   
Private Sub Form_Load()
   Dim strSQL As String
   Dim rsFile As String
   
   If gpRS.State = adStateClosed Then
      strSQL = "SELECT * " & _
               "FROM National_GPs "
      gpRS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
      rsFile = fs.BuildPath(App.Path, "natGPs.adtg")
      Set gpRS.ActiveConnection = Nothing
      
      If fs.FileExists(rsFile) Then
         fs.DeleteFile rsFile
      End If
      
      gpRS.Save rsFile, adPersistADTG
      gpRS.Close
      
      gpRS.Open rsFile
   End If
   
   optGPSel(0).value = True
End Sub

Private Sub lstGP_Click()
   Dim pos As Integer
   pos = InStr(lstGP.List(lstGP.ListIndex), " - ") - 1
   txtGPNatCode.Text = Left(lstGP.List(lstGP.ListIndex), pos)
   txtGPName.Text = Mid(lstGP.List(lstGP.ListIndex), pos + 4)
End Sub

Private Sub optGPSel_Click(Index As Integer)
   Dim i As Integer
   Dim lIndex As Integer
   
   lstGP.Enabled = True
   txtGPName.BackColor = &HC0FFFF
   txtGPNatCode.BackColor = &HC0FFFF
   lstGP.BackColor = &HC0E0FF
   fraGP.Enabled = False
   
   cmdGPFind.Enabled = False
   
   Select Case Index
      Case 0   '  Practice only
         lstGP.Visible = True
         lblWait.Visible = False
         
         gpRS.Filter = "Practice_Code = '" & prId & "' "
         lstGP.Clear
         Do Until gpRS.EOF
            lstGP.AddItem gpRS!Clinician_National_Code & " - " & gpRS!Clinician_Name
            gpRS.MoveNext
         Loop
         gpRS.Filter = ""
         If gpRS.RecordCount > 0 Then
            gpRS.MoveFirst
         End If
         
      Case 1   '  All GP's
         cmdGPFind.Enabled = True
         
         lblWait.Caption = "Enter all or part of the GP's name  in GP Details/Name field, then click Find... to display matching GP's"
         
         txtGPName.BackColor = &HFFFFFF
'         txtGPName.Text = ""
         
'         txtGPNatCode.Text = ""
         txtGPNatCode.Locked = True
         fraGP.Enabled = True
         
         lstGP.Clear
         lstGP.Visible = False
         lblWait.Visible = True
      
      Case 2   '  ICE No required
         
         lblWait.Caption = "Labcomm will generate an 'ICE' number for this GP. This is the usual option if the National code is unknown"
         lstGP.Visible = False
         lblWait.Visible = True
         
         txtGPNatCode.Enabled = False
         txtGPName.BackColor = &HFFFFFF
         txtGPNatCode.Locked = True
         
         
         If UCase(Left(gpId, 3)) = "ICE" Then
            txtGPNatCode.Text = gpId
         Else
            txtGPNatCode = "<System defined>"
         End If
         
         fraGP.Enabled = True
         
      Case 3   '  A manually entered code
         lstGP.Visible = False
         
         lblWait.Visible = True
         lblWait.Caption = "Manually enter a valid national code for the GP. Only use this option if you are sure of the national code."
         
         txtGPNatCode.Enabled = True
         txtGPName.BackColor = &HFFFFFF
         txtGPNatCode.Locked = False
         txtGPNatCode.BackColor = &HFFFFFF
         lstGP.BackColor = &H8000000F
         txtGPNatCode.Text = gpId
         fraGP.Enabled = True
         
   End Select
End Sub

Private Sub txtGPName_KeyUp(KeyCode As Integer, Shift As Integer)
'   Dim gpLen As Integer
'
'   If optGPSel(1).value = True Then
'      gpLen = Len(txtGPName.Text)
'      If gpLen > 0 Then
'         RS.Filter = "Left(Clinician_Name," & gpLen & ") = " & txtGPName.Text
'         lstGP.Visible = False
'         lstGP.Clear
'         Do Until RS.EOF
'            lstGP.AddItem
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
   If Len(txtGPNatCode.Text) > 8 Then
      MsgBox "8 characters is the maximum permissable length for this field", vbExclamation, "Field too long"
      txtGPNatCode.SelStart = 8
      txtGPNatCode.SelLength = Len(txtGPNatCode.Text) - 8
      Cancel = True
   Else
      Cancel = False
   End If
End Sub


