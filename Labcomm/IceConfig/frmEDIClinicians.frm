VERSION 5.00
Begin VB.Form frmEDIClinicians 
   Caption         =   "Name & Local Code"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   390
      Left            =   2265
      TabIndex        =   15
      Top             =   6570
      Width           =   2745
   End
   Begin VB.Frame Frame2 
      Caption         =   "EDI Health Parties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   315
      TabIndex        =   9
      Top             =   810
      Width           =   6630
      Begin VB.CheckBox chkEHP 
         Alignment       =   1  'Right Justify
         Caption         =   "Amend to G888888"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   2010
         TabIndex        =   14
         Top             =   915
         Width           =   2430
      End
      Begin VB.TextBox txtEHPClin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1965
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   375
         Width           =   870
      End
      Begin VB.TextBox txtEHPEDI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   5415
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   375
         Width           =   870
      End
      Begin VB.Label lblEHPEDI 
         Caption         =   "EDI National code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3510
         TabIndex        =   11
         Top             =   420
         Width           =   1650
      End
      Begin VB.Label lblEHPClin 
         Caption         =   "Clin National Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   420
         Width           =   1725
      End
   End
   Begin VB.TextBox txtClinCode 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4170
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   240
      Width           =   1395
   End
   Begin VB.Frame fraEDI 
      Caption         =   "Practice Id: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   300
      TabIndex        =   1
      Top             =   2535
      Width           =   6675
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmEDIClinicians.frx":0000
         Top             =   285
         Width           =   6360
      End
      Begin VB.CommandButton cmdEDI 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Update Clinician"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2925
         Width           =   2220
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2295
         Width           =   1755
      End
      Begin VB.TextBox txtEDICode 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   5055
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2250
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "GP Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   2295
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "GP National Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3405
         TabIndex        =   4
         Top             =   2295
         Width           =   1695
      End
   End
   Begin VB.Label lblClinician 
      Caption         =   "Clinician National Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1155
      TabIndex        =   0
      Top             =   270
      Width           =   2715
   End
End
Attribute VB_Name = "frmEDIClinicians"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cNatCode As String
Private cName As String

Public Property Let ClinicianId(strNewValue As String)
   cNatCode = strNewValue
End Property

Public Property Let ClinicianName(strNewValue As String)
   cName = strNewValue
End Property

Private Sub PrepareForm()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   
   txtClinCode.Text = cNatCode
   
   strSQL = "SELECT Substring(Clinician_Local_code,7,10) As Local_Id, " & _
               "c1.Clinician_Surname, " & _
               "g1.Clinician_Name As EDI_Code_Name, " & _
               "g2.Clinician_Name As Clin_Code_Name, " & _
               "c1.Clinician_Speciality_Code, " & _
               "IsNull(ei2.EDI_NatCode,'') As EDI_NatCode, " & _
               "IsNull(ei2.EDI_Org_NatCode,'N/A') as EDI_Org_NatCode, " & _
               "IsNull(EDI_OP_Name, ei2.EDI_GP_Name) As GP_Name " & _
            "FROM Clinician c1 " & _
               "LEFT JOIN EDI_Recipient_Individuals ei1 " & _
               "ON c1.Clinician_National_Code = ei1.EDI_NatCode " & _
               "INNER JOIN Clinician_Local_Id cl " & _
                  "LEFT JOIN EDI_Matching em " & _
                     "INNER JOIN EDI_Recipient_Individuals ei2 " & _
                     "ON em.Individual_Index = ei2.Individual_Index " & _
                  "ON Substring(Clinician_Local_Code,7,10) = EDI_Local_Key3 " & _
               "ON c1.Clinician_National_Code = cl.Clinician_National_Code " & _
               "LEFT JOIN National_GPs g1 " & _
               "ON ei2.EDI_NatCode = g1.Clinician_National_Code " & _
               "LEFT JOIN National_GPs g2 " & _
               "ON c1.Clinician_National_Code = g2.Clinician_National_Code " & _
            "WHERE c1.Clinician_National_Code = '" & cNatCode & "'"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   frmEDIClinicians.Caption = RS!Clinician_Surname & " (" & RS!Local_Id & ")"
   fraEDI.Caption = RS!EDI_Org_NatCode
   txtEDICode.Text = IIf(RS!EDI_NatCode = "", frmMain.ediPr("NCODE").value, RS!EDI_NatCode)
   txtName.Text = IIf(IsNull(RS!GP_Name), frmMain.ediPr("SURNAME").value, RS!GP_Name)
   
'   lblSHPClin.Caption = cNatCode
'   lblSHPEDI.Caption = txtEDICode.Text
   lblEHPClin.Caption = cNatCode
   lblEHPEDI.Caption = txtEDICode.Text
   
   RS.Close
   
'   strSQL = "SELECT Count(*) " & _
'            "FROM Service_Health_Parties " & _
'            "WHERE Service_HP_Nat_Code = '" & cNatCode & "' " & _
'               "AND (Service_HP_Type = '902' " & _
'                  "OR Service_HP_Type = '906')"
'   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
'   txtSHPClin.Text = RS(0)
'   RS.Close
'
'   strSQL = "SELECT Count(*) " & _
'            "FROM Service_Health_Parties " & _
'            "WHERE Service_HP_Nat_Code = '" & txtEDICode.Text & "' " & _
'               "AND (Service_HP_Type = '902' " & _
'                  "OR Service_HP_Type = '906')"
'   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
'   txtSHPEDI.Text = RS(0)
'   RS.Close
   
   strSQL = "SELECT Count(*) " & _
            "FROM EDI_Health_Parties " & _
            "WHERE EDI_HP_Nat_Code = '" & cNatCode & "' " & _
               "AND (EDI_HP_Type = '902' " & _
                  "OR EDI_HP_Type = '906')"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   txtEHPClin.Text = RS(0)
   RS.Close
   
   strSQL = "SELECT Count(*) " & _
            "FROM EDI_Health_Parties " & _
            "WHERE EDI_HP_Nat_Code = '" & txtEDICode.Text & "' " & _
               "AND (EDI_HP_Type = '902' " & _
                  "OR EDI_HP_Type = '906')"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   txtEHPEDI.Text = RS(0)
   RS.Close
   
'   With chkSHP
'      .Caption = "Amend to " & lblSHPEDI.Caption
'      .Enabled = ((Val(txtSHPClin.Text) + Val(txtSHPEDI.Text)) > 0)
'      .value = Abs(CInt(chkSHP.Enabled))
'   End With
'
   With chkEHP
      .Caption = "Amend to " & lblEHPEDI.Caption
      .Enabled = ((Val(txtEHPClin.Text) + Val(txtEHPEDI.Text)) > 0)
      .value = Abs(CInt(chkEHP.Enabled))
   End With
   Set RS = Nothing
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdEDI_Click()
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim hpInd As Integer
   Dim blnContinue As Boolean
   Dim strPrompt As String
   
   blnContinue = True
   
   If Val(txtEHPClin.Text) + Val(txtEHPEDI.Text) = 0 Then
      hpInd = 0
   Else
'      If chkSHP.value = 1 Then
'         hpInd = 1
'      End If
'
'      If chkEHP.value = 1 Then
'         hpInd = (hpInd Or 2)
'      End If
   
      If chkEHP.value = 0 Then
         strPrompt = "You have chosen NOT to amend the Health party records. This may compromise your " & _
                     "database integrity."
      Else
         strPrompt = "Once these health party records have been amended, the process is irreversible."
      End If
      
      strPrompt = strPrompt & vbCrLf & vbCrLf & "Please call Sunquest on 0845 519 4020 if you are unsure."
      
      blnContinue = MsgBox(strPrompt, vbOKCancel, "Amend " & lblEHPClin.Caption & " to " & lblEHPEDI.Caption) = vbOK
   End If
   
   If blnContinue Then
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc
         .CommandText = "ICECONFIG_Amend_EDIClinician"
         .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
         .Parameters.Append .CreateParameter("NewNat", adVarChar, adParamInput, 8, txtEDICode.Text)
         .Parameters.Append .CreateParameter("OldNat", adVarChar, adParamInput, 8, cNatCode)
         .Parameters.Append .CreateParameter("Update", adBoolean, adParamInput, , (chkEHP.value = 1))
         .Parameters.Append .CreateParameter("Process", adVarChar, adParamOutput, 50)
         .Execute
         
         MsgBox .Parameters("Return").value & " received from proc. Process: " & .Parameters("Process").value, vbInformation, "Debug"
      End With
      
      txtClinCode.Text = txtEDICode.Text
      cNatCode = txtClinCode.Text
      Unload Me
   End If
   
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   
   If iceCmd.Parameters("Return").value = 0 Then
      Resume Next
   Else
      eClass.FurtherInfo = iceCmd.Parameters("Process").value
   End If
   eClass.CurrentProcedure = "frmEDIClinicians.cmdEDI_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub Form_Load()
   PrepareForm
End Sub
