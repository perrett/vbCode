VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColourEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colour Editor"
   ClientHeight    =   1335
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   4410
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Colour Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Colour Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmColourEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error GoTo Col_Error
    CommonDialog1.Color = Picture1.BackColor
    CommonDialog1.ShowColor
    Picture1.BackColor = CommonDialog1.Color
    Exit Sub
Col_Error:
    On Error GoTo 0
End Sub

Private Sub Command3_Click()
    If Label3.Caption <> "" Then
        TempStr = "Update Colours Set Colour_Name='" & Text1.Text & "',Colour_Code='" & Picture1.BackColor & "' Where Colour_Index=" & Label3.Caption
        Debug.Print TempStr
        iceCon.Execute TempStr
    Else
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open "Select Max(Colour_Index) 'MaxCol' From Colours", iceCon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount = 1 And Format(RS!MaxCol & "") <> "" Then
            NewCol = RS!MaxCol + 1
        Else
            NewCol = 1
        End If
        RS.Close
        TempStr = "Insert Into Colours (Colour_Index,Colour_Name,Colour_Code,Date_Added) Values (" & Format(NewCol) & ",'" & Text1.Text & "','" & Picture1.BackColor & "','" & Format(Date, "DD MMM YYYY") & "')"
        Debug.Print TempStr
        iceCon.Execute TempStr
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
   Text1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        MsgBox "The character ' is not permitted, please use the ` character instead"
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1.Text = "" Then
      MsgBox "A name must be specified for this colour", vbExclamation, "No Colour Name"
      Cancel = True
   End If
End Sub
