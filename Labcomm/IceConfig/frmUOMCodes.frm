VERSION 5.00
Begin VB.Form frmUOMCodes 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstSpec 
      Height          =   2010
      Left            =   1545
      TabIndex        =   3
      Top             =   360
      Width           =   2715
   End
   Begin VB.TextBox txtSpec 
      Height          =   345
      Left            =   210
      TabIndex        =   2
      Top             =   1260
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   435
      TabIndex        =   1
      Top             =   2760
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2370
      TabIndex        =   0
      Top             =   2760
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "UOM Text"
      Height          =   285
      Left            =   300
      TabIndex        =   5
      Top             =   780
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "UOM List"
      Height          =   255
      Left            =   2220
      TabIndex        =   4
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "frmUOMCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RS As New ADODB.Recordset
Private keyCount As Integer
Private bText As String
Private pos As Integer

Private Sub Form_Load()

   Dim strSQL

   frmUOMCodes.Caption = "Select the National UOM Code"
   strSQL = "SELECT * FROM EDI_UOM_Nat_Codes ORDER BY EDIUOM_NatCode"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      lstSpec.AddItem RS!EDIUOM_NatCode
      RS.MoveNext
   Loop
   
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   RS.Close
   Set RS = Nothing
End Sub

Private Sub CmdOk_Click()
   frmMain.edipr(frmMain.edipr.Tag).value = lstSpec.List(lstSpec.ListIndex)
   RS.Close
   Set RS = Nothing
   Unload Me
End Sub

Private Sub lstSpec_DblClick()
   CmdOk_Click
End Sub

Private Sub txtSpec_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Dim i As Integer
   
   keyCount = Len(txtSpec.Text)
   If keyCount = 0 Then
      frmUOMCodes.Caption = "<No UOM Code Selected>"
      lstSpec.Selected(pos) = False
   Else
      For i = 0 To lstSpec.ListCount
         If UCase(Left(lstSpec.List(i), keyCount)) = UCase(txtSpec.Text) Then
            lstSpec.Selected(i) = True
            pos = i
            Exit For
         End If
      Next i
   End If
   
End Sub
