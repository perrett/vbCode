VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICE...Configuration Login"
   ClientHeight    =   2415
   ClientLeft      =   3240
   ClientTop       =   4230
   ClientWidth     =   4335
   ControlBox      =   0   'False
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4335
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmLogon.frx":0442
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please enter your username and password into the boxes below in order to access this application."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnNoDBValidation As Boolean

Public Property Let ValidateInForm(blnNewValue As Boolean)
   blnNoDBValidation = blnNewValue
End Property

Private Function GetPassword(userID As String, passW As String) As String
'*******************************************************************
'  This allows for the encryption routine to be missing, and keeps
'  Config to a single version. If the relevant object is missing
'  The input password is returned
'*******************************************************************
   On Error GoTo procEH
   
   Dim cObj As Variant
   Dim pWord As String
   
   pWord = passW
   
   Set cObj = CreateObject("clsVB6CryptoUtility")
   
   pWord = cObj.EncryptString(userID, passW)
   
procEH:
   GetPassword = pWord

End Function

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2.SetFocus
    End If
End Sub

Private Sub Command1_Click()
   Dim RS As New ADODB.Recordset
   Dim strSQl As String
   
   If blnNoDBValidation = False Then
      '  NOTE: We try with both Encrypted and unencrypted password.
      strSQl = "SELECT User_FullName " & _
              "FROM Service_User " & _
              "WHERE User_Name = '" & Combo1.Text & "' " & _
                 "AND (User_Password='" & GetPassword(Combo1.Text, Text2.Text) & "' " & _
                 "OR User_Password = '" & Text2.Text & "')"
      RS.Open strSQl, iceCon, adOpenKeyset, adLockReadOnly
      
      If RS.RecordCount = 0 Then
         Text2.SelStart = 0
         Text2.SelLength = Len(Text2.Text)
         Text2.SetFocus
         MsgBox "Invalid username/password combination, please check and try again.", vbInformation + vbOKOnly, "Invalid Login"
         RS.Close
         Exit Sub
      End If
      
      userID = Combo1.Text
      RS.Close
      Unload Me
   Else
      Me.Hide
   End If
   
   Set RS = Nothing
End Sub

Private Sub Command2_Click()
   If blnNoDBValidation = False Then
      iceCon.Close
      Set iceCon = Nothing
      End
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
   If blnNoDBValidation Then
      Combo1.Text = "Admin"
      Combo1.Enabled = False
      Text2.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select User_Name from Service_User", iceCon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount <> 0 Then
        Do While Not RS.EOF
            Combo1.AddItem RS!User_Name
            RS.MoveNext
        Loop
    End If
    RS.Close
End Sub

Private Function DayCode() As String
    Dim DD As Long, MM As Long, YY As Long, DOW As Long
    Dim TempStr As String, Code As String
    Dim Temp As String
    Dim Present
    Present = Format(Now)
    DOW = Val(Format(Present, "W"))
    YY = Val(Format(Present, "YYYY"))
    MM = Val(Format(Present, "MM"))
    DD = Val(Format(Present, "DD"))
    Code = ""
    Code = Format(DD + DOW)
    Code = Code + Format((((YY - 1203) \ 2) - (MM * DOW)) + DD)
    Code = Code + Format(MM + DOW)
    Temp = Format(Present, "MM") + Format(DD * DOW)
    Code = Code + Mid$(Format(YY - Val(Temp)), 3, 2)
    DayCode = Mid$(Code, 1, 3) + "-" + Mid$(Code, 4, 3) + "-" + Mid$(Code, 7, Len(Code) - 6)
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Command1_Click
    End If
End Sub
