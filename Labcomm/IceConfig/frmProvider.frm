VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Begin VB.Form frmProvider 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provider Selector"
   ClientHeight    =   4290
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5953
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
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double-click the provider you wish to select."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blnUnload As Boolean

Public Property Let CloseMe(blnNewValue As Boolean)
   blnUnload = blnNewValue
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmProviderEditor.Show 1
    LoadProviders
End Sub

Private Sub Command3_Click()
    If vbalGrid1.SelectedCol < 1 Then Exit Sub
    Load frmProviderEditor
    frmProviderEditor.Label1.Caption = vbalGrid1.Cell(vbalGrid1.SelectedRow, 2).Text
    frmProviderEditor.Show 1
    LoadProviders
End Sub

Private Sub Form_Activate()
   If blnUnload Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
    blnUnload = False
    LoadProviders
End Sub

Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If lRow < 1 Then Exit Sub
    PickedProv = vbalGrid1.Cell(lRow, 1).Text
    PickedProvIndex = Val(vbalGrid1.Cell(lRow, 2).Text)
    Unload Me
End Sub

Private Sub LoadProviders()
    Dim RS As ADODB.Recordset
    Dim Col As Long
    vbalGrid1.Clear True
    Set RS = New ADODB.Recordset
    RS.Open "Select * From Service_Providers Order by Provider_Name", iceCon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        With vbalGrid1
            .AddColumn "Col1", "Col1", ecgHdrTextALignCentre, , 249
            .AddColumn "Col2", "Col2", ecgHdrTextALignCentre, , , False
            .rows = RS.RecordCount
            .ScrollBarStyle = ecgSbrEncarta
            '.Columns = 1
            NoProviders = 0
            Do While Not RS.EOF
                .CellDetails NoProviders + 1, 1, Trim(RS!Provider_Name) & ""
                .CellDetails NoProviders + 1, 2, RS!Provider_ID
                NoProviders = NoProviders + 1
                RS.MoveNext
            Loop
        End With
    End If
    RS.Close
End Sub
