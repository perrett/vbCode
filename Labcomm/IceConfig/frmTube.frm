VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Begin VB.Form frmTube 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tube Selector"
   ClientHeight    =   4290
   ClientLeft      =   1635
   ClientTop       =   1470
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   255
      Left            =   1620
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3060
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "Double-click the tube you wish to select."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmTube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    PickedTubeIndex = -1
    Unload Me
End Sub

Private Sub Command2_Click()
    frmTubeEditor.TubeIndex = ""
    frmTubeEditor.Show 1
    LoadTubes
End Sub

Private Sub Command3_Click()
    If vbalGrid1.SelectedCol < 1 Then Exit Sub
    frmTubeEditor.TubeIndex = vbalGrid1.Cell(vbalGrid1.SelectedRow, 3).Text
    Load frmTubeEditor
    'frmColourEditor.Text1.Text = vbalGrid1.Cell(vbalGrid1.SelectedRow, 1).Text
    'frmColourEditor.Picture1.BackColor = vbalGrid1.Cell(vbalGrid1.SelectedRow, 1).BackColor
    frmTubeEditor.Show 1
    LoadTubes
End Sub

Public Property Let CurrentTube(lngNewValue As Long)
   PickedTubeIndex = lngNewValue
End Property

Private Sub Form_Load()
    LoadTubes
End Sub

Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If lRow < 1 Then Exit Sub
    PickedTube = vbalGrid1.Cell(lRow, 2).Text
    PickedTubeIndex = Val(vbalGrid1.Cell(lRow, 3).Text)
    PickedTubeCol = vbalGrid1.Cell(lRow, 1).BackColor
    Unload Me
End Sub

Private Sub LoadTubes()
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim Col As Long
   Dim showTube As Long
   Dim noTubes As Integer
   
   vbalGrid1.Clear True
   
   RS.Open "Select * From Request_Tubes Order by Name", iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      With vbalGrid1
         .AddColumn "Col0", "Col0", ecgHdrTextALignLeft, , 25
         .AddColumn "Col1", "Col1", ecgHdrTextALignLeft, , 245
         .AddColumn "Col2", "Col2", ecgHdrTextALignLeft, , , False
         .rows = RS.RecordCount + 1
         .CellDetails 1, 1, "", , , vbWhite
         .CellDetails 1, 2, "No container required"
         .CellDetails 1, 3, 0
         .ScrollBarStyle = ecgSbrEncarta
         .GridLines = True
         '.Columns = 1
         noTubes = 1
         Do While Not RS.EOF
            If Format(RS!Colour) & "" <> "" Then
               RS2.Open "Select Colour_Code From Colours Where Colour_Index=" & Format(RS!Colour), iceCon, adOpenKeyset, adLockReadOnly
               If RS2.RecordCount = 1 Then
                  Col = TranslateColor(Val(RS2!Colour_Code))
               Else
                  Col = vbWhite
               End If
               RS2.Close
            Else
               Col = vbWhite
            End If
            .CellDetails noTubes + 1, 1, "", , , Col
            If Trim(RS!Description) & "" <> "" Then
               .CellDetails noTubes + 1, 2, Trim(RS!Name) & " - " & Trim(RS!Description)
            Else
               .CellDetails noTubes + 1, 2, Trim(RS!Name) & ""
            End If
            .CellDetails noTubes + 1, 3, RS!Tube_Index
            noTubes = noTubes + 1
            If RS!Tube_Index = PickedTubeIndex Then
               showTube = noTubes
            End If
            RS.MoveNext
         Loop
         If showTube > 0 Then
            .CellSelected(showTube, 2) = True
'            .SelectedRow = showTube
         Else
            .CellSelected(1, 2) = True
'            .SelectedRow = 1
         End If
      End With
   End If
   RS.Close
End Sub
