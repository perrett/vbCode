VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Begin VB.Form frmColour 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colour Selector"
   ClientHeight    =   4290
   ClientLeft      =   3180
   ClientTop       =   4605
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   255
      Left            =   120
      TabIndex        =   2
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
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double-click the colour you wish to select."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colIndex As Long
Private strSQL As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmColourEditor.Show 1
    LoadColours
End Sub

Private Sub Command3_Click()
    If vbalGrid1.SelectedCol < 1 Then Exit Sub
    Load frmColourEditor
    frmColourEditor.Label3.Caption = vbalGrid1.Cell(vbalGrid1.SelectedRow, 3).Text
    frmColourEditor.Text1.Text = vbalGrid1.Cell(vbalGrid1.SelectedRow, 2).Text
    frmColourEditor.Picture1.BackColor = vbalGrid1.Cell(vbalGrid1.SelectedRow, 1).BackColor
    frmColourEditor.Show 1
    LoadColours
End Sub

Private Sub Form_Load()
    LoadColours
End Sub

Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
   If lRow > 0 Then
      Debug.Print "Colour - " & PickedColIndex
      Debug.Print "Colour - " & PickedColName
      PickedCol = vbalGrid1.CellBackColor(lRow, 1)
      PickedColIndex = Val(vbalGrid1.Cell(lRow, 3).Text)
      PickedColName = vbalGrid1.Cell(lRow, 2).Text
   End If
   Unload Me
End Sub

Private Sub LoadColours()
   Dim RS As New ADODB.Recordset
   Dim Col As Long
   Dim showCol As Long
   
   vbalGrid1.Clear True
   strSQL = "SELECT * " & _
            "FROM Colours " & _
            "ORDER BY Colour_Name"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      With vbalGrid1
         .AddColumn "Col0", "Col0", ecgHdrTextALignCentre, , 25
         .AddColumn "Col1", "Col1", ecgHdrTextALignCentre, , 220
         .AddColumn "Col2", "Col2", ecgHdrTextALignCentre, , 0, False
         .GridLines = True
         .rows = RS.RecordCount
         .ScrollBarStyle = ecgSbrEncarta
         '.Columns = 1
         nocolours = 0
         Do While Not RS.EOF
            Col = TranslateColor(Val(RS!Colour_Code))
            .CellDetails nocolours + 1, 2, Trim(RS!Colour_Name) & "", DT_CENTER
            .CellDetails nocolours + 1, 3, RS!Colour_Index
            .CellDetails nocolours + 1, 1, "", DT_CENTER, , Col
            If RS!Colour_Index = colIndex Then
               showCol = nocolours + 1
            End If
            nocolours = nocolours + 1
            RS.MoveNext
         Loop
         If showCol > 0 Then
            .SelectedRow = showCol
         End If
      End With
   End If
   RS.Close
   Set RS = Nothing
End Sub

Public Property Let CurrentColour(lngNewValue As Long)
   colIndex = lngNewValue
End Property
