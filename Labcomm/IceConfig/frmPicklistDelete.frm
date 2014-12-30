VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalgrid6.ocx"
Begin VB.Form frmPicklistDelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Picklist Value"
   ClientHeight    =   4635
   ClientLeft      =   5325
   ClientTop       =   4200
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6165
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
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double-click on the picklist value you wish to delete.  When you are finished click OK"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmPicklistDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub vbalGrid1_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If lRow < 1 Then Exit Sub
    ICECon.Execute "Delete From Request_Picklist_Data Where Picklist_Index=" & Label2.Caption & " And Picklist_Value='" & vbalGrid1.Cell(lRow, lCol).Text & "'"
    vbalGrid1.RemoveRow lRow
End Sub
