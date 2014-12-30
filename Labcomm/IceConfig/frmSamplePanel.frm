VERSION 5.00
Begin VB.Form frmSamplePanel 
   Caption         =   "Sample Panel Options"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox lstPOptions 
      Height          =   2010
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmSamplePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private PanelId As Long
Private OptionName As String
Private OptionId As Long

Public Property Let PanelIdentifier(lngNewValue As Long)
   PanelId = lngNewValue
End Property

Public Property Get PanelOption() As Long
   PanelOption = OptionId
End Property

Public Property Get PanelOptionName() As String
   PanelOptionName = OptionName
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub CmdOk_Click()
   Dim RS As New ADODB.Recordset
   Dim seq As Long
   
   OptionId = lstPOptions.ItemData(lstPOptions.ListIndex)
   OptionName = lstPOptions.List(lstPOptions.ListIndex)
   strSQL = "SELECT Max(Sequence) as Seq " & _
            "FROM Request_Sample_Panels_Options " & _
            "WHERE Sample_Panel_ID = " & PanelId
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   seq = Val(RS!seq & "") + 1
   RS.Close
   Set RS = Nothing
   strSQL = "INSERT INTO Request_Sample_Panels_Options (" & _
               "Sample_Panel_Id, Option_Id, Sequence)" & _
            "VALUES(" & _
               PanelId & ", " & _
               OptionId & ", " & _
               seq & ")"
   iceCon.Execute strSQL
   Me.Hide
End Sub

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   
   strSQL = "SELECT * " & _
            "FROM Request_Sample_Panel_Collection_Day_Options " & _
            "WHERE Option_Id NOT IN (SELECT Option_ID " & _
                                    "FROM Request_Sample_Panels_Options " & _
                                       "WHERE Sample_Panel_ID = " & PanelId & ")"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   Do Until RS.EOF
      lstPOptions.AddItem RS!Description
      lstPOptions.ItemData(lstPOptions.ListCount - 1) = RS!Option_Id
      RS.MoveNext
   Loop
   RS.Close
   Set RS = Nothing
End Sub
