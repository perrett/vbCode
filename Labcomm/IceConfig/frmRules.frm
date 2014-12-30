VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRules 
   Caption         =   "Rules"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Double-Click the rule you wish to add"
      Top             =   120
      Width           =   2895
   End
   Begin MSComctlLib.ListView lvRules 
      Height          =   2775
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data Entry Rule"
         Object.Width           =   7937
      EndProperty
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private PanelId As Long

Public Property Let PanelIdentifier(lngNewValue As Long)
   PanelId = lngNewValue
End Property

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   Dim lKey As String
   
   strSQL = "SELECT * " & _
            "FROM Request_Prompt " & _
            "WHERE Prompt_Type = 'DEN' " & _
               "AND Prompt_Index Not In (SELECT Prompt_Index " & _
                                        "FROM Request_Sample_Prompts " & _
                                        "WHERE Sample_Panel_Id = " & PanelId & ")"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   With lvRules.ListItems
      Do Until RS.EOF
         lKey = "Key_" & RS!Prompt_Index
         .Add , lKey, RS!Prompt_Desc
         RS.MoveNext
      Loop
   End With

   RS.Close
   Set RS = Nothing
End Sub

Private Sub lvRules_DblClick()
   Dim RS As New ADODB.Recordset
   Dim pSeq As Long
   
   strSQL = "SELECT Max(Sequence) " & _
            "FROM Request_Sample_Prompts " & _
            "WHERE Sample_Panel_Id = " & PanelId
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   pSeq = Val(RS(0) & "") + 1
   RS.Close
   
   strSQL = "INSERT INTO Request_Sample_Prompts (" & _
               "Sample_Panel_Id, " & _
               "Prompt_Index, " & _
               "Sequence) " & _
            "VALUES(" & _
               PanelId & ", " & _
               Mid(lvRules.SelectedItem.Key, 5) & ", " & _
               pSeq & ")"
   iceCon.Execute strSQL
   Me.Hide
End Sub
