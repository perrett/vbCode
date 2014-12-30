VERSION 5.00
Begin VB.Form frmNewData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   1185
      TabIndex        =   3
      Top             =   300
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2385
      TabIndex        =   2
      Top             =   300
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000A&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   855
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3540
   End
End
Attribute VB_Name = "frmNewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private ButtonId As String
Private pListId As String
Private panelName As String
Private PanelId As Long

Private Sub cmdCancel_Click()
   ButtonId = "Cancel"
   Unload Me
End Sub

Private Sub cmdDelete_Click()
'   Dim RS As New ADODB.Recordset
'   Dim totEntries As Integer
'   Dim eList As Long
'
'   strSQL = "SELECT COUNT(PanelID) as TotEntries " & _
'            "FROM "
'   If pListId = "PANEL_PAGES" Then
'      strSQL = strSQL & "Request_Panels_Pages "
'   ElseIf pListId = "SCREEN_PANEL" Then
'      strSQL = strSQL & "Request_Panels "
'   End If
'
'   strSQL = strSQL & "WHERE PanelID = " & PanelId
'
'   If pListId = "PANEL_PAGES" Then
'      strSQL = strSQL & _
'               "AND PageName = '" & txtValue.Text & "'"
'   End If
'
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'   totEntries = RS!totEntries
'   RS.Close
'
'   If pListId = "SCREEN_PANEL" Then
'      If MsgBox("This will delete ALL the pages associated with this panel " & _
'               "(Affecting " & totEntries & " tests). Are you sure?", vbYesNo, "IceConfig Deletion Warning") = vbYes Then
'         strSQL = "DELETE FROM Request_Panels_Pages " & _
'                  "WHERE PanelID = " & PanelId
'         ICECon.Execute strSQL
''         frmMain.ediPr(pListId).ListItems.Remove panelID
'         strSQL = "DELETE FROM Request_Panels" & _
'                  " WHERE PanelID = " & PanelId
'         ICECon.Execute strSQL
'      End If
'
'   Else
'      If MsgBox("There are " & totEntries & " tests of this page. Are you sure you want to delete this page?", vbYesNo, "IceConfig Deletion Warning") = vbYes Then
'         strSQL = "DELETE FROM Request_Panels_Pages " & _
'                  "WHERE PanelId = '" & PanelId & "' " & _
'                     "AND Page_Name = '" & txtValue.Text & "'"
'         ICECon.Execute strSQL
'      End If
'   End If
'   With frmMain.ediPr(pListId)
'      eList = .ListItems.NameToIndex(txtValue.Text)
'      .ListItems.Remove eList
'      .value = ""
'   End With
   ButtonId = "Delete"
   Me.Hide
End Sub

Private Sub cmdSave_Click()
   If txtValue.Text = "" Then
      MsgBox "You cannot leave this field blank. An identifier is mandatory", vbExclamation, "Validation"
   Else
      ButtonId = "Save"
      Me.Hide
   End If
End Sub

Public Property Get ExitMethod() As String
   ExitMethod = ButtonId
End Property

Public Property Let PanelDetails(strNewValue As String)
   Dim RS As New ADODB.Recordset
   
   panelName = strNewValue
   strSQL = "SELECT * " & _
            "FROM Request_Panels " & _
            "WHERE PanelName = '" & panelName & "'"
   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
   If RS.RecordCount > 0 Then
      PanelId = RS!PanelId
   End If
   RS.Close
   Set RS = Nothing
End Property

Public Property Let PropertylistItem(strNewValue As String)
   pListId = strNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
   If transCount > 0 Then
      ICECon.CommitTrans
      transCount = transCount - 1
   End If
End Sub

Private Sub txtValue_GotFocus()
   txtValue.SelStart = 0
   txtValue.SelLength = Len(txtValue.Text)
End Sub
