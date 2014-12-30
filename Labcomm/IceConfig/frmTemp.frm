VERSION 5.00
Object = "{3F118CA4-97A8-4926-AA0A-6FD80DFB3DCE}#2.6#0"; "PrpList2.ocx"
Begin VB.Form frmSpecs 
   Caption         =   "Specialty"
   ClientHeight    =   4680
   ClientLeft      =   9060
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3015
      TabIndex        =   2
      Top             =   4110
      Width           =   1275
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   345
      Left            =   750
      TabIndex        =   1
      Top             =   4110
      Width           =   1305
   End
   Begin PropertiesListCtl.PropertiesList ediPrSub 
      Height          =   3765
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6641
      LicenceData     =   "351606201721010B00172539012755210A0A1D221D3F5520001607391773250717160C2406"
      Caption         =   "Details"
      ShowPageStrip   =   -1  'True
      DescriptionHeight=   35
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   
   Me.Visible = False
   Unload frmSpecs
   
End Sub

Private Sub CmdOk_Click()

   Dim i As Integer
   
   For i = 1 To ediPrSub.PropertyItems.Count
      If ediPrSub(i).Selected Then
         ediPrSub_PropertyDblClick ediPrSub(i)
         Exit For
      End If
   Next i
   
End Sub

Private Sub ediPrSub_PropertyDblClick(PropertyItem As PropertiesListCtl.PropertyItem)

   Dim idx As Integer
   
   If strPropVal = "SP-MS3" Then
      frmmain.ediPr(strPropVal).Value = PropertyItem.Caption & "," & PropertyItem.Value
   ElseIf strPropVal = "SP-MS1" Then
      frmmain.ediPr(Left(strPropVal, 2)).Value = PropertyItem.Caption
      frmmain.ediPr(strPropVal).Value = PropertyItem
      idx = frmmain.ediPr.PropertyItems.KeyToIndex(strPropVal) + 1
      frmmain.ediPr(idx).Value = PropertyItem.Caption
   End If
   Me.Visible = False
   Unload frmSpecs
   
End Sub

Private Sub Form_Activate()

   If ediPrSub.Tag = "Korner" Then
      ediPrSub.ActivePage = ediPrSub("S" & frmmain.ediPr(strPropVal).Value).PageKeys
      ediPrSub.Pages(ediPrSub("S" & frmmain.ediPr(strPropVal).Value).PageKeys).Selected = True
      ediPrSub("S" & frmmain.ediPr(strPropVal).Value).Selected = True
      ediPrSub("S" & frmmain.ediPr(strPropVal).Value).EnsureVisible
   End If
   
End Sub


Private Sub Form_Load()
   
   Dim RS As New ADODB.Recordset
   Dim pageId As Integer
   Dim pos As Integer
   Dim dLen As Integer

   If strPropVal = "SP-MS3" Then
      RS.Open "SELECT DISTINCT EDI_Msg_Format FROM EDI_Msg_Types ORDER BY EDI_Msg_Format", ICECon, adOpenKeyset, adLockReadOnly
      ediPrSub.PropertyItems.Clear
      ediPrSub.Pages.Clear
      ediPrSub.UsePageKeys = False
      Do Until RS.EOF
         pos = InStr(RS("EDI_Msg_Format"), ",")
         dLen = Len(RS("EDI_Msg_Format"))
         ediPrSub.PropertyItems.Add RS("EDI_Msg_Format"), Left(RS("EDI_Msg_Format"), dLen - pos), plpsString, Mid(RS("EDI_Msg_Format"), pos + 1), , , , pldsNormal
         RS.MoveNext
      Loop
   ElseIf strPropVal = "SP-MS1" Then
      pageId = 0
      RS.Open "SELECT * FROM Specialty", ICECon, adOpenKeyset, adLockPessimistic
      ediPrSub.PropertyItems.Clear
      ediPrSub.Pages.Clear
      ediPrSub.UsePageKeys = True
      Do Until RS.EOF
         If Int(Val(RS("Specialty_Code")) / 100) > Int(pageId / 100) Then
            pageId = Int(Val(RS("Specialty_code")))
            frmSpecs.ediPrSub.Pages.Add "P" & pageId, pageId & "'s"
         End If
         frmSpecs.ediPrSub.PropertyItems.Add "S" & RS("Specialty_Code"), Trim(RS("Specialty")), plpsString, Trim(RS("Specialty_Code"))
         frmSpecs.ediPrSub("S" & RS("Specialty_Code")).PageKeys = "P" & pageId
         RS.MoveNext
      Loop
   End If
   
End Sub

Private Sub SetReturnData(ByVal prop As PropertiesListCtl.PropertyItem)

   Dim idx As Integer
   
   frmmain.ediPr(Left(strPropVal, 2)).Value = prop.Caption
   frmmain.ediPr(strPropVal).Value = prop
   idx = frmmain.ediPr.PropertyItems.KeyToIndex(strPropVal) + 1
   frmmain.ediPr(idx).Value = prop.Caption
   Me.Visible = False
   Unload frmSpecs
   
End Sub
