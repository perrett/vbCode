VERSION 5.00
Object = "{3F118CA4-97A8-4926-AA0A-6FD80DFB3DCE}#2.6#0"; "PrpList2.ocx"
Begin VB.Form frmSpecs 
   Caption         =   "Specialty"
   ClientHeight    =   4680
   ClientLeft      =   5775
   ClientTop       =   3795
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   5220
   Visible         =   0   'False
   Begin PropertiesListCtl.PropertiesList ediPrSub 
      Height          =   3492
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   4572
      _ExtentX        =   8070
      _ExtentY        =   6165
      LicenceData     =   "351606201721010B00172539012755210A0A1D221D3F553204110570212314100E011A"
      DescriptionHeight=   44
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3015
      TabIndex        =   0
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
End
Attribute VB_Name = "frmSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private callOrigin As String
Private writeBack As String
   
Public Property Let EDI_WriteBack(strNewValue As String)
   writeBack = strNewValue
End Property

Public Property Let CalledFrom(strNewValue As String)
   callOrigin = strNewValue
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub CmdOk_Click()
   Dim i As Integer
   
   For i = 1 To ediPrSub.PropertyItems.Count
      If ediPrSub(i).Selected Then
         ediprsub_PropertyDblClick ediPrSub(i)
         Exit For
      End If
   Next i
   callOrigin = ""
   Unload Me
End Sub

Private Sub ediprsub_PropertyDblClick(PropertyItem As PropertiesListCtl.PropertyItem)
   Dim idx As Integer
   Dim thisOrigin As String
   

   Select Case callOrigin
      Case "SP+MS3"
         frmMain.ediPr(callOrigin).value = PropertyItem.Caption & "," & PropertyItem.value
      
      Case "SP+MS1"
         frmMain.ediPr(callOrigin).value = PropertyItem
         idx = frmMain.ediPr.PropertyItems.KeyToIndex(callOrigin) + 1
         frmMain.ediPr(idx).value = PropertyItem.Caption
      
      Case "NATCODE"
         frmMain.ediPr("NATDESC").value = PropertyItem.Caption
         frmMain.ediPr("DMSPEC").value = PropertyItem.value
      
      Case "TESTSPEC"
         If writeBack = "" Then
            frmMain.ediPr(callOrigin).value = PropertyItem.value
         Else
            frmMain.ediPr(writeBack).value = PropertyItem.value
         End If
      
      Case Else
         frmMain.ediPr(callOrigin).value = PropertyItem.value
      
   End Select
End Sub

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   Dim PageId As Integer
   Dim pos As Integer
   Dim dLen As Integer
   Dim strSQL As String
   Dim pract As String
   Dim AddSpecs As String
   Dim ExtraSpecs As String

   ExtraSpecs = ""
   
   AddSpecs = Read_Ini_Var("General", "ExtraSpecialties", iniFile)
   
   Dim sp() As String
   Dim i As Integer
   
   sp = Split(AddSpecs, ",")
   For i = 0 To UBound(sp)
      If IsNumeric(sp(i)) Then
         If ExtraSpecs = "" Then
            ExtraSpecs = sp(i)
         Else
            ExtraSpecs = ExtraSpecs & "," & sp(i)
         End If
      End If
   Next i
   
   If ExtraSpecs = "" Then
      ExtraSpecs = "502"
   End If
   
   If callOrigin = "" Then
      callOrigin = frmMain.ediPr.Tag
   End If
   
   pract = objTV.NodeLevel(objTV.TopLevelNode)
   
   Select Case callOrigin
      Case "SP+MS3"
         frmSpecs.Caption = "Message Types"
         strSQL = "SELECT EDI_Msg_Format FROM EDI_Msg_Types " & _
                  "WHERE EDI_Org_NatCode = '" & pract & "' " & _
                  "ORDER BY EDI_Msg_Format"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         ediPrSub.PropertyItems.Clear
         ediPrSub.Pages.Clear
         Do Until RS.EOF
            pos = InStr(RS!EDI_Msg_Format, ",")
            dLen = Len(RS!EDI_Msg_Format)
            ediPrSub.PropertyItems.Add RS("EDI_Msg_Format"), Left(RS("EDI_Msg_Format"), pos - 1), plpsString, Mid(RS("EDI_Msg_Format"), pos + 1), , , , pldsNormal
            ediPrSub(RS("EDI_Msg_Format")).ReadOnly = True
            RS.MoveNext
         Loop
            
      Case "SP+MS1"
         frmSpecs.Caption = "Specialties"
         PageId = 0
         strSQL = "SELECT * " & _
                  "FROM CRIR_Specialty " & _
                  "WHERE Specialty_code between 800 and 899 or Specialty_code in (" & ExtraSpecs & ")" & _
                  "ORDER BY Specialty_Code"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockPessimistic
         ediPrSub.PropertyItems.Clear
         ediPrSub.Pages.Clear
         
         frmSpecs.ediPrSub.PropertyItems.Add "S002", "Clinical Letter", plpsString, "002"
         
         Do Until RS.EOF
            frmSpecs.ediPrSub.PropertyItems.Add "S" & RS("Specialty_Code"), Trim(RS("Specialty")), plpsString, Trim(RS("Specialty_Code"))
            ediPrSub("S" & RS("Specialty_Code")).ReadOnly = True
            RS.MoveNext
         Loop
   
      Case "TESTSPEC"
         frmSpecs.Caption = "Specialties"
         PageId = 0
         strSQL = "SELECT * " & _
                  "FROM CRIR_Specialty " & _
                  "WHERE Specialty_code between '800' and '899' or Specialty_code = '502'" & _
                  "ORDER BY Specialty_Code"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockPessimistic
         ediPrSub.PropertyItems.Clear
         ediPrSub.Pages.Clear
         Do Until RS.EOF
            frmSpecs.ediPrSub.PropertyItems.Add "S" & RS("Specialty_Code"), Trim(RS("Specialty")), plpsString, Trim(RS("Specialty_Code"))
            ediPrSub("S" & RS("Specialty_Code")).ReadOnly = True
            RS.MoveNext
         Loop
      
      Case "NATCODE"
         PageId = 0
         strSQL = "SELECT * " & _
                  "FROM CRIR_Specialty " & _
                  "WHERE Specialty_code between '800' and '899' or Specialty_code = '502'" & _
                  "ORDER BY Specialty_Code"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockPessimistic
         ediPrSub.PropertyItems.Clear
         ediPrSub.Pages.Clear
         Do Until RS.EOF
            frmSpecs.ediPrSub.PropertyItems.Add "S" & RS("Specialty_Code"), Trim(RS("Specialty")), plpsString, Trim(RS("Specialty_Code"))
            ediPrSub("S" & RS("Specialty_Code")).ReadOnly = True
            RS.MoveNext
         Loop
      
      Case Else
         frmSpecs.Caption = "Active Recipients"
         ediPrSub.PropertyItems.Add "Blank", "(Clear Entry)", plpsString, "", "No Recipient"
         strSQL = "SELECT * " & _
                  "FROM EDI_Recipients " & _
                  "WHERE EDI_Active = 1"
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         Do Until RS.EOF
            ediPrSub.PropertyItems.Add "R" & RS!EDI_NatCode, RS!EDI_Name, plpsString, RS!EDI_NatCode, "Select the Practice required"""
            ediPrSub("R" & RS!EDI_NatCode).ReadOnly = True
            RS.MoveNext
         Loop
      
   End Select
End Sub
