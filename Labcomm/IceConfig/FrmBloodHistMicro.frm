VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmBloodHistMicro 
   Caption         =   "Bloodbank"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8145
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5953
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "BloodBank"
      TabPicture(0)   =   "FrmBloodHistMicro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAmend"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstProducts"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Histology"
      TabPicture(1)   =   "FrmBloodHistMicro.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Microbiology"
      TabPicture(2)   =   "FrmBloodHistMicro.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).ControlCount=   1
      Begin VB.ListBox lstProducts 
         Height          =   2205
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Frame fraAmend 
         Caption         =   "Amend BloodBank Products"
         Height          =   2445
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   5385
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0FFFF&
            Height          =   780
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "FrmBloodHistMicro.frx":0054
            Top             =   1320
            Width           =   4965
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   285
            Left            =   3795
            TabIndex        =   5
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Height          =   285
            Left            =   2130
            TabIndex        =   4
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   285
            Left            =   465
            TabIndex        =   3
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtProduct 
            Height          =   285
            Left            =   2160
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   330
            Width           =   2310
         End
         Begin VB.Label Label1 
            Caption         =   "Product"
            Height          =   240
            Left            =   795
            TabIndex        =   7
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Option not yet implemented"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   12
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Option not yet implemented"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   11
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Current Products"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmBloodHistMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String

Private Sub cmdAdd_Click()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim iMax As Integer
   
   strSQL = "SELECT MAX(Product_Name_Index) PNI " & _
            "FROM Request_Bloodbank_Product_Names"
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   iMax = RS("PNI") + 1
   RS.Close
   strSQL = "INSERT INTO Request_BloodBank_Product_Names " & _
               "(Product_Name_Index, Product_Name) " & _
            "VALUES (" & _
               iMax & ", '" & txtProduct.Text & "')"
   iceCon.Execute strSQL
   lstProducts.AddItem txtProduct.Text
   lstProducts.ItemData(lstProducts.ListCount - 1) = iMax
   txtProduct.Tag = iMax
   txtProduct.Text = ""
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmBloodHistMicro.cmdPAdd_Click"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   Dim RS As New ADODB.Recordset
   Dim mbText As String
   Dim pCount As Integer
   
   If MsgBox("Delete " & txtProduct.Text & " from BloodBank products?", vbYesNo, "Confirm Delete") = vbYes Then
      strSQL = "SELECT DISTINCT Request_BloodBank_Products.Test_Index, Request_Tests.Screen_Caption " & _
               "FROM Request_BloodBank_Products " & _
                  "INNER JOIN Request_Tests ON " & _
                  "Request_BloodBank_Products.Test_Index = Request_Tests.Test_Index " & _
               "WHERE Product_Name_Index = " & txtProduct.Tag
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.EOF Then
         strSQL = "DELETE FROM Request_BloodBank_Product_Names " & _
                  "WHERE Product_Name_index = " & txtProduct.Tag
         iceCon.Execute strSQL
         lstProducts.RemoveItem lbIndex(txtProduct.Tag, True)
         txtProduct.Text = ""
      Else
         mbText = "Unable to delete product. Product is used by:" & vbCrLf
         pCount = 0
         Do Until RS.EOF Or pCount > 10
            mbText = mbText & RS!Screen_Caption & vbCrLf
            RS.MoveNext
            pCount = pCount + 1
         Loop
         If RS.EOF = False Then
            mbText = mbText & "..." & vbCrLf
         End If
         mbText = mbText & "Remove product from these tests before deleting"
         MsgBox mbText, vbInformation, txtProduct.Text
      End If
      RS.Close
   End If
   Set RS = Nothing
End Sub

Private Sub cmdUpdate_Click()
   lstProducts.List(lbIndex(txtProduct.Tag, True)) = txtProduct.Text
   strSQL = "UPDATE Request_BloodBank_Product_Names " & _
            "SET Product_Name = '" & txtProduct.Text & "' " & _
            "WHERE Product_Name_Index = " & lstProducts.ItemData(lstProducts.ListIndex)
   iceCon.Execute strSQL
End Sub

Private Sub PopulateListBox()
   Dim intRow As Integer
   Dim RS As New ADODB.Recordset
   
   With lstProducts
      .Visible = False
      .Clear
   
      strSQL = "SELECT * " & _
               "FROM Request_BloodBank_Product_Names " & _
               "ORDER BY Product_Name_Index"
      eClass.FurtherInfo = strSQL
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      intRow = 0
      Do Until RS.EOF
         .AddItem RS!Product_Name
         .ItemData(intRow) = RS!Product_Name_Index
         RS.MoveNext
         intRow = intRow + 1
      Loop
      .Visible = True
   End With
   RS.Close
   Set RS = Nothing
End Sub

Public Function lbIndex(lbItem As String, _
                        Optional lbItemData As Boolean = False) As Integer
   Dim i As Integer
   
   With lstProducts
      For i = 0 To .ListCount - 1
         If lbItemData Then
            If .ItemData(i) = lbItem Then
               lbIndex = i
               Exit For
            End If
         Else
            If .List(i) = lbItem Then
               lbIndex = i
               Exit For
            End If
         End If
      Next i
   End With
End Function

Private Sub Form_Load()
   SSTab1.Tab = 0
   SSTab1_Click (0)
End Sub

Private Sub lstProducts_Click()
   txtProduct.Text = lstProducts.List(lstProducts.ListIndex)
   txtProduct.Tag = lstProducts.ItemData(lstProducts.ListIndex)
End Sub

Private Sub lstProducts_DblClick()
   txtProduct.Text = lstProducts.List(lstProducts.ListIndex)
   txtProduct.Tag = lstProducts.ItemData(lstProducts.ListIndex)
End Sub

Private Sub txtProduct_Validate(Cancel As Boolean)
   If Len(txtProduct.Text) > 20 Then
      MsgBox "20 characters is the maximum permissable length for this field", vbExclamation, "Field too long"
      txtProduct.SelStart = 20
      txtProduct.SelLength = Len(txtProduct.Text) - 20
      Cancel = True
   Else
      Cancel = False
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case SSTab1.Tab
      Case 0
         PopulateListBox
         If lstProducts.ListIndex > -1 Then
            txtProduct.Text = lstProducts.List(lstProducts.ListIndex)
            txtProduct.Tag = lstProducts.ItemData(lstProducts.ListIndex)
         Else
            txtProduct.Text = ""
            txtProduct.Tag = ""
         End If
      Case 1
      
      Case 2
   
   End Select
End Sub

