VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBloodBank 
   Caption         =   "Blood Bank"
   ClientHeight    =   5955
   ClientLeft      =   3660
   ClientTop       =   5805
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9660
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   5040
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   5040
      Width           =   1680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Blood Bank"
      TabPicture(0)   =   "frmBloodBank.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProducts"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDesc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Histology"
      TabPicture(1)   =   "frmBloodBank.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkTextBox"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Microbiology"
      TabPicture(2)   =   "frmBloodBank.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).ControlCount=   1
      Begin VB.CheckBox chkTextBox 
         Caption         =   "Show Textbox"
         Height          =   375
         Left            =   -71400
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00C0FFFF&
         Height          =   765
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmBloodBank.frx":0054
         Top             =   3240
         Width           =   6510
      End
      Begin VB.Frame Frame1 
         Caption         =   "Available Products"
         ForeColor       =   &H00C00000&
         Height          =   2685
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2025
         Begin VB.ListBox lstProducts 
            Height          =   1815
            Left            =   225
            TabIndex        =   7
            Top             =   360
            Width           =   1560
         End
      End
      Begin VB.Frame fraProducts 
         Caption         =   "Products selected for"
         Height          =   2670
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   6510
         Begin MSComCtl2.UpDown qScroll 
            Height          =   1800
            Left            =   6125
            TabIndex        =   2
            Top             =   390
            Width           =   240
            _ExtentX        =   503
            _ExtentY        =   3175
            _Version        =   393216
            OrigLeft        =   8640
            OrigTop         =   750
            OrigRight       =   8880
            OrigBottom      =   1350
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chkGAS 
            Caption         =   "Group and save"
            Height          =   240
            Left            =   2490
            TabIndex        =   4
            Top             =   2325
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   6045
            TabIndex        =   3
            Top             =   315
            Width           =   375
         End
         Begin vbAcceleratorGrid6.vbalGrid vbgBBank 
            Height          =   1920
            Left            =   195
            TabIndex        =   5
            Top             =   315
            Width           =   5910
            _ExtentX        =   10425
            _ExtentY        =   3387
            GridLines       =   -1  'True
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
            DisableIcons    =   -1  'True
         End
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
         Left            =   -72120
         TabIndex        =   12
         Top             =   1920
         Width           =   3255
      End
   End
   Begin MSComctlLib.ImageList VBGIcons 
      Left            =   0
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBloodBank.frx":011F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBloodBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private testId As String
Private thisTest As String
Private strSQL As String
Private BBCancelled As Boolean
Private strTemp As String
Private blnEditMode As Boolean
Private Const vbgText = 2
Private Const vbgQuantity = 3
Private Const vbgVariable = 4
Private Const vbgMandatory = 5
Private Const vbgHidden = 6
Private Const bbAmendText As String = "Note:" & vbTab & "Amending an entry will affect ALL blood bank records" & vbCrLf & _
                                      vbTab & "which reference this product, not just this test. Please" & vbCrLf & _
                                      vbTab & "be sure that this is the intention."
Private Const bbProductText As String = "To amend a value, click on the relevant cell and use the " & _
                                        "Up/Down arrows (right). " & vbCrLf & _
                                        "To add a blood bank products, double-click an  " & _
                                        "entry in the Available Products list " & vbCrLf & _
                                        "To remove and entry, click the red 'X' Icon (Left) "

Public Property Get CancelClicked() As Boolean
   CancelClicked = BBCancelled
End Property

Public Property Let SQLData(strNewValue As String)
   strSQL = strNewValue
End Property

Public Property Get SQLData() As String
   SQLData = strSQL
End Property

Public Property Let TestIndex(strNewValue As String)
   testId = strNewValue
End Property

Public Property Let TestName(strNewValue As String)
   thisTest = strNewValue
'   lblTestName.Caption = "Products currently associated with " & thisTest
End Property

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

Private Sub cmdCancel_Click()
   Me.Hide
   BBCancelled = True
   strSQL = ""
   Unload Me
End Sub

Private Sub CmdOk_Click()
   On Error GoTo procEH
   Dim i As Integer
   Dim RS As New ADODB.Recordset
   Dim blnReset As Boolean
   
   blnReset = False
   Select Case SSTab1.Tab
      Case 0
         With vbgBBank
'            loadCtrl.DeleteBloodBank testId
            If chkGAS.value > 0 Then
               strSQL = "INSERT INTO Request_BloodBank_Procedures " & _
                           "(Test_Index, Group_And_Save) " & _
                        "VALUES (" & _
                           testId & ", " & Abs(CInt(chkGAS.value)) & "); "
            End If
            
            If .rows > 0 Then
               For i = 1 To .rows
                  strSQL = strSQL & _
                           "INSERT INTO Request_BloodBank_Products " & _
                              "(Test_Index, Product_Name_Index, Quantity, Variable, Required) " & _
                           "VALUES (" & _
                              testId & ", " & .Cell(i, vbgHidden).Text & ", " & .Cell(i, vbgQuantity).Text & ", " & _
                              Abs(CInt((.Cell(i, vbgVariable).Text = "Variable"))) & ", " & _
                              Abs(CInt((.Cell(i, vbgMandatory).Text = "Required"))) & "); "
               Next i
            Else
               If chkGAS.value = 0 Then
                  blnReset = True
                  frmMain.edipr("TEST_TYPE").value = 0
               End If
            End If
'            BBCancelled = False
         End With
         
      Case 1
'         strSQL = "DELETE FROM Request_Histology_Input_Type " & _
                  "WHERE Test_Index = " & testId
'         ICECon.Execute strSQL
         If chkTextBox.value = 1 Then
            strSQL = "INSERT INTO Request_Histology_Input_Type (" & _
                        "Test_Index, ShowTextbox) " & _
                     "VALUES (" & _
                        testId & ", " & _
                        chkTextBox.value & ")"
         Else
            blnReset = True
            frmMain.edipr("TEST_TYPE").value = 0
         End If
         
      End Select
      Delete
      Me.Hide
      If blnReset = False Then
         frmMain.edipr("TEST_TYPE").value = SSTab1.Tab + 1
      End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.FurtherInfo = strSQL
   eClass.CurrentProcedure = "frmBloodBank.cmdOK_Click"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub
'
'Private Sub cmdPAdd_Click()
'   On Error GoTo procEH
'   Dim RS As New ADODB.Recordset
'   Dim iMax As Integer
'
'   strSQL = "SELECT MAX(Product_Name_Index) PNI " & _
'            "FROM Request_Bloodbank_Product_Names"
'   RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'   iMax = RS("PNI") + 1
'   RS.Close
'   strSQL = "INSERT INTO Request_BloodBank_Product_Names " & _
'               "(Product_Name_Index, Product_Name) " & _
'            "VALUES (" & _
'               iMax & ", '" & txtProduct.Text & "')"
'   ICECon.Execute strSQL
'   lstProducts.AddItem txtProduct.Text
'   lstProducts.ItemData(lstProducts.ListCount - 1) = iMax
'   txtProduct.Tag = iMax
'   Set RS = Nothing
'   Exit Sub
'
'procEH:
'   If eClass.Behaviour = -1 Then
'      Stop
'      Resume
'   End If
'   eClass.CurrentProcedure = "frmBloodBank.cmdPAdd_Click"
'   eClass.Add Err.Number, Err.Description, Err.Source
'End Sub
'
'Private Sub cmdPDelete_Click()
'   Dim RS As New ADODB.Recordset
'   Dim mbText As String
'   Dim pCount As Integer
'
'   If MsgBox("Delete " & txtProduct.Text & " from BloodBank products?", vbYesNo, "Confirm Delete") = vbYes Then
'      strSQL = "SELECT DISTINCT Request_BloodBank_Products.Test_Index, Request_Tests.Screen_Caption " & _
'               "FROM Request_BloodBank_Products " & _
'                  "INNER JOIN Request_Tests ON " & _
'                  "Request_BloodBank_Products.Test_Index = Request_Tests.Test_Index " & _
'               "WHERE Product_Name_Index = " & txtProduct.Tag
'      RS.Open strSQL, ICECon, adOpenKeyset, adLockReadOnly
'      If RS.EOF Then
'         strSQL = "DELETE FROM Request_BloodBank_Product_Names " & _
'                  "WHERE Product_Name_index = " & txtProduct.Tag
'         ICECon.Execute strSQL
'         lstProducts.RemoveItem lbIndex(txtProduct.Tag, True)
'      Else
'         mbText = "Unable to delete product. Product is used by:" & vbCrLf
'         pCount = 0
'         Do Until RS.EOF Or pCount > 10
'            mbText = mbText & RS!Screen_Caption & vbCrLf
'            RS.MoveNext
'            pCount = pCount + 1
'         Loop
'         If RS.EOF = False Then
'            mbText = mbText & "..." & vbCrLf
'         End If
'         mbText = mbText & "Remove product from these tests before deleting"
'         MsgBox mbText, vbInformation, txtProduct.Text
'      End If
'      RS.Close
'   End If
'   Set RS = Nothing
'End Sub
'
'Private Sub cmdPUpdate_Click()
'   lstProducts.List(lbIndex(txtProduct.Tag, True)) = txtProduct.Text
'   strSQL = "UPDATE Request_BloodBank_Product_Names " & _
'            "SET Product_Name = '" & txtProduct.Text & "' " & _
'            "WHERE Product_Name_Index = " & lstProducts.ItemData(lstProducts.ListIndex)
'   ICECon.Execute strSQL
'End Sub

Private Sub Delete()
   strSQL = "DELETE FROM Request_BloodBank_Procedures " & _
            "WHERE Test_index = " & testId & "; " & _
            "DELETE FROM Request_Bloodbank_Products " & _
            "WHERE Test_Index = " & testId & "; " & _
            "DELETE FROM Request_Histology_Input_Type " & _
            "WHERE Test_Index = " & testId & ";" & _
            strSQL
End Sub

Private Sub Form_Load()
   On Error GoTo procEH
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim intRow As Integer
   
'   testId = "0"
   PopulateListBox
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmBloodBank.Form_Load"
   eClass.Add Err.Number, Err.Description, Err.Source, False
End Sub

Private Sub lstProducts_DblClick()
   Dim intRow As Integer
   
   With vbgBBank
      .AddRow 1
      intRow = 1
      .CellIcon(intRow, 1) = 0
      .Cell(intRow, vbgText).Text = lstProducts.List(lstProducts.ListIndex)
      .Cell(intRow, vbgQuantity).Text = "1"
      .Cell(intRow, vbgVariable).Text = "Variable"
      .Cell(intRow, vbgMandatory).Text = "Optional"
      .Cell(intRow, vbgHidden).Text = lstProducts.ItemData(lstProducts.ListIndex)
      lstProducts.RemoveItem lstProducts.ListIndex
   End With
End Sub

Private Sub qScroll_DownClick()
   With vbgBBank
      Select Case .SelectedCol
         Case vbgQuantity
            If .SelectedCol = vbgQuantity And .SelectedRow > 0 And Val(.Cell(.SelectedRow, vbgQuantity).Text) > 1 Then
               .Cell(.SelectedRow, vbgQuantity).Text = Val(.Cell(.SelectedRow, vbgQuantity).Text) - 1
            End If
         
         Case vbgVariable
            If .Cell(.SelectedRow, vbgVariable).Text = "Variable" Then
               .Cell(.SelectedRow, vbgVariable).Text = "Fixed"
            Else
               .Cell(.SelectedRow, vbgVariable).Text = "Variable"
            End If
         
         Case vbgMandatory
            If .Cell(.SelectedRow, vbgMandatory).Text = "Required" Then
               .Cell(.SelectedRow, vbgMandatory).Text = "Optional"
            Else
               .Cell(.SelectedRow, vbgMandatory).Text = "Required"
            End If
      
      End Select
   End With
End Sub

Private Sub qScroll_UpClick()
   With vbgBBank
      Select Case .SelectedCol
         Case vbgQuantity
            If .SelectedRow > 0 And Val(.Cell(.SelectedRow, vbgQuantity).Text) < 10 Then
               .Cell(.SelectedRow, vbgQuantity).Text = Val(.Cell(.SelectedRow, vbgQuantity).Text) + 1
            End If
         
         Case vbgVariable
            If .Cell(.SelectedRow, vbgVariable).Text = "Variable" Then
               .Cell(.SelectedRow, vbgVariable).Text = "Fixed"
            Else
               .Cell(.SelectedRow, vbgVariable).Text = "Variable"
            End If
         
         Case vbgMandatory
            If .Cell(.SelectedRow, vbgMandatory).Text = "Required" Then
               .Cell(.SelectedRow, vbgMandatory).Text = "Optional"
            Else
               .Cell(.SelectedRow, vbgMandatory).Text = "Required"
            End If
      
      End Select
   End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim intRow As Integer
   
   Select Case SSTab1.Tab
      Case 0
         txtDesc.Text = bbProductText
         intRow = 1
         
         lstProducts.ListIndex = 0
         strSQL = "SELECT * " & _
                  "FROM Request_BloodBank_Procedures " & _
                  "WHERE Test_Index = " & testId
         eClass.FurtherInfo = strSQL
         RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         If RS.BOF And RS.EOF Then
            chkGAS.value = 0
         Else
            chkGAS.value = Abs(CInt(RS!Group_And_Save))
         End If
         RS.Close
         PopulateListBox
         With vbgBBank
            .Clear True
            .ImageList = VBGIcons
            .AddColumn , "X", ecgHdrTextALignCentre, 1, 20, , True
            .AddColumn , "Blood Bank Product", ecgHdrTextALignLeft, , 120
            .AddColumn , "Quantity", ecgHdrTextALignRight, , 60
            .AddColumn , "Allow Changes", ecgHdrTextALignLeft, , 90
            .AddColumn , "Mandatory", ecgHdrTextALignLeft, , 80
            .AddColumn , , , , 0, False
      '      .Width = 360 * Screen.TwipsPerPixelX
         
            strSQL = "SELECT * " & _
                     "FROM Request_Bloodbank_Products " & _
                     "WHERE Test_Index = " & testId
            RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
            If RS.RecordCount > 0 Then
      '         intRow = 1
               Do Until RS.EOF
                  .AddRow
                  .CellIcon(intRow, 1) = 1
                  .CellForeColor(intRow, 1) = vbRed
                  .Cell(intRow, 1).Text = "X"
                  strSQL = "SELECT * " & _
                           "FROM Request_Bloodbank_Product_Names " & _
                           "WHERE Product_Name_Index = " & RS!Product_Name_Index
                  RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
                  .Cell(intRow, vbgText).Text = RS2!Product_Name
                  .Cell(intRow, vbgText).ForeColor = vbBlack
                  .Cell(intRow, vbgHidden).Text = RS2!Product_Name_Index
                  lstProducts.RemoveItem lbIndex(RS2!Product_Name)
                  RS2.Close
                  
                  .Cell(intRow, vbgQuantity).Text = RS!Quantity
                  .CellTextAlign(intRow, vbgQuantity) = DT_RIGHT
                  If RS!variable Then
                     .Cell(intRow, vbgVariable).Text = "Variable"
                  Else
                     .Cell(intRow, vbgVariable).Text = "Fixed"
                  End If
                  
                  If RS!Required Then
                     .Cell(intRow, vbgMandatory).Text = "Required"
                  Else
                     .Cell(intRow, vbgMandatory).Text = "Optional"
                  End If
                  intRow = intRow + 1
                  RS.MoveNext
               Loop
               
               .SelectedCol = 2
               .SelectedRow = 1
            End If
      '      .AddRow
      '      .Cell(intRow, 1).Text = "New"
      '      .CellBackColor(intRow, 1) = &H8000000F
      '      .CellForeColor(intRow, 1) = BPGREEN
         End With
         RS.Close
         cmdOK.Enabled = True
      
      Case 1
         strSQL = "SELECT * " & _
                  "FROM Request_Histology_Input_Type " & _
                  "WHERE Test_Index = " & testId
         RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
         If RS.EOF = False Then
            chkTextBox.value = Abs(CInt(RS!ShowTextbox))
         End If
         RS.Close
         cmdOK.Enabled = True
         
      Case 2
         cmdOK.Enabled = False
         
   End Select
   Set RS = Nothing
   Set RS2 = Nothing
End Sub

Private Sub vbgBBank_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   With vbgBBank
      If .SelectedCol = 1 Then
         If MsgBox("Remove " & .Cell(.SelectedRow, vbgText).Text, vbYesNo, "Confirm Removal") = vbYes Then
            lstProducts.AddItem .Cell(.SelectedRow, vbgText).Text
            .RemoveRow .SelectedRow
         End If
      End If
   End With
End Sub
