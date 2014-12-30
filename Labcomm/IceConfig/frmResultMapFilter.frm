VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmResultMapFilter 
   Caption         =   "View Read Codes"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sTabRC 
      Height          =   1860
      Left            =   45
      TabIndex        =   2
      Top             =   75
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3281
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Filter"
      TabPicture(0)   =   "frmResultMapFilter.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraFilter"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Find"
      TabPicture(1)   =   "frmResultMapFilter.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraSearch"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraSearch 
         Caption         =   "Search for..."
         Height          =   1260
         Left            =   105
         TabIndex        =   8
         Top             =   465
         Width           =   4230
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   150
            TabIndex        =   12
            Top             =   510
            Width           =   2355
         End
         Begin VB.OptionButton optFind 
            Caption         =   "Local Code"
            Height          =   270
            Index           =   0
            Left            =   2655
            TabIndex        =   11
            Top             =   180
            Width           =   1425
         End
         Begin VB.OptionButton optFind 
            Caption         =   "Test Description"
            Height          =   270
            Index           =   1
            Left            =   2655
            TabIndex        =   10
            Top             =   510
            Width           =   1500
         End
         Begin VB.OptionButton optFind 
            Caption         =   "Read Code"
            Height          =   270
            Index           =   2
            Left            =   2655
            TabIndex        =   9
            Top             =   840
            Width           =   1200
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   "Only show..."
         Height          =   1260
         Left            =   -74880
         TabIndex        =   3
         Top             =   435
         Width           =   4230
         Begin VB.OptionButton optFilter 
            Caption         =   "Non Read-coded tests"
            Height          =   240
            Index           =   0
            Left            =   195
            TabIndex        =   7
            Top             =   330
            Width           =   1905
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "Inactive Tests"
            Height          =   240
            Index           =   1
            Left            =   195
            TabIndex        =   6
            Top             =   705
            Width           =   1815
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "Suppressed  tests"
            Height          =   240
            Index           =   2
            Left            =   2235
            TabIndex        =   5
            Top             =   330
            Width           =   1905
         End
         Begin VB.OptionButton optFilter 
            Caption         =   "Flagged as deleted"
            Height          =   240
            Index           =   3
            Left            =   2220
            TabIndex        =   4
            Top             =   705
            Width           =   1905
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   420
      Left            =   2625
      TabIndex        =   1
      Top             =   2115
      Width           =   1230
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   405
      Left            =   630
      TabIndex        =   0
      Top             =   2130
      Width           =   1305
   End
End
Attribute VB_Name = "frmResultMapFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private FilterText As String

Public Property Get FilterStatus() As String
   FilterStatus = FilterText
End Property

Private Sub cmdCancel_Click()
   FilterText = "Filter..."
   Me.Hide
End Sub

Private Sub CmdOk_Click()
   Dim mainSQL As String
   Dim tmpSQL As String
   Dim i As Integer
   Dim tLen As Integer
   
   mainSQL = "SELECT ec.*, Sample_Text " & _
             "FROM EDI_InvTest_Codes ec " & _
                "INNER JOIN EDI_Local_Trader_Settings el " & _
                "ON ec.EDI_LTS_Index = el.EDI_LTS_Index " & _
                "INNER JOIN CRIR_Sample_Type " & _
                "ON EDI_Sample_TypeCode = Sample_Code " & _
             "WHERE ec.Organisation = '" & frmMain.cboTrust.Text & "' " ' & _
                "AND ec.EDI_LTS_index = " & frmMain.CurrentLTSIndex & " "

   If sTabRC.Tab = 1 Then
      
      For i = 0 To 2
         If optFind(i).value Then
            Exit For
         End If
      Next i
      
      Select Case i
         Case 0
            strSQL = "AND EDI_Local_Test_Code like '%" & txtFind.Text & "%'"
                     
         Case 1
            strSQL = "AND EDI_Local_Rubric like '%" & txtFind.Text & "%'"
         
         Case 2
            If Len(txtFind.Text) = 5 Then
               strSQL = "AND Convert(Binary(5),EDI_Read_Code) = Convert(Binary(5),'" & txtFind.Text & "')"
            Else
               strSQL = "AND EDI_Read_Code like '" & txtFind.Text & "' + '%'"
            End If
         
      End Select
      If sTabRC.TabEnabled(0) = False Then
         i = 10
      End If
   Else
      
      For i = 0 To 3
         If optFilter(i).value Then
            Exit For
         End If
      Next i
      
      Select Case i
         Case 0
            strSQL = "AND (EDI_Read_Code IS Null OR EDI_Read_Code = '')"
         Case 1
            strSQL = "AND EDI_Op_Active = 0"
         
         Case 2
            strSQL = "AND EDI_OP_Suppress = 1"
                     
         Case 3
            loadCtrl.FirstViewSQL = "SELECT DISTINCT Left(EDI_Local_Rubric,1) As Initial " & _
                                    "FROM EDI_InvTest_Codes " & _
                                       "INNER JOIN Read_Codes " & _
                                       "ON Convert(Binary(5),EDI_Read_Code) = Convert(Binary(5),Read_V2RC) " & _
                                    "WHERE Organisation = '" & frmMain.cboTrust.Text & "' " & _
                                       "AND Read_Status = 'D'"
'                                    " ORDER BY Initial"

            
            loadCtrl.LocalTestsSQL = "SELECT * " & _
                                     "FROM EDI_InvTest_Codes ec " & _
                                        "INNER JOIN Read_Codes " & _
                                        "ON Convert(Binary(5),EDI_Read_Code) = Convert(Binary(5),Read_V2RC) " & _
                                        "LEFT JOIN CRIR_Sample_Type " & _
                                        "ON EDI_Sample_TypeCode = Sample_Code " & _
                                     "WHERE Organisation = '" & frmMain.cboTrust.Text & "' " & _
                                        "AND Read_Status = 'D'"
      End Select
      
   End If
   
   If i < 3 Then
'                                 "AND EDI_LTS_Index = " & frmMain.CurrentLTSIndex & " "
      loadCtrl.FirstViewSQL = "SELECT DISTINCT Left(EDI_Local_Rubric,1) As Initial " & _
                              "FROM EDI_InvTest_Codes " & _
                              "WHERE Organisation = '" & frmMain.cboTrust.Text & "' " & _
                              strSQL
'                              " ORDER BY Initial"
      loadCtrl.LocalTestsSQL = mainSQL & _
                               strSQL
   End If
   
   FilterText = "Clear filter"
   Me.Hide
End Sub

Private Sub Form_Load()
   optFind(0).value = True
   optFilter(0).value = True
End Sub

Private Sub fraSearch_Click()
   fraSearch.Appearance = 0
   fraFilter.Appearance = 1
End Sub

Private Sub frafilter_Click()
   fraFilter.Appearance = 0
   fraSearch.Appearance = 1
End Sub
