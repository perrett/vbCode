VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestMnt 
   Caption         =   "Included Test"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid flexTest 
      Height          =   3615
      Left            =   15
      TabIndex        =   3
      Top             =   660
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   6376
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4410
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   450
      TabIndex        =   1
      Top             =   4410
      Width           =   1335
   End
   Begin VB.Label lblQualF 
      BackColor       =   &H00C0E0FF&
      Caption         =   "<No qualifying tests>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   915
      TabIndex        =   5
      Top             =   2085
      Width           =   2220
   End
   Begin VB.Label lblQualB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   15
      TabIndex        =   4
      Top             =   660
      Width           =   4245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double-click on the test you wish to add."
      Height          =   420
      Left            =   15
      TabIndex        =   0
      Top             =   120
      Width           =   4230
   End
End
Attribute VB_Name = "frmTestMnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private curTestIndex As Long
Private curTestCode As String
Private nType As String

Private Sub cmdAdd_Click()
   Dim strSQL As String
   
'   If flexTest.Row > 0 Then
      Select Case nType
         Case "INC"
            strSQL = "INSERT INTO Request_Included_Tests " & _
                        "(Test_Index, Included_Test_Index)" & _
                        "VALUES (" & curTestIndex & ", " & flexTest.TextMatrix(flexTest.Row, 0) & ")"
         
         Case "EXC"
            strSQL = "INSERT INTO Request_Excluded_Tests " & _
                     "(Test_Index, Excluded_Test_Index) " & _
                     "VALUES (" & curTestIndex & ", " & flexTest.TextMatrix(flexTest.Row, 0) & ")"
         
         Case "RFX"
            strSQL = "INSERT INTO Request_Reflex_Tests " & _
                     "(Test_Index, Reflex_Test_Index) " & _
                     "VALUES (" & curTestIndex & ", " & flexTest.TextMatrix(flexTest.Row, 0) & ")"
         
         Case "PRF"
            strSQL = "INSERT INTO Request_Profile_Tests " & _
                     "(Profile_Index, Profile_Test_Index) " & _
                     "VALUES (" & curTestIndex & ", " & flexTest.TextMatrix(flexTest.Row, 0) & ")"
         
            
      End Select
      iceCon.Execute strSQL
'   End If
   Unload Me
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Property Let CurrentTest(Index As Long)
   curTestIndex = Index
End Property

Public Property Let CurrentTestCode(strNewValue As String)
   curTestCode = strNewValue
End Property

Public Property Let NodeType(strNewValue As String)
   nType = strNewValue
End Property

Private Sub flexTest_DblClick()
   cmdAdd_Click
End Sub

Private Sub flexTest_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim i As Integer
   
   If KeyCode >= 65 And KeyCode <= 90 Then
      i = 0
      With flexTest
         Do Until Asc(Mid(.TextMatrix(i, 1), 2, 1)) >= KeyCode
            i = i + 1
         Loop
         If Asc(Mid(.TextMatrix(i, 1), 2, 1)) > KeyCode Then
            i = i - 1
         End If
         .Row = i
         .Col = 0
         .ColSel = 1
         .TopRow = i
         .SetFocus
      End With
   End If
End Sub

Private Sub Form_Load()
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim provID As Integer
   
   With flexTest
      .Row = 0
      .cols = 2
      .ColWidth(0) = 0
      .ColWidth(1) = 4430
      .CellAlignment = 1
      .SelectionMode = flexSelectionByRow
   
      provID = objTV.NodeLevel(objTV.TopLevelNode)
      Select Case nType
         Case "INC"
            strSQL = "SELECT DISTINCT Screen_Caption, Test_Index " & _
                     "FROM Request_Tests " & _
                     "WHERE Test_Index NOT IN " & _
                        "(SELECT Included_Test_Index " & _
                        "FROM Request_Included_Tests " & _
                        "WHERE Test_Index = " & curTestIndex & ") " & _
                           "AND Test_Index NOT IN " & _
                           "(SELECT Excluded_Test_Index " & _
                           "FROM Request_Excluded_Tests " & _
                           "WHERE Test_Index = " & curTestIndex & ") " & _
                              "AND Test_Index NOT IN " & _
                              "(SELECT Reflex_Test_Index " & _
                              "FROM Request_Reflex_Tests " & _
                              "WHERE Test_Index = " & curTestIndex & ") " & _
                        "AND Test_Index <> " & curTestIndex & _
                     " ORDER BY Screen_Caption"
         
         Case "EXC"
            strSQL = "SELECT Screen_Caption, Test_Index " & _
                     "FROM Request_Tests " & _
                     "WHERE Test_Index NOT IN " & _
                        "(SELECT Excluded_Test_Index " & _
                        "FROM Request_Excluded_Tests " & _
                        "WHERE Test_Index = " & curTestIndex & ") " & _
                           "AND Test_Index NOT IN " & _
                           "(SELECT Included_Test_Index " & _
                           "FROM Request_Included_Tests " & _
                           "WHERE Test_Index = " & curTestIndex & ") " & _
                              "AND Test_Index NOT IN " & _
                              "(SELECT Reflex_Test_Index " & _
                              "FROM Request_Reflex_Tests " & _
                              "WHERE Test_Index = " & curTestIndex & ") " & _
                        "AND Test_Index <> " & curTestIndex & _
                     " ORDER BY Screen_Caption"
         
         Case "RFX"
            strSQL = "SELECT Screen_Caption, Test_Index " & _
                     "FROM Request_Tests " & _
                     "WHERE Test_Index NOT IN " & _
                        "(SELECT Reflex_Test_Index " & _
                        "FROM Request_Reflex_Tests " & _
                        "WHERE Test_Index = " & curTestIndex & ") " & _
                           "AND Test_Index NOT IN " & _
                           "(SELECT Included_Test_Index " & _
                           "FROM Request_Included_Tests " & _
                           "WHERE Test_Index = " & curTestIndex & ") " & _
                              "AND Test_Index NOT IN " & _
                              "(SELECT Excluded_Test_Index " & _
                              "FROM Request_Excluded_Tests " & _
                              "WHERE Test_Index = " & curTestIndex & ") " & _
                        "AND Test_Index <> " & curTestIndex & _
                     " ORDER BY Screen_Caption"

'                        " AND Provider_ID = " & provID & _

         Case "PRF"
            strSQL = "SELECT Screen_Caption, Test_Index as Test_Index " & _
                     "FROM Request_Tests " & _
                     "WHERE Test_Index NOT IN " & _
                        "(SELECT Profile_Test_Index " & _
                        "FROM Request_Profile_Tests " & _
                        "WHERE Profile_Index = " & curTestIndex & ") " & _
                           "AND Test_Index <> " & curTestIndex & _
                     " ORDER BY Screen_Caption"
            
         
      End Select
      RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
      If RS.RecordCount > 0 Then
         lblQualB.Visible = False
         lblQualF.Visible = False
         .Visible = True
         i = 0
         .rows = RS.RecordCount
         
         Do Until RS.EOF
            .TextMatrix(i, 0) = Trim(RS!Test_Index & "")
            .TextMatrix(i, 1) = " " & RS!Screen_Caption
            RS.MoveNext
            i = i + 1
         Loop
      Else
         lblQualB.Visible = True
         lblQualF.Visible = True
         cmdAdd.Enabled = False
         .Visible = False
      End If
      RS.Close
   End With
   flexTest.Row = 0
   flexTest.Col = 0
   Set RS = Nothing
End Sub
