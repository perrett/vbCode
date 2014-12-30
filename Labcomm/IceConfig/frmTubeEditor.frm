VERSION 5.00
Object = "{3F118CA4-97A8-4926-AA0A-6FD80DFB3DCE}#2.6#0"; "PrpList2.ocx"
Begin VB.Form frmTubeEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tube Editor"
   ClientHeight    =   5415
   ClientLeft      =   2655
   ClientTop       =   2295
   ClientWidth     =   4710
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   2055
   End
   Begin PropertiesListCtl.PropertiesList TL1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8493
      LicenceData     =   "043E5E24294320255427005827381117235F203E5E386C6135395D741F41353E5A313F"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Ownerdraw       =   1
      MaxDropItems    =   0
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmTubeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tubeID As String
Private strSQL As String
Private Type RECT
 Left   As Long
 Top    As Long
 Right  As Long
 Bottom As Long
End Type

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function DrawText& Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long)
Private Declare Function FillRect& Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Sub Command1_Click()
   If tubeID <> "" Then
      TempStr = "Update Request_Tubes Set "
      TempStr = TempStr + "Code='" & TL1.PropertyItems("TUBE_CODE").value & "',"
      TempStr = TempStr + "Name='" & TL1.PropertyItems("TUBE_NAME").value & "',"
      TempStr = TempStr + "Description='" & TL1.PropertyItems("TUBE_DESC").value & "',"
      If Format(TL1.PropertyItems("VOLUME").value) & "" = "" Then
         TempStr = TempStr + "Vol=NULL,"
      Else
         TempStr = TempStr + "Vol=" & Format(TL1.PropertyItems("VOLUME").value) + ","
      End If
      If Format(TL1.PropertyItems("MIN_VOL").value) & "" = "" Then
         TempStr = TempStr + "Min_Vol=NULL,"
      Else
         TempStr = TempStr + "Min_Vol=" & Format(TL1.PropertyItems("MIN_VOL").value) + ","
      End If
      'TempStr = TempStr + "Paed_Min_Vol=" & Format(TL1.PropertyItems("PAED_MIN").value) + ","
      TempStr = TempStr + "Colour=" & Left(TL1.PropertyItems("COLOUR").Tag, InStr(TL1.PropertyItems("COLOUR").Tag, "-") - 1) + ","
      TempStr = TempStr + "EDI_SampleCode='" & TL1.PropertyItems("EDI_CODE").value + "',"
      TempStr = TempStr + "EDI_SampleText='" & TL1.PropertyItems("EDI_TEXT").value + "',"
      TempStr = TempStr + "Stock_Code='" & TL1.PropertyItems("STOCK_CODE").value + "',"
      TempStr = TempStr + "Unit_Of_Issue='" & Format(TL1.PropertyItems("UNIT").value) + "',"
      TempStr = TempStr + "ReOrder_Qty='" & Format(TL1.PropertyItems("REORDER").value) + "' "
      TempStr = TempStr + "Where Tube_Index=" & tubeID
      Debug.Print TempStr
      iceCon.Execute TempStr
   Else
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
      RS.Open "Select Max(Tube_Index) 'TubeMax' From Request_Tubes", iceCon, adOpenKeyset, adLockReadOnly
      If Format(RS!TubeMax & "") <> "" Then
         TI = RS!TubeMax + 1
      Else
         TI = 1
      End If
      RS.Close
      TempStr = "Insert Into Request_Tubes (Tube_Index,Date_Added,Code,Name,Description,Vol,Min_Vol,Paed_Min_Vol,Colour,EDI_SampleCode,EDI_SampleText,Stock_Code,Unit_Of_Issue,ReOrder_Qty) Values ("
      TempStr = TempStr + Format(TI) + ",'" + Format(Date, "DD MMM YYYY") + "','"
      TempStr = TempStr + TL1.PropertyItems("TUBE_CODE").value & "','"
      TempStr = TempStr + TL1.PropertyItems("TUBE_NAME").value & "','"
      TempStr = TempStr + TL1.PropertyItems("TUBE_DESC").value & "',"
      If Format(TL1.PropertyItems("VOLUME").value) & "" = "" Then
         TempStr = TempStr + "NULL,"
      Else
         TempStr = TempStr + Format(TL1.PropertyItems("VOLUME").value) & ","
      End If
      If Format(TL1.PropertyItems("MIN_VOL").value) & "" = "" Then
         TempStr = TempStr + "NULL,NULL,"
      Else
         TempStr = TempStr + Format(TL1.PropertyItems("MIN_VOL").value) & ",NULL,"
      End If
      'TempStr = TempStr + Format(TL1.PropertyItems("PAED_MIN").value) & ","
      If TL1.PropertyItems("COLOUR").Tag = "" Then
         TempStr = TempStr & "NULL,'"
      Else
         TempStr = TempStr + Left(TL1.PropertyItems("COLOUR").Tag, InStr(TL1.PropertyItems("COLOUR").Tag, "-") - 1) & ",'"
      End If
      TempStr = TempStr + TL1.PropertyItems("EDI_CODE").value & "','"
      TempStr = TempStr + TL1.PropertyItems("EDI_TEXT").value & "','"
      TempStr = TempStr + TL1.PropertyItems("STOCK_CODE").value & "','"
      TempStr = TempStr + Format(TL1.PropertyItems("UNIT").value) & "','"
      TempStr = TempStr + Format(TL1.PropertyItems("REORDER").value) & "')"
      Debug.Print TempStr
      iceCon.Execute TempStr
   End If
   Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

'Private Sub Form_Activate()
'    If Label1.Caption <> "" Then
'        Dim RS As ADODB.Recordset
'        Dim RS2 As ADODB.Recordset
'        Set RS = New ADODB.Recordset
'        Set RS2 = New ADODB.Recordset
'        RS.Open "Select * From Request_Tubes Where Tube_Index=" & Label1.Caption, ICECon, adOpenKeyset, adLockReadOnly
'        TL1.PropertyItems("TUBE_CODE").value = RS!Code
'        TL1.PropertyItems("TUBE_NAME").value = RS!Name
'        TL1.PropertyItems("TUBE_DESC").value = RS!Description
'        TL1.PropertyItems("VOLUME").value = RS!Vol
'        TL1.PropertyItems("MIN_VOL").value = RS!Min_VOl
'        'TL1.PropertyItems("PAED_MIN").value = RS!Paed_Min_Vol
'        If Format(RS!Colour) <> "" Then
'            RS2.Open "Select * From Colours Where Colour_Index=" & RS!Colour, ICECon, adOpenKeyset, adLockReadOnly
'            TL1.PropertyItems("COLOUR").value = Val(RS2!Colour_Code)
'            TL1.PropertyItems("COLOUR").Tag = Format(RS!Colour) + "-" + RS2!Colour_Name
'            RS2.Close
'        End If
'        TL1.PropertyItems("EDI_CODE").value = RS!EDI_SampleCode
'        TL1.PropertyItems("EDI_TEXT").value = RS!EDI_SampleText
'        TL1.PropertyItems("STOCK_CODE").value = RS!Stock_Code
'        TL1.PropertyItems("UNIT").value = Val(Format(RS!Unit_Of_Issue))
'        TL1.PropertyItems("REORDER").value = Val(Format(RS!Reorder_Qty))
'        RS.Close
'        TL1.Refresh True
'    End If
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        MsgBox "The character ' is not permitted, please use the ` character instead"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    TL1.PropertyItems.Add "TUBE_CODE", "Tube Code", plpsString, "Tube01", "Lab tube reference ID"
    TL1.PropertyItems("TUBE_CODE").max = 6
    TL1.PropertyItems.Add "TUBE_NAME", "Tube Name", plpsString, "New Tube", "Generic tube name"
    TL1.PropertyItems("TUBE_NAME").max = 25
    TL1.PropertyItems.Add "TUBE_DESC", "Tube Description", plpsString, "New Tube", "Generic description of tube"
    TL1.PropertyItems("TUBE_DESC").max = 50
    TL1.PropertyItems.Add "VOLUME", "Volume", plpsNumber, , "Maximum volume of tube in millilitres"
    With TL1.PropertyItems("VOLUME")
      .Min = 0
      .max = 10000
      .Increment = 1
      .defaultValue = 0
      .value = 0
    End With
    TL1.PropertyItems.Add "MIN_VOL", "Minimum Volume", plpsNumber, , "Minimum volume of sample required to perform tests"
    With TL1.PropertyItems("MIN_VOL")
      .Min = 0
      .max = 10000
      .Increment = 1
      .defaultValue = 0
      .value = 0
    End With
    'TL1.PropertyItems.Add "PAED_MIN","Paediatric Minimum Volume",
    TL1.PropertyItems.Add "COLOUR", "Colour", plpsColor, , "Colour of tube cap"
    TL1.PropertyItems.Add "EDI_CODE", "EDI Sample Code", plpsString, , "EDI code for sample type"
    TL1.PropertyItems("EDI_CODE").max = 8
    TL1.PropertyItems.Add "EDI_TEXT", "EDI Sample Text", plpsString, , "EDI text for sample type"
    TL1.PropertyItems("EDI_TEXT").max = 65
    TL1.PropertyItems.Add "STOCK_CODE", "Stock Code", plpsString, , "Internal stock code for tube"
    TL1.PropertyItems("STOCK_CODE").max = 25
    TL1.PropertyItems.Add "UNIT", "Unit of Issue", plpsNumber, , "Units in which tubes are issued"
    With TL1.PropertyItems("UNIT")
      .Min = 0
      .max = 10000
      .Increment = 5
      .defaultValue = 0
      .value = 0
    End With
    TL1.PropertyItems.Add "REORDER", "Re-Order Quantity", plpsNumber, , "At what level of stock should a re-order be made"
    With TL1.PropertyItems("REORDER")
      .Min = 0
      .max = 100
      .Increment = 5
      .defaultValue = 0
      .value = 0
    End With
    If tubeID <> "" Then
      TubeDetails
   End If
End Sub

Private Sub TubeDetails()
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   
   strSQL = "SELECT * " & _
            "FROM Request_Tubes " & _
            "WHERE Tube_Index = " & tubeID
   RS.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
   With TL1
      .PropertyItems("TUBE_CODE").value = RS!Code
      .PropertyItems("TUBE_NAME").value = RS!Name
      .PropertyItems("TUBE_DESC").value = RS!Description
      .PropertyItems("VOLUME").value = RS!Vol
      .PropertyItems("MIN_VOL").value = RS!Min_VOl
      If Format(RS!Colour) <> "" Then
         strSQL = "SELECT * " & _
                  "FROM Colours " & _
                  "WHERE Colour_Index = " & RS!Colour
          RS2.Open strSQL, iceCon, adOpenKeyset, adLockReadOnly
         .PropertyItems("COLOUR").value = Val(RS2!Colour_Code)
         .PropertyItems("COLOUR").Tag = Format(RS!Colour) + "-" + RS2!Colour_Name
          RS2.Close
      End If
      .PropertyItems("EDI_CODE").value = RS!EDI_SampleCode
      .PropertyItems("EDI_TEXT").value = RS!EDI_SampleText
      .PropertyItems("STOCK_CODE").value = RS!Stock_Code
      .PropertyItems("UNIT").value = Val(Format(RS!Unit_Of_Issue))
      .PropertyItems("REORDER").value = Val(Format(RS!Reorder_Qty))
      RS.Close
      .Refresh True
   End With
   Set RS = Nothing
   Set RS2 = Nothing
End Sub

Private Sub TL1_AfterEdit(PropertyItem As PropertiesListCtl.PropertyItem, newValue As Variant, Cancel As Boolean)
   Dim blnNoEntry As Boolean
   
   blnNoEntry = (newValue = "")
   If Left(PropertyItem.Key, 4) <> "TUBE" Then
      blnNoEntry = False
   End If
   If blnNoEntry Then
      MsgBox PropertyItem.Caption & " must not be blank", _
             vbExclamation, "Mandatory field missing"
      newValue = PropertyItem.value
   End If
   Cancel = blnNoEntry
End Sub

Private Sub TL1_BeforeEdit(PropertyItem As PropertiesListCtl.PropertyItem, Cancel As Boolean)
    If PropertyItem.Style = plpsColor Then
        frmColour.Show 1
        If PickedColIndex > 0 Then
            Debug.Print "Tube - " & PickedColIndex
            Debug.Print "Tube - " & PickedColName
            PropertyItem.Tag = Format(PickedColIndex) + "-" + PickedColName
            PropertyItem.value = PickedCol
        End If
        'PropertyItem.Description = PropertyItem.Tag
        Cancel = True
    End If
End Sub

Private Sub TL1_RequestDisplayValue(PropertyItem As PropertiesListCtl.PropertyItem, DisplayValue As String)
   If PropertyItem.Tag <> "" Then
      If PropertyItem.Key = "COLOUR" Then
         DisplayValue = Mid$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") + 1, Len(PropertyItem.Tag) - InStr(PropertyItem.Tag, "-") + 1) + " (" + Left$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") - 1) + ")"
      End If
   End If
End Sub

Public Property Let TubeIndex(strNewValue As String)
   tubeID = strNewValue
End Property
