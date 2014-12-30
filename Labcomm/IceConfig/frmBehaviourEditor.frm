VERSION 5.00
Object = "{3F118CA4-97A8-4926-AA0A-6FD80DFB3DCE}#2.6#0"; "PrpList2.ocx"
Begin VB.Form frmBehaviourEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Behaviour Editor"
   ClientHeight    =   6105
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   4935
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin PropertiesListCtl.PropertiesList BS 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9763
      LicenceData     =   "12212D232721363A27200E3A312762102D3D36212D3F621127212C3A277312363021272736"
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmBehaviourEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BN As String
Private testId As Long

Public Property Let TestIndex(lngNewValue As Long)
   testId = lngNewValue
End Property

Private Sub BS_AfterEdit(PropertyItem As PropertiesListCtl.PropertyItem, NewValue As Variant, Cancel As Boolean)
    If PropertyItem.Key = "STYLE" And Not IsNull(PropertyItem.value) And (PropertyItem.value <> 0 Or NewValue <> "") Then
        BS.PropertyItems.Clear
        Select Case NewValue
            Case "QUE"
                SetupBS ("QUE")
                SetupBSQuestion
                SetupBS2
            Case "DEN"
                SetupBS ("DEN")
                SetupBSData
                SetupBS2
            Case "MCD"
                SetupBS ("MCD")
                'SetupBSClin
                'SetupBS2
            Case "EIM"
                SetupBS ("EIM")
                'SetupBSEIM
                'SetupBS2
            Case "EIF"
                SetupBS ("EIF")
                'SetupBSEIF
                'SetupBS2
            Case "CNL"
                SetupBS ("CNL")
            Case "HLP"
                SetupBS ("HLP")
                SetupBSHLP
            Case Null
                SetupBS Null
        End Select
        If BN <> "" Then BS.PropertyItems("NAME").value = BN
    End If
    If PropertyItem.Key = "TYPE" And Not IsNull(PropertyItem.value) Then
        If PropertyItem.value = "N" Then
            BS.PropertyItems("NUMMAX").Enabled = True
            BS.PropertyItems("NUMMIN").Enabled = True
            BS.PropertyItems("PICKLIST").Enabled = False
        ElseIf PropertyItem.value = "P" Then
            BS.PropertyItems("PICKLIST").Enabled = True
        Else
            BS.PropertyItems("PICKLIST").Enabled = False
            BS.PropertyItems("PICKLIST").value = Null
            BS.PropertyItems("NUMMAX").Enabled = False
            BS.PropertyItems("NUMMAX").value = Null
            BS.PropertyItems("NUMMIN").Enabled = False
            BS.PropertyItems("NUMMIN").value = Null
        End If
    End If
End Sub

Private Sub BS_BeforeEdit(PropertyItem As PropertiesListCtl.PropertyItem, Cancel As Boolean)
    If PropertyItem.Key = "STYLE" And PropertyItem.value <> 0 Then
        If MsgBox("Changing the style will cause changed property details to be lost.  Are you sure you wish to change the behaviour style?", vbQuestion + vbYesNo, "Change Behaviour Style") = vbNo Then Cancel = True
        BN = BS.PropertyItems("NAME").value
    ElseIf PropertyItem.Key = "STYLE" Then
        BN = BS.PropertyItems("NAME").value & ""
    End If
End Sub

Private Sub Command1_Click()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    If testId > -1 Then
        tempStr1 = "Update Request_Prompt Set Prompt_Desc='" & BS.PropertyItems("NAME").value & "',"
        tempstr2 = " Where Prompt_Index=" & testId
        Select Case BS.PropertyItems("STYLE").value
            Case "QUE"
                If Format(BS.PropertyItems("YESACTION").value & "") = "" Then BS.PropertyItems("YESACTION").value = "NULL"
                If Format(BS.PropertyItems("NOACTION").value & "") = "" Then BS.PropertyItems("NOACTION").value = "NULL"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Type='QUE',"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Text='" & BS.PropertyItems("QUESTION").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "Yes_Text='" & BS.PropertyItems("YESTEXT").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "No_Text='" & BS.PropertyItems("NOTEXT").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "Yes_Action_Type=" & BS.PropertyItems("YESACTION").value & ","
                tempStr1 = tempStr1 + vbCrLf + "No_Action_Type=" & BS.PropertyItems("NOACTION").value & ","
                tempStr1 = tempStr1 + vbCrLf + "Cancel_Text='" & BS.PropertyItems("CANCELTEXT").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "Save_As_Type='" & BS.PropertyItems("SAVETYPE").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "Save_As_String='" & BS.PropertyItems("SAVEHEADER").value & "'"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "DEN"
                'Prompt_Text,DataEntry_Type,DataEntry_Upper_Val,DataEntry_Lower_Val,PickList_Index,Save_As_Type,Save_As_String"
                If Format(BS.PropertyItems("NUMMAX").value & "") = "" Then
                    NumMax = "NULL"
                Else
                    NumMax = Format(BS.PropertyItems("NUMMAX").value)
                End If
                If Format(BS.PropertyItems("NUMMIN").value & "") = "" Then
                    NumMin = "NULL"
                Else
                    NumMin = Format(BS.PropertyItems("NUMMIN").value)
                End If
                If NumMin <> "NULL" And NumMax <> "NULL" Then
                    If Val(NumMin) >= Val(NumMax) Then
                        MsgBox "The numeric range you have entered is invalid.  Please check and try again.", vbInformation + vbOKOnly, "Save Rule"
                        Exit Sub
                    End If
                End If
                If Format(BS.PropertyItems("PICKLIST").value & "") = "" Then
                    PL = "NULL"
                Else
                    PL = Format(BS.PropertyItems("PICKLIST").value)
                End If
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Type='DEN',"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Text='" & BS.PropertyItems("PROMPT").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "DataEntry_Type='" & BS.PropertyItems("TYPE").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "DataEntry_Upper_Val=" & NumMax & ","
                tempStr1 = tempStr1 + vbCrLf + "DataEntry_Lower_Val=" & NumMin & ","
                tempStr1 = tempStr1 + vbCrLf + "PickList_Index=" & PL & ","
                tempStr1 = tempStr1 + vbCrLf + "Save_As_Type='" & BS.PropertyItems("SAVETYPE").value & "',"
                tempStr1 = tempStr1 + vbCrLf + "Save_As_String='" & BS.PropertyItems("SAVEHEADER").value & "'"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "MCD"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Type='MCD'"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "EIF"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Type='EIF'"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "EIM"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Type='EIM'"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "CNL"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Type='CNL'"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "HLP"
                tempStr1 = tempStr1 + vbCrLf + "Prompt_Type='HLP',"
                tempStr1 = tempStr1 + vbCrLf + "Dialog_Text='" & BS.PropertyItems("HELPTEXT").value & "'"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
        End Select
    Else
        RS.Open "Select Max(Prompt_Index) 'MPI' From Request_Prompt", ICECon, adOpenKeyset, adLockReadOnly
        If Format(RS!MPI) = "" Then
            NPI = 1
        Else
            NPI = RS!MPI + 1
        End If
        RS.Close
        tempStr1 = "Insert Into Request_Prompt (Prompt_Index,Prompt_Desc,Prompt_Type,"
        tempstr2 = ") Values (" & Format(NPI) & ",'" & BS.PropertyItems("NAME").value & "',"
        Select Case BS.PropertyItems("STYLE").value
            Case "QUE"
                tempStr1 = tempStr1 + "Prompt_Text,Yes_Text,No_Text,Yes_Action_Type,No_Action_Type,Cancel_Text,Save_As_Type,Save_As_String"
                If Format(BS.PropertyItems("YESACTION").value & "") = "" Then BS.PropertyItems("YESACTION").value = "NULL"
                If Format(BS.PropertyItems("NOACTION").value & "") = "" Then BS.PropertyItems("NOACTION").value = "NULL"
                tempstr2 = tempstr2 + "'QUE','" + BS.PropertyItems("QUESTION").value + "','" + BS.PropertyItems("YESTEXT").value + "','" + BS.PropertyItems("NOTEXT").value + "'," + Format(BS.PropertyItems("YESACTION").value & "") + "," + Format(BS.PropertyItems("NOACTION").value & "") + ",'" + BS.PropertyItems("CANCELTEXT").value + "','" + BS.PropertyItems("SAVETYPE").value + "','" + BS.PropertyItems("SAVEHEADER").value + "')"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "DEN"
                tempStr1 = tempStr1 + "Prompt_Text,DataEntry_Type,DataEntry_Upper_Val,DataEntry_Lower_Val,PickList_Index,Save_As_Type,Save_As_String"
                If Format(BS.PropertyItems("NUMMAX").value & "") = "" Then
                    NumMax = "NULL"
                Else
                    NumMax = Format(BS.PropertyItems("NUMMAX").value)
                End If
                If Format(BS.PropertyItems("NUMMIN").value & "") = "" Then
                    NumMin = "NULL"
                Else
                    NumMin = Format(BS.PropertyItems("NUMMIN").value)
                End If
                If Format(BS.PropertyItems("PICKLIST").value & "") = "" Then
                    PL = "NULL"
                Else
                    PL = Format(BS.PropertyItems("PICKLIST").value)
                End If
                tempstr2 = tempstr2 + "'DEN','" + BS.PropertyItems("PROMPT").value + "','" + BS.PropertyItems("TYPE").value + "'," + NumMax + "," + NumMin + "," + PL + ",'" + BS.PropertyItems("SAVETYPE").value + "','" + BS.PropertyItems("SAVEHEADER").value + "')"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "MCD"
                tempStr1 = Left(tempStr1, Len(tempStr1) - 1)
                tempstr2 = tempstr2 + "'MCD')"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "EIF"
                tempStr1 = Left(tempStr1, Len(tempStr1) - 1)
                tempstr2 = tempstr2 + "'EIF')"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "EIM"
                tempStr1 = Left(tempStr1, Len(tempStr1) - 1)
                tempstr2 = tempstr2 + "'EIM')"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "CNL"
                tempStr1 = Left(tempStr1, Len(tempStr1) - 1)
                tempstr2 = tempstr2 + "'CNL')"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
            Case "HLP"
                tempStr1 = tempStr1 + "Dialog_Text"
                tempstr2 = tempstr2 + "'HLP','" + BS.PropertyItems("HELPTEXT").value + "')"
                TempStr = tempStr1 + tempstr2
                Debug.Print TempStr
                ICECon.Execute TempStr
        End Select
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If testId > -1 Then
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open "Select * From Request_Prompt Where Prompt_Index=" & testId, ICECon, adOpenKeyset, adLockReadOnly
        Select Case RS!Prompt_Type
            Case "QUE"
                SetupBS "QUE"
                SetupBSQuestion
                SetupBS2
                BS.PropertyItems("QUESTION").value = RS!Prompt_Text
                BS.PropertyItems("YESTEXT").value = RS!Yes_Text
                BS.PropertyItems("NOTEXT").value = RS!No_Text
                BS.PropertyItems("YESACTION").value = Format(RS!Yes_Action_Type)
                BS.PropertyItems("NOACTION").value = Format(RS!No_Action_Type)
                BS.PropertyItems("CANCELTEXT").value = RS!Cancel_Text
                BS.PropertyItems("SAVETYPE").value = Format(RS!Save_As_Type)
                BS.PropertyItems("SAVEHEADER").value = RS!Save_As_String
            Case "DEN"
                SetupBS "DEN"
                SetupBSData
                SetupBS2
                BS.PropertyItems("PROMPT").value = RS!Prompt_Text
                BS.PropertyItems("TYPE").value = Format(RS!DataEntry_Type)
                BS.PropertyItems("NUMMAX").value = Format(RS!DataEntry_Upper_Val)
                BS.PropertyItems("NUMMIN").value = Format(RS!DataEntry_Lower_Val)
                BS.PropertyItems("PICKLIST").value = Format(RS!Picklist_Index)
                If BS.PropertyItems("TYPE").value = "P" Then
                    BS.PropertyItems("PICKLIST").Enabled = True
                ElseIf BS.PropertyItems("TYPE").value = "N" Then
                    BS.PropertyItems("NUMMAX").Enabled = True
                    BS.PropertyItems("NUMMIN").Enabled = True
                End If
                BS.PropertyItems("SAVETYPE").value = Format(RS!Save_As_Type)
                BS.PropertyItems("SAVEHEADER").value = RS!Save_As_String
            Case "MCD"
                SetupBS "MCD"
            Case "EIM"
                SetupBS "EIM"
            Case "EIF"
                SetupBS "EIF"
            Case "CNL"
                SetupBS "CNL"
            Case "HLP"
                SetupBS "HLP"
                SetupBSHLP
                BS.PropertyItems("HELPTEXT").value = RS!Dialog_Text
        End Select
        BS.PropertyItems("NAME").value = RS!Prompt_Desc
        RS.Close
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        MsgBox "The character ' is not permitted, please use the ` character instead"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    SetupBS (0)
End Sub

Private Sub SetupBS(value)
    BS.PropertyItems.Clear
    BS.PropertyItems.Add "NAME", "Rule Name", plpsString, , "Name of this rule"
    BS.PropertyItems("NAME").Max = 50
    BS.PropertyItems.Add "STYLE", "Rule Style", plpsList, value, "Type of rule"
    BS.PropertyItems("STYLE").ListItems.Add "Question", "QUE"
    BS.PropertyItems("STYLE").ListItems.Add "Data Entry", "DEN"
    BS.PropertyItems("STYLE").ListItems.Add "Mandatory Clinical Details", "MCD"
    BS.PropertyItems("STYLE").ListItems.Add "Exclude If Male", "EIM"
    BS.PropertyItems("STYLE").ListItems.Add "Exclude If Female", "EIF"
    BS.PropertyItems("STYLE").ListItems.Add "Cancel Test", "CNL"
    BS.PropertyItems("STYLE").ListItems.Add "Help Dialog", "HLP"
End Sub
Private Sub SetupBS2()
    BS.PropertyItems.Add "SAVETYPE", "Save Data As", plpsList, , "Select how the data is stored.  Clinical Detail is encoded in the PDF417 barcode, Test Information is not"
    BS.PropertyItems("SAVETYPE").ListItems.Add "Clinical Detail", "CD"
    BS.PropertyItems("SAVETYPE").ListItems.Add "Test Information", "TI"
    BS.PropertyItems.Add "SAVEHEADER", "Save As Header", plpsString, , "Enter the tag you wish to use to identify this data"
    BS.PropertyItems("SAVEHEADER").Max = 25
End Sub
Private Sub SetupBSQuestion()
    BS.PropertyItems.Add "QUESTION", "Question to Ask", plpsString, , "Enter the question you wish to ask the user on selection of the associated request"
    BS.PropertyItems("QUESTION").Max = 140
    BS.PropertyItems.Add "YESTEXT", "Yes Value", plpsString, "Yes", "Enter the text you wish to save when the user clicks the Yes button"
    BS.PropertyItems("YESTEXT").Max = 25
    BS.PropertyItems.Add "NOTEXT", "No Value", plpsString, "No", "Enter the text you wish to save when the user clicks the No button"
    BS.PropertyItems("NOTEXT").Max = 25
    BS.PropertyItems.Add "YESACTION", "Yes Action", plpsList, , "Select the action you wish to take on clicking the Yes button"
    BS.PropertyItems.Add "NOACTION", "No Action", plpsList, , "Select the action you wish to take on clicking the No button"
    If frmBehaviour.vbalGrid1.Rows > 0 Then
        For i = 1 To frmBehaviour.vbalGrid1.Rows
            BS.PropertyItems("YESACTION").ListItems.Add frmBehaviour.vbalGrid1.Cell(i, 1).Text, frmBehaviour.vbalGrid1.Cell(i, 2).Text
            BS.PropertyItems("NOACTION").ListItems.Add frmBehaviour.vbalGrid1.Cell(i, 1).Text, frmBehaviour.vbalGrid1.Cell(i, 2).Text
        Next i
    End If
    BS.PropertyItems.Add "CANCELTEXT", "Cancel Text", plpsString, , "If selected action cancels the test, message to display on the screen to inform user"
End Sub

Private Sub SetupBSData()
    BS.PropertyItems.Add "PROMPT", "User Prompt", plpsString, , "Enter the prompt you wish to display for data entry"
    BS.PropertyItems("PROMPT").Max = 140
    BS.PropertyItems.Add "TYPE", "Type", plpsList, , "Select the type of data you wish the user to enter"
    BS.PropertyItems("TYPE").ListItems.Add "Free Text", "FT"
    BS.PropertyItems("TYPE").ListItems.Add "Numeric", "N"
    BS.PropertyItems("TYPE").ListItems.Add "Date", "D"
    BS.PropertyItems("TYPE").ListItems.Add "Time", "T"
    BS.PropertyItems("TYPE").ListItems.Add "Date and Time", "DT"
    BS.PropertyItems("TYPE").ListItems.Add "Picklist", "P"
    BS.PropertyItems.Add "NUMMAX", "Numeric Upper Limit", plpsNumber, , "Enter the maximum numeric value a user can enter"
    BS.PropertyItems("NUMMAX").Enabled = False
    BS.PropertyItems.Add "NUMMIN", "Numeric Lower Limit", plpsNumber, , "Enter the minimum numeric value a user can enter"
    BS.PropertyItems("NUMMIN").Enabled = False
    BS.PropertyItems.Add "PICKLIST", "Picklist", plpsList, , "Select the picklist you wish the user to select from"
    BS.PropertyItems("PICKLIST").Enabled = False
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select * From Request_Picklist Order by Picklist_Name", ICECon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            BS.PropertyItems("PICKLIST").ListItems.Add RS!PickList_Name, RS!Picklist_Index
            RS.MoveNext
        Loop
    End If
    RS.Close
End Sub

Private Sub SetupBSHLP()
    BS.PropertyItems.Add "HELPTEXT", "Help Text", plpsString, , "Enter the help you wish to appear in a popup dialog box"
    BS.PropertyItems("HELPTEXT").Max = 255
    BS.PropertyItems("HELPTEXT").MultiLine = True
End Sub
