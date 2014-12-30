Attribute VB_Name = "frmMainCode"
Option Explicit

Dim DragNode As Object
Dim InDrag As Boolean
Public CfgPanelHelp As String
Public RequeueStr As String
Public itemId As String
Public tvNode As String
'Private tPList As PropertiesList
Dim Activity() As String
Dim ActivityDirs() As String
Dim ActivityDirPattern() As String
Dim Reports As AHSLReporting.Reports
Dim Report As AHSLReporting.Report
Dim ReportList As AHSLReporting.ReportList
Private Type RECT
 Left   As Long
 Top    As Long
 Right  As Long
 Bottom As Long
End Type
Private nOrigin As String

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

'Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim mfX As Single
Dim mfY As Single
Dim moNode As MSComctlLib.Node
Dim m_iScrollDir As Integer 'Which way to scroll
Dim mbFlag As Boolean
Dim CfgProgID As String
Dim CfgCfgID As String
Dim CfgWardID As String
Dim CfgUserID As String
Dim ShowRequesting As Boolean
Dim ShowUsers As Boolean
Dim ShowEDI As Boolean
Dim ShowAudit As Boolean
Dim ShowConnections As Boolean
Dim ShowHelp As Boolean
Dim DefaultItem As String

Private Sub Display_DirActivity()
   On Error GoTo ProcEH
   Dim i As Integer
   Dim FL As Integer
   
   eClass.FurtherInfo = "Displaying System Activity directory(ies)"
   FL = 0
   For i = 0 To UBound(ActivityDirs)
      If (ActivityDirs(i) <> "") And (FL <= FLActivity.UBound) Then
         FLlbl(FL).Caption = Activity(i)
         FLActivity(FL).Path = ActivityDirs(i)
         FLActivity(FL).Pattern = ActivityDirPattern(i)
         If i <= FLlbl.UBound Then
            FLlbl(FL).Visible = True
         End If
         FLActivity(FL).Visible = True
         FL = FL + 1
      End If
   Next i
   Timer2.Enabled = True
   Exit Sub
   
ProcEH:
   eClass.CurrentProcedure = "frmMain.Display_DirActivity"
   eClass.Add Err.Number, Err.Description, Err.Source
   
End Sub

Private Sub Check1_Click()
    Command7.Enabled = True
    Command8.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = False
    TreeView1.Enabled = False
    SSListBar1.Enabled = False
    OrgList.Enabled = False
End Sub

Private Sub cmdSrchCancel_Click()

  fView.Show Fra_HELP, ""

End Sub

Private Sub CmdSrchOk_Click()

   On Local Error GoTo ProcEH
   Dim RS As New ADODB.Recordset
   Dim gpCode As String
   
   If optLogSearch(0).value = True Then
'     Search by date
      objTView.LoadLogs "Dates", dtPFrom.value, dtPTo.value
   ElseIf optLogSearch(1).value Then
'     Search by practice
      RS.Open "SELECT DISTINCT EDI_Local_Key1 FROM EDI_Recipient_Individuals WHERE EDI_Org_NatCode = '" & _
                      ComboSrchPractice.Text & "'", ICECon, adOpenKeyset, adLockPessimistic
      gpCode = RS("EDI_Local_Key1")
      RS.Close
      Set RS = Nothing
      objTView.LoadLogs "Practice", OrgList.Text & " " & gpCode
   ElseIf optLogSearch(2).value Then
'     Search by Patient name
      objTView.LoadPatient txtSrchSurname.Text, txtSrchForename.Text, 0
   ElseIf optLogSearch(3).value Then
'     Search by Hospital/NHS Number
      objTView.LoadPatient txtSrchNHS.Text, , 1
   ElseIf optLogSearch(4).value Then
'     Search by report id
      objTView.LoadReports txtLogSearchLab.Text
   End If
'   fraSrchLog.Visible = False
   frmMain.LogText.Visible = frmMain.TreeView1.Visible
   fView.Show Fra_LOGVIEW, ""
   Exit Sub
   
ProcEH:
   eClass.Show
   
End Sub

Private Sub Command10_Click()

   Dim strArray() As String
   
'    If TreeView1.SelectedItem.Children = 0 Then Exit Sub
    Load frmIncExcRfxDelete
    frmIncExcRfxDelete.vbalGrid1.Clear True
    frmIncExcRfxDelete.vbalGrid1.AddColumn "COL1", "Test"
    frmIncExcRfxDelete.vbalGrid1.AddColumn "COL2", "Index", , , , False
'    frmIncExcRfxDelete.vbalGrid1.Rows = TreeView1.SelectedItem.Children
'    frmIncExcRfxDelete.vbalGrid1.CellDetails 1, 1, TreeView1.SelectedItem.Child.Text
'    frmIncExcRfxDelete.vbalGrid1.CellDetails 1, 2, Mid$(TreeView1.SelectedItem.Child.Key, InStr(1, TreeView1.SelectedItem.Child.Key, "-") + 1, Len(TreeView1.SelectedItem.Child.Key) - InStr(1, TreeView1.SelectedItem.Child.Key, "-") + 1)
    Dim TVN As Node
'    Set TVN = TreeView1.SelectedItem.Child
'    For i = 1 To TreeView1.SelectedItem.Children - 1
'        frmIncExcRfxDelete.vbalGrid1.CellDetails i + 1, 1, TVN.Next.Text
'        strarray = Split(objTView.NodeKey(TVN.Next.Key), ":")
'        frmIncExcRfxDelete.vbalGrid1.CellDetails i + 1, 2, strarray(2)
'        frmIncExcRfxDelete.vbalGrid1.CellDetails i + 1, 2, Mid$(TVN.Next.Key, InStr(1, TVN.Next.Key, "-") + 1, Len(TVN.Next.Key) - InStr(1, TVN.Next.Key, "-") + 1)
'        Set TVN = TVN.Next
'    Next i
'    frmIncExcRfxDelete.vbalGrid1.AutoWidthColumn "COL1"
    frmIncExcRfxDelete.Show 1
End Sub

Private Sub Command11_Click()
    If PR1.PropertyItems("WARD").value = "" Then
        MsgBox "You must specify a ward to allocate this profile to", vbInformation + vbOKOnly, "Save Profile"
        Exit Sub
    End If
    Command11.Enabled = False
    Command12.Enabled = False
    Command13.Enabled = True
    Command14.Enabled = True
    SSListBar1.Enabled = True
    TreeView1.Enabled = True
    OrgList.Enabled = True
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    If objTView.NodeKey(TreeView1.SelectedItem.Key) = "New" Then
        RS.Open "Select Max(Profile_Index) 'MPI' From Request_Profiles", ICECon, adOpenKeyset, adLockReadOnly
        If Format(RS!MPI) & "" = "" Then
            NPI = 1
        Else
            NPI = RS!MPI + 1
        End If
        RS.Close
        TempStr1 = "Insert Into Request_Profiles (Profile_Index,Enabled,Date_Added,Profile_Location_Code,Profile_Position,Profile_Caption,Profile_Colour,Profile_Help,Profile_Help_Backcolour,Profile_TextString) Values ("
        If PR1.PropertyItems("ENABLED").value = True Then
            en = 1
        Else
            en = 0
        End If
        TempStr2 = Format(NPI) & "," & Format(en) & ",'" & Format(Date, "DD MMM YYYY") & "','" & PR1.PropertyItems("WARD").value & "',"
        If Format(PR1.PropertyItems("POSITION").value) <> "" Then
            TempStr2 = TempStr2 & Format(PR1.PropertyItems("POSITION").value)
        Else
            TempStr2 = TempStr2 & "NULL"
        End If
        TempStr2 = TempStr2 & ",'" & PR1.PropertyItems("CAPTION").value & "',"
        If PR1.PropertyItems("PROFILE_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + Left(PR1.PropertyItems("PROFILE_COLOUR").Tag, InStr(1, PR1.PropertyItems("PROFILE_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        TempStr2 = TempStr2 + "'" + PR1.PropertyItems("HELP").value & "',"
        If PR1.PropertyItems("HELP_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + Left(PR1.PropertyItems("HELP_COLOUR").Tag, InStr(1, PR1.PropertyItems("HELP_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        TempStr2 = TempStr2 + "'" + PR1.PropertyItems("TEXT_STRING").value + "')"
        TempStr = TempStr1 + TempStr2
        Debug.Print TempStr
        ICECon.Execute TempStr
    Else
        NPI = Mid$(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
        TempStr1 = "Update Request_Profiles Set "
        If PR1.PropertyItems("ENABLED").value = True Then
            en = 1
        Else
            en = 0
        End If
        TempStr2 = "Enabled=" & Format(en) + ","
        TempStr2 = TempStr2 + "Profile_Location_Code='" & PR1.PropertyItems("WARD").value + "',"
        If Format(PR1.PropertyItems("POSITION").value) <> "" Then TempStr2 = TempStr2 + "Profile_Position=" & Format(PR1.PropertyItems("POSITION").value) & ","
        TempStr2 = TempStr2 + "Profile_Caption='" & Format(PR1.PropertyItems("CAPTION").value) & "',"
        If PR1.PropertyItems("PROFILE_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + "Profile_Colour=" + Left(PR1.PropertyItems("PROFILE_COLOUR").Tag, InStr(1, PR1.PropertyItems("PROFILE_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "Profile_Colour=NULL,"
        End If
        TempStr2 = TempStr2 + "Profile_Help='" & PR1.PropertyItems("HELP").value & "',"
        If PR1.PropertyItems("HELP_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + "Profile_Help_Backcolour=" + Left(PR1.PropertyItems("HELP_COLOUR").Tag, InStr(1, PR1.PropertyItems("HELP_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "Profile_Help_Backcolour=NULL,"
        End If
        TempStr2 = TempStr2 + "Profile_TextString='" & PR1.PropertyItems("TEXT_STRING").value & "' "
        TempStr2 = TempStr2 + "Where Profile_Index=" & NPI
        TempStr = TempStr1 + TempStr2
        Debug.Print TempStr
        ICECon.Execute TempStr
    End If
    fView.Show Fra_PROFILE
'    objTView.LoadOrgProfiles (OrgList.Text)
'    Profile.Visible = False
End Sub

Private Sub Command12_Click()
    Command11.Enabled = False
    Command12.Enabled = False
    Command13.Enabled = True
    Command14.Enabled = True
    SSListBar1.Enabled = True
    TreeView1.Enabled = True
    OrgList.Enabled = True
End Sub

Private Sub Command13_Click()
    frmIncExcRfx.Show 1
End Sub

Private Sub Command14_Click()
    If TreeView1.SelectedItem.Children = 0 Then Exit Sub
    Load frmIncExcRfxDelete
    frmIncExcRfxDelete.vbalGrid1.Clear True
    frmIncExcRfxDelete.vbalGrid1.AddColumn "COL1", "Test"
    frmIncExcRfxDelete.vbalGrid1.AddColumn "COL2", "Index", , , , False
    frmIncExcRfxDelete.vbalGrid1.Rows = TreeView1.SelectedItem.Children
    frmIncExcRfxDelete.vbalGrid1.CellDetails 1, 1, TreeView1.SelectedItem.Child.Text
    frmIncExcRfxDelete.vbalGrid1.CellDetails 1, 2, Mid$(TreeView1.SelectedItem.Child.Key, InStr(1, TreeView1.SelectedItem.Child.Key, "-") + 1, Len(TreeView1.SelectedItem.Child.Key) - InStr(1, TreeView1.SelectedItem.Child.Key, "-") + 1)
    Dim TVN As Node
    Set TVN = TreeView1.SelectedItem.Child
    For i = 1 To TreeView1.SelectedItem.Children - 1
        frmIncExcRfxDelete.vbalGrid1.CellDetails i + 1, 1, TVN.Next.Text
        frmIncExcRfxDelete.vbalGrid1.CellDetails i + 1, 2, Mid$(TVN.Next.Key, InStr(1, TVN.Next.Key, "-") + 1, Len(TVN.Next.Key) - InStr(1, TVN.Next.Key, "-") + 1)
        Set TVN = TVN.Next
    Next i
    frmIncExcRfxDelete.vbalGrid1.AutoWidthColumn "COL1"
    frmIncExcRfxDelete.Show 1
End Sub

Private Sub Command15_Click()
    'Delete Profile
End Sub

Private Sub Command16_Click()
    'Delete Test
    Load frmWait
    frmWait.Label1.Caption = "Please wait whilst test dependencies are calculated..."
    frmWait.Show
    frmWait.Refresh
    
End Sub

Private Sub Command17_Click()
    'Delete Picklist
End Sub

Private Sub Command18_Click()
    'Save Changes to Configuration Option
    If CfgName.Caption = "New" Then
        If ConfigList.PropertyItems("ProgramID").value = "" Or ConfigList.PropertyItems("Name").value = "" Or ConfigList.PropertyItems("Type").value = "" Or ConfigList.PropertyItems("Desc").value = "" Then
            MsgBox "You must supply values for all four fields", vbInformation + vbOKOnly, "New Configuration Option"
            Exit Sub
        End If
        TempStr = "Insert Into Configuration (Organisation,ProgramID,CfgID,CfgType,CfgNotes) Values ('" & OrgList.Text & "','" & ConfigList.PropertyItems("ProgramID").value & "','" & ConfigList.PropertyItems("Name").value & "'," & ConfigList.PropertyItems("Type").value & ",'" & ConfigList.PropertyItems("Desc").value & "')"
        ICECon.Execute TempStr
        objTView.LoadConfiguration OrgList.Text
        Frame2.Visible = False
    Else
        'MsgBox "Saving Configuration Option: " & CfgCfgID & ", " & CfgWardID & CfgUserID & ", " & CfgProgID
        CfgValue = ConfigList.PropertyItems("Value").value
        If CfgValue = "True" Then CfgValue = "1"
        If CfgValue = "False" Then CfgValue = "0"
        If CfgWardID = "" And CfgUserID = "" Then
            TempStr = "Update Configuration Set CfgValue='" & CfgValue & "' Where Organisation='" & OrgList.Text & "' And ProgramID='" & CfgProgID & "' And CfgID='" & CfgCfgID & "' And Location Is Null and Username Is Null"
        ElseIf CfgWardID <> "" Then
            TempStr = "Update Configuration Set CfgValue='" & CfgValue & "' Where Organisation='" & OrgList.Text & "' And ProgramID='" & CfgProgID & "' And CfgID='" & CfgCfgID & "' And Location='" & CfgWardID & "'"
        ElseIf CfgUserID <> "" Then
            TempStr = "Update Configuration Set CfgValue='" & CfgValue & "' Where Organisation='" & OrgList.Text & "' And ProgramID='" & CfgProgID & "' And CfgID='" & CfgCfgID & "' And Username='" & CfgUserID & "'"
        End If
        ICECon.Execute TempStr
        Frame2.Visible = False
    End If
End Sub

Private Sub Command19_Click()
    'Cancel Changes to Configuration Option
    Frame2.Visible = False
End Sub

Private Sub Command21_Click()
    'Add New User Overrider
    Load frmOverride
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select Distinct User_Name From Service_User Order By User_Name", ICECon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            frmOverride.Combo1.AddItem RS!User_Name
            RS.MoveNext
        Loop
    End If
    RS.Close
    frmOverride.Label1.Caption = "Please select the user for which you wish to override the default configuration setting from the list below"
    frmOverride.Show 1
    If PickedOverride <> "" Then
        RS.Open "Select CfgType,CfgNotes,CfgValue From Configuration Where Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "' And Location Is Null And UserName Is Null", ICECon, adOpenKeyset, adLockReadOnly
        TempStr = "Insert Into Configuration (Organisation,Username,ProgramID,CfgID,CfgType,CfgNotes,CfgValue) Values ('" & OrgList.Text & "','" & PickedOverride & "','" & CfgProgID & "','" & CfgCfgID & "'," & Format(RS!CfgType) & ",'" & RS!CfgNotes & "','" & RS!CfgValue & "')"
        RS.Close
        ICECon.Execute TempStr
        objTView.LoadConfiguration OrgList.Text
    End If
End Sub

Private Sub Command20_Click()
    'Mew Location Overrider
    Load frmOverride
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select Distinct Local_Location_Code From Location Order By Local_Location_Code", ICECon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            frmOverride.Combo1.AddItem RS!Local_Location_Code
            RS.MoveNext
        Loop
    End If
    RS.Close
    frmOverride.Label1.Caption = "Please select the location for which you wish to override the default configuration setting from the list below"
    frmOverride.Show 1
    If PickedOverride <> "" Then
        RS.Open "Select CfgType,CfgNotes,CfgValue From Configuration Where Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "' And Location Is Null And UserName Is Null", ICECon, adOpenKeyset, adLockReadOnly
        TempStr = "Insert Into Configuration (Organisation,Location,ProgramID,CfgID,CfgType,CfgNotes,CfgValue) Values ('" & OrgList.Text & "','" & PickedOverride & "','" & CfgProgID & "','" & CfgCfgID & "'," & Format(RS!CfgType) & ",'" & RS!CfgNotes & "','" & RS!CfgValue & "')"
        RS.Close
        ICECon.Execute TempStr
        objTView.LoadConfiguration OrgList.Text
    End If
End Sub

Private Sub Command22_Click()
    'Delete Configuration Option
    If MsgBox("Are you sure you wish to delete this configuration option as doing so could affect program operation?", vbQuestion + vbYesNo, "Delete Configuration Option") = vbYes Then
        TempStr = "Delete From Configuration Where ProgramID='" & CfgProgID & "' And Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "'"
        ICECon.Execute TempStr
        objTView.LoadConfiguration (OrgList.Text)
        Command22.Visible = False
    End If
End Sub

Private Sub Command23_Click()
    'Delete Location Override
    Load frmOverride
    frmOverride.Label1.Caption = "Please select the location you wish to revert to using the default setting from the list below"
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select Distinct Location From Configuration Where Location Is Not Null And Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "'", ICECon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            frmOverride.Combo1.AddItem RS!Location
            RS.MoveNext
        Loop
        frmOverride.Show 1
        If PickedOverride <> "" Then
            TempStr = "Delete From Configuration Where Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "' And Location='" & PickedOverride & "'"
            ICECon.Execute TempStr
            objTView.LoadConfiguration (OrgList.Text)
        End If
    End If
    RS.Close
End Sub

Private Sub Command24_Click()
    'Delete User Override
    Load frmOverride
    frmOverride.Label1.Caption = "Please select the location you wish to revert to using the default setting from the list below"
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select Distinct Username From Configuration Where Username Is Not Null And Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "'", ICECon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            frmOverride.Combo1.AddItem RS!Username
            RS.MoveNext
        Loop
        frmOverride.Show 1
        If PickedOverride <> "" Then
            TempStr = "Delete From Configuration Where Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "' And Username='" & PickedOverride & "'"
            ICECon.Execute TempStr
            objTView.LoadConfiguration (OrgList.Text)
        End If
    End If
    RS.Close
End Sub

Private Sub Command4_Click()
    If TreeView1.SelectedItem.Children = 0 Then
        Exit Sub
    Else
        Load frmBehaviourDelete
        frmBehaviourDelete.vbalGrid1.Clear True
        frmBehaviourDelete.vbalGrid1.AddColumn "COL1", "Rule", ecgHdrTextALignCentre
        frmBehaviourDelete.vbalGrid1.AddColumn "COL2", "Key", ecgHdrTextALignLeft, , , False
        frmBehaviourDelete.vbalGrid1.Rows = TreeView1.SelectedItem.Children
        frmBehaviourDelete.vbalGrid1.CellDetails 1, 1, TreeView1.SelectedItem.Child.Text
        frmBehaviourDelete.vbalGrid1.CellDetails 1, 2, TreeView1.SelectedItem.Child.Index
        Dim TVN As Node
        Set TVN = TreeView1.SelectedItem.Child
        For i = 1 To TreeView1.SelectedItem.Children - 1
            frmBehaviourDelete.vbalGrid1.CellDetails i + 1, 1, TreeView1.Nodes(TVN.Index).Next.Text
            frmBehaviourDelete.vbalGrid1.CellDetails i + 1, 2, TreeView1.Nodes(TVN.Index).Next.Index
            Set TVN = TreeView1.Nodes(TVN.Index).Next
        Next i
        frmBehaviourDelete.Show 1
    End If
End Sub

Private Sub Command5_Click()
    TempStr = InputBox("Please enter the picklist value you wish to add to the '" & TreeView1.SelectedItem.Text & "' picklist:", "Add Picklist Value")
    If Trim(TempStr) = "" Then Exit Sub
    If Len(Trim(TempStr)) > 50 Then
        MsgBox "The picklist value you have entered is too long.  Picklist values can be upto 50 characters.", vbInformation + vbOKOnly, "Add Picklist Value"
        Exit Sub
    End If
    If TreeView1.SelectedItem.Children > 0 Then
        Dim TVN As Node
        Set TVN = TreeView1.SelectedItem.Child
        If UCase(TempStr) = UCase(TVN.Text) Then
            MsgBox "This value already exists on the selected picklist and cannot be added again"
            Exit Sub
        End If
        For i = 1 To TreeView1.SelectedItem.Children - 1
            Set TVN = TVN.Next
            If TempStr = TVN.Text Then
                MsgBox "This value already exists on the selected picklist and cannot be added again"
                Exit Sub
            End If
        Next i
    End If
    NPI = Mid$(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
    ICECon.Execute "Insert Into Request_Picklist_Data (Picklist_Index,Picklist_Value) Values (" & NPI & ",'" & TempStr & "')"
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    If TreeView1.SelectedItem.Children > 0 Then
        For i = 1 To TreeView1.SelectedItem.Children
            TreeView1.Nodes.Remove TreeView1.SelectedItem.Child.Index
        Next i
    End If
    RS.Open "Select * From Request_Picklist_Data Where Picklist_Index=" & NPI & " Order by Picklist_Value", ICECon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            TreeView1.Nodes.Add TreeView1.SelectedItem, tvwChild, "D" + Format(RS!Picklist_Index) + "-" + RS!Picklist_Value, RS!Picklist_Value, 1, 1
            RS.MoveNext
        Loop
    End If
    RS.Close
End Sub

Private Sub Command6_Click()
    Load frmPicklistDelete
    If TreeView1.SelectedItem.Children > 0 Then
        Dim TVN As Node
        Set TVN = TreeView1.SelectedItem.Child
        frmPicklistDelete.vbalGrid1.AddColumn "COL1"
        frmPicklistDelete.vbalGrid1.Rows = TreeView1.SelectedItem.Children
        frmPicklistDelete.vbalGrid1.CellDetails 1, 1, TVN.Text
        NoValues = 1
        For i = 1 To TreeView1.SelectedItem.Children - 1
            NoValues = NoValues + 1
            Set TVN = TVN.Next
            frmPicklistDelete.vbalGrid1.CellDetails NoValues, 1, TVN.Text
        Next i
        frmPicklistDelete.vbalGrid1.AutoWidthColumn "COL1"
        NPI = Mid$(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
        frmPicklistDelete.Label2.Caption = NPI
        frmPicklistDelete.Show 1
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        If TreeView1.SelectedItem.Children > 0 Then
            For i = 1 To TreeView1.SelectedItem.Children
                TreeView1.Nodes.Remove TreeView1.SelectedItem.Child.Index
            Next i
        End If
        RS.Open "Select * From Request_Picklist_Data Where Picklist_Index=" & NPI & " Order by Picklist_Value", ICECon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            Do While Not RS.EOF
                TreeView1.Nodes.Add TreeView1.SelectedItem, tvwChild, "D" + Format(RS!Picklist_Index) + "-" + RS!Picklist_Value, RS!Picklist_Value, 1, 1
                RS.MoveNext
            Loop
        End If
        RS.Close
    End If
End Sub

Private Sub Command7_Click()
    If objTView.NodeKey(TreeView1.SelectedItem.Key) = "New" Then
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open "Select Max(Picklist_Index) 'MPI' from Request_Picklist", ICECon, adOpenKeyset, adLockReadOnly
        If Format(RS!MPI) & "" = "" Then
            NPI = 1
        Else
            NPI = RS!MPI + 1
        End If
        RS.Close
        TempStr = "Insert Into Request_Picklist (Picklist_Index,Picklist_Name,Multichoice) Values (" & Format(NPI) & ",'" & Text1.Text & "',"
        If Check1.value Then
            TempStr = TempStr + "1)"
        Else
            TempStr = TempStr + "0)"
        End If
        ICECon.Execute TempStr
    Else
        NPI = Mid$(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
        TempStr = "Update Request_Picklist Set Picklist_Name='" & Text1.Text & "',Multichoice="
        If Check1.value Then
            TempStr = TempStr + "1 Where Picklist_Index=" & NPI
        Else
            TempStr = TempStr + "0 Where Picklist_Index=" & NPI
        End If
        ICECon.Execute TempStr
    End If
    Command7.Enabled = False
    Command8.Enabled = False
    Command5.Enabled = True
    Command6.Enabled = True
'    TreeView1.Enabled = True
'    SSListBar1.Enabled = True
    OrgList.Enabled = True
'    Picklist.Visible = False
'    objTView.LoadPicklists
End Sub

Private Sub Command8_Click()
    Command7.Enabled = False
    Command8.Enabled = False
    Command5.Enabled = True
    Command6.Enabled = True
    TreeView1.Enabled = True
    SSListBar1.Enabled = True
    OrgList.Enabled = True
End Sub

Private Sub Command9_Click()
    frmIncExcRfx.Show 1
End Sub

Private Sub DeleteEntry_Click()

End Sub

Private Sub ediPr_AfterEdit(PropertyItem As PropertiesListCtl.PropertyItem, NewValue As Variant, Cancel As Boolean)
   If objctrl.AddData(PropertyItem.Key, NewValue) Then
      objctrl.DataChanged = True
   Else
     Cancel = True
   End If
End Sub

Private Sub edipr_PropertyBrowseClick(PropertyItem As PropertiesListCtl.PropertyItem)
   
   Dim RS As New ADODB.Recordset
   
   ediPr.Tag = PropertyItem.Key
   
   Select Case PropertyItem.Key
      Case "EC"
         frmReadCodes.Show 1
         
      Case "LSA"
         If ediPr("SA").value = "" Then
            MsgBox "Please set an Anatomical origin for this sample before adding descriptions", vbInformation, "Anatomical Origin not defined"
            ediPr("SA").Selected = True
         Else
            frmSampleData.Tag = "EDI_Local_Sample_AnatOrigin"
            frmSampleData.txtNatCode.Text = ediPr("SA").value
            frmSampleData.fraDesc.Caption = "Anatomical Origin Description"
            frmSampleData.Show
         End If
      
      Case "LSC"
         If ediPr("SC").value = "" Then
            MsgBox "Please set a Collection Code for this sample before adding descriptions", vbInformation, "Collection Code not defined"
            ediPr("SC").Selected = True
         Else
            frmSampleData.Tag = "EDI_Local_Sample_CollectionTypes"
            frmSampleData.txtNatCode.Text = ediPr("SC").value
            frmSampleData.fraDesc.Caption = "Collection Type Description"
            frmSampleData.Show
         End If
      
      Case "LST"
         If ediPr("ST").value = "" Then
            MsgBox "Please set a Sample Code for this sample before adding descriptions", vbInformation, "Sample Code not defined"
            ediPr("ST").Selected = True
         Else
            frmSampleData.Tag = "EDI_Local_Sample_Types"
            frmSampleData.txtNatCode.Text = ediPr("ST").value
            frmSampleData.fraDesc.Caption = "Sample Description"
            frmSampleData.Show
         End If
         
      Case "NatUOM"
         frmUOMCodes.Show 1
      
      Case "NT"
         frmSampCodes.Show 1
         ediPr("NS").value = frmSampCodes.txtSpec.Text
         Unload frmSampCodes
         ediPr_AfterEdit ediPr("NS"), ediPr("NS").value, False
         
      Case "RC"
         frmReadCodes.Show 1
         
      Case "SA"
         frmSampCodes.dbTable = "CRIR_Sample_AnatOrigin"
         frmSampCodes.NationalCodeField = "Origin_Code"
         frmSampCodes.NationalDescriptionField = "Origin_Text"
         frmSampCodes.ReturnDataTo = "SA"
         frmSampCodes.Show 1
         Unload frmSampCodes
      
      Case "SC"
         frmSampCodes.dbTable = "CRIR_Sample_CollectionType"
         frmSampCodes.NationalCodeField = "Collection_Code"
         frmSampCodes.NationalDescriptionField = "Collection_Text"
         frmSampCodes.ReturnDataTo = "SC"
         frmSampCodes.Show 1
         Unload frmSampCodes
         
      Case "ST"
         frmSampCodes.dbTable = "CRIR_Sample_Type"
         frmSampCodes.NationalCodeField = "Sample_Code"
         frmSampCodes.NationalDescriptionField = "Sample_Text"
         frmSampCodes.ReturnDataTo = "ST"
         frmSampCodes.Show 1
         Unload frmSampCodes
      
      Case Else
'         frmSpecs.ediPrSub.UsePageKeys = (PropertyItem.Key = "SP+MS1")
'         frmSpecs.ediPrSub.ShowPageStrip = (PropertyItem.Key = "SP+MS1")
         frmSpecs.Show 1
         If ediPr.Tag = "SP+MS1" Then
            ediPr_AfterEdit ediPr("SP+MS2"), ediPr("SP+MS2").value, False
         End If
   End Select
   ediPr_AfterEdit PropertyItem, ediPr(PropertyItem.Key).value, False
   
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   If Source.Name = "TreeView1" Then
      Timer1.Enabled = False
   End If
   blnAddSpec = False
End Sub
Private Sub Command1_Click()
    If Trim(TD1.PropertyItems("SCREEN_CAPTION").value) = "" Then
        MsgBox "A screen caption must be entered in order to save the test", vbInformation + vbOKOnly, "Save Test"
        Exit Sub
    End If
    If Format(TD1.PropertyItems("VOL").value) <> "" Then
        If Val(TD1.PropertyItems("VOL").value) > 5000 Then
            MsgBox "The volume required value you have entered is invalid, please check and try again.", vbInformation + vbOKOnly, "Save Test"
            Exit Sub
        End If
    End If
    If Format(TD1.PropertyItems("PAED_VOL").value) <> "" Then
        If Val(TD1.PropertyItems("PAED_VOL").value) > 5000 Then
            MsgBox "The volume required value you have entered is invalid, please check and try again.", vbInformation + vbOKOnly, "Save Test"
            Exit Sub
        End If
    End If
'    If TreeView1.SelectedItem.Key = "T9999" Then
    If objTView.NodeKey(TreeView1.SelectedItem.Key) = "New" Then
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open "Select Max(Test_Index) 'MTI' From Request_Tests", ICECon, adOpenKeyset, adLockReadOnly
        If Format(RS!MTI) & "" = "" Then
            NTI = 1
        Else
            NTI = Val(RS!MTI) + 1
        End If
        If Format(TD1.PropertyItems("DYNAMIC").value) = "" Or TD1.PropertyItems("DYNAMIC").value = False Then
            DYN = 0
        Else
            DYN = 1
        End If
        If Format(TD1.PropertyItems("SENSITIVE").value) = "" Or TD1.PropertyItems("SENSITIVE").value = False Then
            SENS = 0
        Else
            SENS = 1
        End If
        If Format(TD1.PropertyItems("WORKLIST").value) = "" Or TD1.PropertyItems("WORKLIST").value = False Then
            WL = 0
        Else
            WL = 1
        End If
        If Format(TD1.PropertyItems("ENABLED").value) = "" Or TD1.PropertyItems("ENABLED").value = False Then
            en = 0
        Else
            en = 1
        End If
        TempStr1 = "Insert Into Request_Tests (Test_Index,Test_Code,Department,Screen_Panel,Screen_Panel_Page,Screen_Position,Screen_Caption,"
        TempStr1 = TempStr1 + vbCrLf + "Screen_Help,Screen_Colour,Screen_Help_Backcolour,Information,Provider_ID,Read_Code,"
        TempStr1 = TempStr1 + vbCrLf + "Tube_Code,PaedTube_Code,ResHistory_SearchType,ResHistory_SearchString,ResHistory_Message,"
        TempStr1 = TempStr1 + vbCrLf + "Dynamic,Sensitive,Worklist_Enabled,Enabled,Date_Added,Test_Volume,PaedTube_Test_Volume) Values ("
        TempStr2 = Format(NTI) + ",'" + Format(OrgList.Text, "!@@@@@@") + Format(TD1.PropertyItems("TEST_CODE").value) + "','" + Format(TD1.PropertyItems("DEPT").value) + "',"
        If Format(TD1.PropertyItems("SCREEN_PANEL").value) <> "" Then
            TempStr2 = TempStr2 + Format(TD1.PropertyItems("SCREEN_PANEL").value) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        TempStr2 = TempStr2 + "'" + TD1.PropertyItems("PANEL_PAGE").value + "',"
        If Format(TD1.PropertyItems("SCREEN_POSN").value) <> "" Then
            TempStr2 = TempStr2 + Format(TD1.PropertyItems("SCREEN_POSN").value) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        TempStr2 = TempStr2 + "'" + LTrim(Format(TD1.PropertyItems("SCREEN_CAPTION").value)) + "',"
        TempStr2 = TempStr2 + vbCrLf + "'" + Format(TD1.PropertyItems("SCREEN_HELP").value) + "',"
        If TD1.PropertyItems("SCREEN_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + Left(TD1.PropertyItems("SCREEN_COLOUR").Tag, InStr(1, TD1.PropertyItems("SCREEN_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        If TD1.PropertyItems("HELP_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + Left(TD1.PropertyItems("HELP_COLOUR").Tag, InStr(1, TD1.PropertyItems("HELP_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        TempStr2 = TempStr2 + "'" + Format(TD1.PropertyItems("INFO").value) + "',"
        If TD1.PropertyItems("PROV_ID").Tag <> "" Then
            TempStr2 = TempStr2 + Format(TD1.PropertyItems("PROV_ID").Tag) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        TempStr2 = TempStr2 + "'" + Format(TD1.PropertyItems("READ_CODE").value) + "',"
        If TD1.PropertyItems("TUBE_CODE").Tag <> "" Then
            TempStr2 = TempStr2 + vbCrLf + Left(TD1.PropertyItems("TUBE_CODE").Tag, InStr(1, TD1.PropertyItems("TUBE_CODE").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        If TD1.PropertyItems("PAED_TUBE_CODE").Tag <> "" Then
            TempStr2 = TempStr2 + Left(TD1.PropertyItems("PAED_TUBE_CODE").Tag, InStr(1, TD1.PropertyItems("PAED_TUBE_CODE").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        TempStr2 = TempStr2 + "'" + Format(TD1.PropertyItems("HIST_ST").value) + "','" + Format(TD1.PropertyItems("HIST_SS").value) + "','" + Format(TD1.PropertyItems("HIST_MSG").value) + "',"
        TempStr2 = TempStr2 + vbCrLf + Format(DYN) + "," + Format(SENS) + "," + Format(WL) + "," + Format(en) + ",'" + Format(Date, "DD MMM YYYY") + "',"
        If Format(TD1.PropertyItems("VOL").value) <> "" Then
            TempStr2 = TempStr2 + Format(TD1.PropertyItems("VOL").value) + ","
        Else
            TempStr2 = TempStr2 + "NULL,"
        End If
        If Format(TD1.PropertyItems("PAED_VOL").value) <> "" Then
            TempStr2 = TempStr2 + Format(TD1.PropertyItems("PAED_VOL").value) + ")"
        Else
            TempStr2 = TempStr2 + "NULL)"
        End If
        TempStr = TempStr1 + vbCrLf + TempStr2
        Debug.Print TempStr
        ICECon.Execute TempStr
        Dim TVN As Node
        Set TVN = TreeView1.Nodes.Add(, tvwLast, "T" + Format(NTI), LTrim(TD1.PropertyItems("SCREEN_CAPTION").value), 1, 1)
        TreeView1.Nodes.Add TVN, tvwChild, "I" + Format(NTI), "Included Tests", 1, 1
        TreeView1.Nodes.Add TVN, tvwChild, "X" + Format(NTI), "Excluded Tests", 1, 1
        TreeView1.Nodes.Add TVN, tvwChild, "R" + Format(NTI), "Reflex Tests", 1, 1
        TreeView1.Nodes.Add TVN, tvwChild, "B" + Format(NTI), "Rule", 1, 1
        If en = 0 Then TVN.ForeColor = vbRed
        TreeView1.Visible = False
        TreeView1.Sorted = True
        TreeView1.Visible = True
    Else
        NTI = Val(Mid$(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1))
        If Format(TD1.PropertyItems("DYNAMIC").value) = "" Or TD1.PropertyItems("DYNAMIC").value = False Then
            DYN = 0
        Else
            DYN = 1
        End If
        If Format(TD1.PropertyItems("SENSITIVE").value) = "" Or TD1.PropertyItems("SENSITIVE").value = False Then
            SENS = 0
        Else
            SENS = 1
        End If
        If Format(TD1.PropertyItems("WORKLIST").value) = "" Or TD1.PropertyItems("WORKLIST").value = False Then
            WL = 0
        Else
            WL = 1
        End If
        If Format(TD1.PropertyItems("ENABLED").value) = "" Or TD1.PropertyItems("ENABLED").value = False Then
            en = 0
        Else
            en = 1
        End If
        TempStr1 = " Where Test_Index=" + Format(NTI)
        '"Test_Index,Test_Code,Department,Screen_Position,Screen_Caption,"
        'TempStr1 = TempStr1 + vbCrLf + "Screen_Help,Screen_Colour,Screen_Help_Backcolour,Information,Provider_ID,Read_Code,"
        'TempStr1 = TempStr1 + vbCrLf + "Tube_Code,PaedTube_Code,ResHistory_SearchType,ResHistory_SearchString,ResHistory_Message,"
        'TempStr1 = TempStr1 + vbCrLf + "Dynamic,Sensitive,Worklist_Enabled,Enabled,Date_Added,Test_Volume,PaedTube_Test_Volume) Values ("
        TempStr2 = "Update Request_Tests Set Test_Code='" + Format(OrgList.Text, "!@@@@@@") + Format(TD1.PropertyItems("TEST_CODE").value) + "',Department='" + Format(TD1.PropertyItems("DEPT").value) + "',"
        If Format(TD1.PropertyItems("SCREEN_PANEL").value) <> "" Then
            TempStr2 = TempStr2 + "Screen_Panel=" + Format(TD1.PropertyItems("SCREEN_PANEL").value) + ","
        Else
            TempStr2 = TempStr2 + "Screen_Panel=NULL,"
        End If
        TempStr2 = TempStr2 + "Screen_Panel_Page='" + TD1.PropertyItems("PANEL_PAGE").value + "',"
        If Format(TD1.PropertyItems("SCREEN_POSN").value) <> "" Then
            TempStr2 = TempStr2 + "Screen_Position=" + Format(TD1.PropertyItems("SCREEN_POSN").value) + ","
        Else
            TempStr2 = TempStr2 + "Screen_Position=NULL,"
        End If
        TempStr2 = TempStr2 + "Screen_Caption='" + LTrim(Format(TD1.PropertyItems("SCREEN_CAPTION").value)) + "',"
        TempStr2 = TempStr2 + vbCrLf + "Screen_Help='" + Format(TD1.PropertyItems("SCREEN_HELP").value) + "',"
        If TD1.PropertyItems("SCREEN_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + "Screen_Colour=" + Left(TD1.PropertyItems("SCREEN_COLOUR").Tag, InStr(1, TD1.PropertyItems("SCREEN_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "Screen_Colour=NULL,"
        End If
        If TD1.PropertyItems("HELP_COLOUR").Tag <> "" Then
            TempStr2 = TempStr2 + "Screen_Help_Backcolour=" + Left(TD1.PropertyItems("HELP_COLOUR").Tag, InStr(1, TD1.PropertyItems("HELP_COLOUR").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "Screen_Help_Backcolour=NULL,"
        End If
        TempStr2 = TempStr2 + "Information='" + Format(TD1.PropertyItems("INFO").value) + "',"
        If TD1.PropertyItems("PROV_ID").Tag <> "" Then
            TempStr2 = TempStr2 + "Provider_ID=" + Format(TD1.PropertyItems("PROV_ID").Tag) + ","
        Else
            TempStr2 = TempStr2 + "Provider_ID=NULL,"
        End If
        TempStr2 = TempStr2 + "Read_Code='" + Format(TD1.PropertyItems("READ_CODE").value) + "',"
        If TD1.PropertyItems("TUBE_CODE").Tag <> "" Then
            TempStr2 = TempStr2 + vbCrLf + "Tube_Code=" + Left(TD1.PropertyItems("TUBE_CODE").Tag, InStr(1, TD1.PropertyItems("TUBE_CODE").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "Tube_Code=NULL,"
        End If
        If TD1.PropertyItems("PAED_TUBE_CODE").Tag <> "" Then
            TempStr2 = TempStr2 + "PaedTube_Code=" + Left(TD1.PropertyItems("PAED_TUBE_CODE").Tag, InStr(1, TD1.PropertyItems("PAED_TUBE_CODE").Tag, "-") - 1) + ","
        Else
            TempStr2 = TempStr2 + "PaedTube_Code=NULL,"
        End If
        TempStr2 = TempStr2 + "ResHistory_SearchType='" + Format(TD1.PropertyItems("HIST_ST").value) + "',ResHistory_SearchString='" + Format(TD1.PropertyItems("HIST_SS").value) + "',ResHistory_Message='" + Format(TD1.PropertyItems("HIST_MSG").value) + "',"
        TempStr2 = TempStr2 + vbCrLf + "Dynamic=" + Format(DYN) + ",Sensitive=" + Format(SENS) + ",Worklist_Enabled=" + Format(WL) + ",Enabled=" + Format(en) + ",Date_Added='" + Format(Date, "DD MMM YYYY") + "',"
        If Format(TD1.PropertyItems("VOL").value) <> "" Then
            TempStr2 = TempStr2 + "Test_Volume=" + Format(TD1.PropertyItems("VOL").value) + ","
        Else
            TempStr2 = TempStr2 + "Test_Volume=NULL,"
        End If
        If Format(TD1.PropertyItems("PAED_VOL").value) <> "" Then
            TempStr2 = TempStr2 + "PaedTube_Test_Volume=" + Format(TD1.PropertyItems("PAED_VOL").value)
        Else
            TempStr2 = TempStr2 + "PaedTube_Test_Volume=NULL"
        End If
        TempStr = TempStr2 + vbCrLf + TempStr1
        Debug.Print TempStr
        ICECon.Execute TempStr
        TreeView1.SelectedItem.Text = LTrim(TD1.PropertyItems("SCREEN_CAPTION").value)
        If en = 0 Then
            TreeView1.SelectedItem.ForeColor = vbRed
        Else
            TreeView1.SelectedItem.ForeColor = vbBlack
        End If
        TreeView1.Visible = False
        TreeView1.Sorted = True
        TreeView1.Visible = True
    End If
    Command1.Enabled = False
    Command2.Enabled = False
    SSListBar1.Enabled = True
    TreeView1.Enabled = True
    'LoadOrgTests OrgList.Text
    OrgList.Enabled = True
    fView.Show Fra_TESTDETAILS
'    TestDetails.Visible = False
End Sub

Private Sub Command2_Click()
    Command1.Enabled = False
    Command2.Enabled = False
    SSListBar1.Enabled = True
    TreeView1.Enabled = True
    OrgList.Enabled = True
    TI = Mid$(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
    LoadTestDetails TI
End Sub

Private Sub Command3_Click()
    frmBehaviour.Show 1
    WriteTestRules TreeView1.SelectedItem
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        MsgBox "The character ' is not permitted, please use the ` character instead"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    SSSplitter1.Panes(0).LockHeight = True
    SSSplitter1.Panes(1).LockWidth = False
    SSSplitter1.Panes(4).LockHeight = True
    SSSplitter1.Panes(5).LockHeight = True
    SSSplitter1.Panes(4).Width = SSSplitter1.Panes(1).Width
    SSSplitter1.Panes(4).LockWidth = False
    For i = 0 To FLActivity.UBound
      FLActivity(i).Visible = False
      If i <= FLlbl.UBound Then
         FLlbl(i).Visible = False
      End If
    Next i
    SSListBar1.Enabled = False
    'frmMain.Left = 195
    'frmMain.Top = 0
    'frmMain.Height = 11265
    Timer2.Enabled = False
    LogBackColour = LogText.BackColor
    StatusBar1.Panels(3).Text = "Copyright © Anglia Healthcare Systems Ltd, 2000.  All Rights Reserved."
    ReadINI
    If Dir(ConfigPath + "\ICETest.UDL") <> "" Or Dir(ConfigPath + "\ICETraining.UDL") <> "" Then
        frmDatabase.Show 1
    Else
        DB_Name = "Live"
        DB_UDL_FILE = ConfigPath + "\ICE.UDL"
    End If
    If Dir(DB_UDL_FILE) = "" Then
        MsgBox "An error has occured whilst trying to open the database connection.  Please contact your system administrator.", vbCritical + vbOKOnly, "ICE...Configuration"
        End
    Else
        Set ICECon = GetConnection
    End If
    NewTop = frmSplash.Top - (frmLogon.Height / 2) - 200
    For i = frmSplash.Top To NewTop Step -1
        frmSplash.Top = i
    Next i
    frmLogon.Left = (Screen.Width - frmLogon.Width) / 2
    frmLogon.Top = ((Screen.Height - frmSplash.Height) / 2) + frmSplash.Height - (frmLogon.Height / 2)
    frmLogon.Show 1
    GetOrganisations
 '   SetupTestDetailsList
    Me.Caption = Me.Caption + " (" & DB_Name & " Database) - Version " & Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
'    CfgPanelHelp = Label6.Caption
    Timer1.Enabled = False
    Timer1.Interval = 200
    blnAddStatus = False
    fView.Hide
    Unload frmSplash
'    fView.Show fra_HELP, ""
'    Form1.Show 1
    If ShowRequesting Then
        SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Configuration Items", "Configuration Items"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Clinical Science Tests", "Clinical Science Tests"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Profiles", "Profiles"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Picklists", "Picklists"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "System Configuration", "System Configuration"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
    End If
    If ShowUsers Then
        SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "User Management", "User Management"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Users", "Users"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 2
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Groups", "Groups"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 3
    End If
    If ShowEDI Then
        SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "EDI Management", "EDI Management"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "My Settings", "My Settings"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 4
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "EDI Recipients", "EDI Recipients"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 3
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "READ Codes", "READ Codes"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 7
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Result Mapping", "Result Mapping"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 1
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "UOM Mapping", "UOM Mapping"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 9
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Specimen Types", "Specimen Types"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 6
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Körner Medical Specialties", "Körner Medical Specialties"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 7
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Specimen Anatomical Origins", "Specimen Anatomical Origins"
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 12
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Specimen Collection Procedures", "Specimen Collection Procedures"
'        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 11
    End If
    If ShowAudit Then
        SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Audit Logs", "Audit Logs"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Logs", "Logs"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 13
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Search", "Search"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 16
    End If
    If ShowConnections Then
        SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Connections", "Connections"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Connections", "Connections"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 5
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Maps", "Maps"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 15
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Modules", "Modules"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 14
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "System Activity", "System Activity"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 16
    End If
    SSListBar1.Groups.Remove SSListBar1.Groups(1)
    SSListBar1.Groups.Add SSListBar1.Groups.Count + 1, "Help", "Help"
    SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "About", "About"
    SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 17
    If Dir(App.Path + "\ICEKeyGen.EXE") <> "" Then
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Add SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count + 1, "Key Generation", "Key Generation"
        SSListBar1.Groups(SSListBar1.Groups.Count).ListItems(SSListBar1.Groups(SSListBar1.Groups.Count).ListItems.Count).IconLarge = 18
    End If
    Select Case DefaultItem
        Case "Clinical Science Tests"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Configuration Items" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 2
            CI = 1
        Case "Profiles"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Configuration Items" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 2
            CI = 2
        Case "Picklists"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Configuration Items" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 2
            CI = 3
        Case "System Configuration"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Configuration Items" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 2
            CI = 4
        Case "Users"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "User Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 3
            CI = 1
        Case "Groups"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "User Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 3
            CI = 1
        Case "My Settings"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 1
        Case "EDI Recipients"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 2
        Case "READ Codes"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 3
        Case "Result Mapping"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 4
        Case "UOM Mapping"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 5
        Case "Specimen Type"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 6
        Case "Körner Medical Specialties"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 7
        Case "Specimen Anatomical Origins"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 8
        Case "Specimen Collection Procedures"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "EDI Management" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 4
            CI = 9
        Case "Logs"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Audit Logs" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 5
            CI = 1
        Case "Search"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Audit Logs" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            CG = 5
            CI = 2
        Case "Connections"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Connections" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 6
            CI = 1
        Case "Maps"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Connections" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 6
            CI = 2
        Case "Modules"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Connections" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 6
            CI = 3
        Case "System Activity"
            For i = 1 To SSListBar1.Groups.Count
                If SSListBar1.Groups(i).Caption = "Connections" Then
                    CG = i + 1
                    Exit For
                End If
            Next i
            'CG = 6
            CI = 4
    End Select
    SSListBar1.CurrentGroup = SSListBar1.Groups(CG - 1)
    SSListBar1_ListItemClick SSListBar1.CurrentGroup.ListItems(CI)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If MsgBox("Are you sure you wish to close ICE...Configuration?", vbQuestion + vbYesNo, "ICE...Configuration") = vbNo Then
      Cancel = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ICECon.Close
   Set objTView = Nothing
   Set objctrl = Nothing
   Set fView = Nothing
   Set eClass = Nothing
   End
End Sub

Public Sub ReadINI()
    ini = App.Path + "\ICEConfig.INI"
    ConfigPath = Read_Ini_Var("General", "ConfigPath", ini)
    DefOrgID = Read_Ini_Var("General", "DefOrgID", ini)
    ShowRequesting = Read_Ini_Var("General", "ShowRequesting", ini)
    ShowUsers = Read_Ini_Var("General", "ShowUsers", ini)
    ShowEDI = Read_Ini_Var("General", "ShowEDI", ini)
    ShowAudit = Read_Ini_Var("General", "ShowAudit", ini)
    ShowConnections = Read_Ini_Var("General", "ShowConnections", ini)
    ShowHelp = Read_Ini_Var("General", "ShowHelp", ini)
    DefaultItem = Read_Ini_Var("General", "DefaultItem", ini)
End Sub

Public Sub GetOrganisations()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select Organisation_National_Code From Organisation", ICECon, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            OrgList.AddItem Trim(RS!Organisation_National_Code)
            RS.MoveNext
        Loop
    End If
    OrgList.AddItem "Create New..."
    For i = 0 To OrgList.ListCount - 1
        If OrgList.List(i) = DefOrgID Then
            OrgList.ListIndex = i
            Exit For
        End If
    Next i
    If OrgList.ListIndex = -1 Then
      MsgBox "Default Organisation from config file (" & DefOrgID & ") not found. Organisation set to " & OrgList.List(0), vbInformation, "Invalid Organisation"
      DefOrgID = OrgList.List(0)
      OrgList.ListIndex = 0
   End If
End Sub

Private Sub MailBtn_Click()
   Dim RS As New ADODB.Recordset
   Dim RS2 As New ADODB.Recordset
   Dim maxLen As Integer
   Dim cnt As Integer
   Dim SQLstr As String
   frmRequeue.QueList.Text = ""
   maxLen = 0
   ReDim RepQue(0 To 1)
   cnt = 0
   
   Select Case MailBtn.Caption
      Case "REQUEUE ALL REPORTS"
         RS.Open "SELECT * FROM Service_ImpExp_Messages WHERE Service_ImpExp_Id = " & Val(Trim(RequeueStr)) & _
                       " ORDER BY Service_Message_Id", ICECon, adOpenKeyset, adLockReadOnly
         If RS.RecordCount > 0 Then
            ICECon.BeginTrans
            While Not RS.EOF
               SQLstr = "SELECT EDI_Org_NatCode from EDI_Recipient_Individuals WHERE EDI_Local_Key1 = '" & Trim(Mid$(RS!Destination, 7, 10)) & _
                            "' AND EDI_Local_Key3 = 'ZZZ'"
               RS2.Open SQLstr, ICECon, adOpenKeyset, adLockReadOnly
               If RS2.BOF And RS2.EOF Then
                  RS2.Close
                  SQLstr = "SELECT DISTINCT EDI_Org_NatCode from EDI_Recipient_Individuals WHERE EDI_Local_Key1 = '" & _
                                 Trim(Mid$(RS!Destination, 7, 10)) & "'"
                  RS2.Open SQLstr, ICECon, adOpenKeyset, adLockReadOnly
               End If
               SQLstr = "Insert into EDI_Rep_List " & _
                              "(EDI_Report_Index,EDI_Provider_Org,EDI_Loc_Nat_Code_to,EDI_Service_ID,date_added) " & _
                              "Values (" & RS!Service_Report_Index & ",'" & _
                                                 Mid$(RS!Destination, 1, 6) & "','" & _
                                                 RS2!EDI_Org_NatCode & "','" & _
                                                 RS!Service_Id & "','" & _
                                                 CLDate(Format(Date, "dd mmm yyyy")) & " " & Format(Time, "hh:mm:ss") & "')"
               ICECon.Execute SQLstr
               frmRequeue.QueList.Text = frmRequeue.QueList.Text & RS2!EDI_Org_NatCode & vbTab & RS!Service_Id & _
                                                             vbTab & Trim(RS!Patient_Name) & vbCrLf
               If maxLen < (Len(RS2!EDI_Org_NatCode) + Len(RS!Service_Id) + Len(Trim(RS!Patient_Name)) + 10) Then
                  maxLen = (Len(RS2!EDI_Org_NatCode) + Len(RS!Service_Id) + Len(Trim(RS!Patient_Name)) + 10)
               End If
               RS2.Close
               RS.MoveNext
            Wend
         End If
         RS.Close
         SQLstr = "Insert into Service_ImpExp_Comments " & _
                        "(Service_ImpExp_Id, Service_ImpExp_Comment) " & _
                        "Values (" & Val(Trim(RequeueStr)) & ", " & _
                        "'ALL Reports Requeued for Output " & _
                        CLDate(Format(Date, "dd mmm yyyy")) & " " & Format(Time, "hh:mm:ss") & "')"
         ICECon.Execute SQLstr
         SQLstr = "Update Service_ImpExp_Headers set Comment_marker = " & 1 & " where Service_ImpExp_ID = " & Val(Trim(RequeueStr))
         ICECon.Execute SQLstr
         frmRequeue.Label1.Caption = "Reports to be requeued for retransmission"
         
      Case "REQUEUE THIS REPORT"
          RS.Open "Select * From service_impexp_messages Where service_impexp_message_id = " & Val(Trim(RequeueStr)) & " order by service_message_id", ICECon, adOpenKeyset, adLockReadOnly
          If RS.RecordCount > 0 Then
            ICECon.BeginTrans
            While Not RS.EOF
               SQLstr = "Insert into EDI_Rep_List " & _
                              "(EDI_Report_Index,EDI_Provider_Org,EDI_Loc_Nat_Code_to,EDI_Service_ID,date_added) " & _
                              "Values (" & RS!Service_Report_Index & ", '" & _
                                             Mid$(RS!Destination, 1, 6) & "','" & _
                                             Mid$(RS!Destination, 7, 10) & "','" & _
                                             RS!Service_Id & "','" & _
                                             CLDate(Format(Date, "dd mmm yyyy")) & " " & Format(Time, "hh:mm:ss") & "')"
               ICECon.Execute SQLstr
               SQLstr = "Insert into Service_ImpExp_Comments " & _
                              "(Service_ImpExp_Id,Service_ImpExp_Comment) " & _
                              "Values (" & RS!Service_ImpExp_Id & "," & _
                                            "'Requeued for Ouput  " & CLDate(Format(Date, "dd mmm yyyy")) & " " & Format(Time, "hh:mm:ss") & "')" & _
                                            "Insert into Service_ImpExp_Comments " & _
                                            "(Service_ImpExp_Id,Service_ImpExp_Comment) " & _
                                            "Values (" & RS!Service_ImpExp_Id & "," & _
                                                          "' " & vbTab & Mid$(RS!Destination, 7, 10) & vbTab & RS!Service_Id & vbTab & Trim(RS!Patient_Name) & "')"
               ICECon.Execute SQLstr
               SQLstr = "Update Service_ImpExp_Headers set Comment_marker = " & 1 & " where Service_ImpExp_ID = " & RS!Service_ImpExp_Id
               ICECon.Execute SQLstr
               frmRequeue.QueList.Text = frmRequeue.QueList.Text & Mid$(RS!Destination, 7, 10) & vbTab & RS!Service_Id & vbTab & Trim(RS!Patient_Name) & vbCrLf
               If maxLen < (Len(Mid$(RS!Destination, 7, 10)) + Len(RS!Service_Id) + Len(Trim(RS!Patient_Name)) + 10) Then
                  maxLen = (Len(Mid$(RS!Destination, 7, 10)) + Len(RS!Service_Id) + Len(Trim(RS!Patient_Name)) + 10)
               End If
               RS.MoveNext
            Wend
         End If
         RS.Close
         frmRequeue.Label1.Caption = "Report to be requeued for retransmission"
         
      Case "RESEND EDI FILE"
         RS.Open "Select * From service_impexp_headers Where service_impexp_id = " & Val(Trim(RequeueStr)), ICECon, adOpenKeyset, adLockReadOnly
         If RS.RecordCount > 0 Then
            ICECon.BeginTrans
            SQLstr = "Insert into Service_ImpExp_Comments " & _
                           "(Service_ImpExp_Id,Service_ImpExp_Comment) " & _
                           "Values (" & Val(Trim(RequeueStr)) & "," & _
                                         "'EDI File Requeued for Output " & _
                                          CLDate(Format(Date, "dd mmm yyyy")) & " " & Format(Time, "hh:mm:ss") & "')" & _
                           "Insert into Service_ImpExp_Comments " & _
                                    "(Service_ImpExp_Id,Service_ImpExp_Comment) " & _
                           "Values (" & Val(Trim(RequeueStr)) & "," & _
                                         "'" & vbTab & RS!ImpExp_File & "')"
            ICECon.Execute SQLstr
            SQLstr = "Update Service_ImpExp_Headers set Comment_marker = " & 1 & " where Service_ImpExp_ID = " & Val(Trim(RequeueStr))
            ICECon.Execute SQLstr
         End If
         frmRequeue.QueList.Text = Trim(RS!ImpExp_File) & vbCrLf
         If maxLen < (Len(Trim(RS!ImpExp_File))) Then
            maxLen = (Len(Trim(RS!ImpExp_File)))
         End If
         RS.Close
         
         RS.Open "Select * From service_impexp_messages Where service_impexp_id = " & Val(Trim(RequeueStr)) & " order by service_message_id", ICECon, adOpenKeyset, adLockReadOnly
         If RS.RecordCount > 0 Then
            While Not RS.EOF
               frmRequeue.QueList.Text = frmRequeue.QueList.Text & Mid$(RS!Destination, 7, 10) & vbTab & RS!Service_Id & vbTab & Trim(RS!Patient_Name) & vbCrLf
               If maxLen < (Len(Mid$(RS!Destination, 7, 10)) + Len(RS!Service_Id) + Len(Trim(RS!Patient_Name)) + 10) Then
                  maxLen = (Len(Mid$(RS!Destination, 7, 10)) + Len(RS!Service_Id) + Len(Trim(RS!Patient_Name)) + 10)
               End If
               RS.MoveNext
            Wend
         End If
         RS.Close
         frmRequeue.Label1.Caption = "EDI File to be requeued for retransmission"
         
   End Select
   frmRequeue.QueList.Width = maxLen * 110
   frmRequeue.Width = frmRequeue.QueList.Width + 1000
   Load frmRequeue
   frmRequeue.Show
End Sub

Private Sub optLogSearch_Click(Index As Integer)

   If Index = 0 Then
      dtPTo.value = Now()
      dtPFrom.value = DateAdd("d", -10, dtPTo.value)
   End If
   fView.SetUpPanel Fra_LOGSEARCH, CStr(Index)

End Sub

Private Sub itemAdd_Click()
    objTView.MenuAddEntry itemId
End Sub
'
Private Sub itemDelete_Click()

   If nOrigin = "T" Then
      objTView.MenuDeleteEntry TreeView1.SelectedItem.Text
   Else
      objTView.MenuDeleteEntry ediPr(objTView.NodeKey(TreeView1.SelectedItem.Key)).value
      fView.RefreshDisplay
   End If
   
End Sub

Private Sub OrgList_Click()
    If OrgList.Text <> "" Then
        If OrgList.Text = "Create New..." Then
            MsgBox "Not yet available"
        Else
            OrgName.Caption = GetOrganisationName(OrgList.Text)
            SSListBar1.Enabled = True
            TreeView1.Visible = False
            TreeView1.Nodes.Clear
            TreeView1.Visible = True
'            TestDetails.Visible = False
'            Behaviour.Visible = False
        End If
    End If
End Sub

Private Sub PR1_BeforeEdit(PropertyItem As PropertiesListCtl.PropertyItem, Cancel As Boolean)
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = False
    Command14.Enabled = False
    SSListBar1.Enabled = False
    TreeView1.Enabled = False
    OrgList.Enabled = False
    If PropertyItem.Style = plpsColor Then
        frmColour.Show 1
        If PickedColIndex > 0 Then
            PropertyItem.value = PickedCol
            PropertyItem.Tag = Format(PickedColIndex) + "-" + PickedColName
        End If
        'PropertyItem.Description = PropertyItem.Tag
        Cancel = True
    End If
End Sub

Private Sub PR1_PropertyBrowseClick(PropertyItem As PropertiesListCtl.PropertyItem)
    Select Case PropertyItem.Key
        Case "PROV_ID"
            frmProvider.Show 1
            PropertyItem.value = PickedProv
            PropertyItem.Tag = PickedProvIndex
    End Select
End Sub

Private Sub PR1_RequestDisplayValue(PropertyItem As PropertiesListCtl.PropertyItem, DisplayValue As String)
    If PropertyItem.Tag = "" Then Exit Sub
    Select Case PropertyItem.Key
        Case "PROFILE_COLOUR"
            DisplayValue = Mid$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") + 1, Len(PropertyItem.Tag) - InStr(PropertyItem.Tag, "-") + 1) + " (" + Left$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") - 1) + ")"
        Case "HELP_COLOUR"
            DisplayValue = Mid$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") + 1, Len(PropertyItem.Tag) - InStr(PropertyItem.Tag, "-") + 1) + " (" + Left$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") - 1) + ")"
    End Select
End Sub

Private Sub ssCmdEdiCancel_Click()
   
   fView.Show Fra_HELP, ""
   fView.RefreshDisplay
'   blnAddSpec = False
      
End Sub

Private Sub ssCmdEDIok_Click()
   On Local Error GoTo ProcEH
   '  Update database with new values
         
   If objctrl.GroupKey = "New" And objTView.NodeOrigin = "U" Then
      objctrl.GroupKey = ""
      tPList.Pages("X400").Selected = True
      tPList("X4").EnsureVisible
   Else
      objctrl.UdateDatabase
      fView.RefreshDisplay
   End If
         
   Exit Sub
         
ProcEH:
   eClass.Show
   
End Sub

Private Sub SSListBar1_GroupClick(ByVal GroupClicked As Listbar.SSGroup, ByVal PreviousGroup As Listbar.SSGroup)
Dim i As Integer
   fView.Show Fra_HELP, "1"
'  LogFrame.Visible = False
'  LogText.Visible = False
'  LogViews.Visible = False
  Timer2.Enabled = False
'  For i = 0 To FLActivity.UBound
'    FLActivity(i).Visible = False
'    If i <= FLlbl.UBound Then
'      FLlbl(i).Visible = False
'    End If
'  Next i
'  LogViews.Caption = " Log Information"
Select Case GroupClicked.Index
    Case 6
        'Help button
'      LogText.BackColor = vbYellow
'      LogViews.Caption = " Help"
'      LogText.Text = "This will be  help information" & vbCrLf
'      LogFrame.Visible = True
'      LogText.Visible = True
'      LogViews.Visible = True
   fView.Show Fra_HELP, "1"
    
End Select
End Sub

Private Sub SSListBar1_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
   On Local Error GoTo ProcEH
   Dim i As Integer
   Dim Temp(2) As String
    
   Timer2.Enabled = False
   MailBtn.Visible = False
   
   Set objctrl = Nothing
   Set objctrl = New Class1
   
   
   Select Case ItemClicked.Text
      Case "Clinical Science Tests"
         fView.RefreshProc = "LoadOrgTests"
         fView.RefreshProcParams = OrgList.Text
'         LoadOrgTests (OrgList.Text)
        
      Case "Profiles"
         fView.RefreshProc = "LoadOrgProfiles"
         fView.RefreshProcParams = OrgList.Text
'         LoadOrgProfiles (OrgList.Text)
        
      Case "Picklists"
         fView.RefreshProc = "LoadPickLists"
'         LoadPicklists
        
      Case "Users"
      
      Case "Groups"
        
      Case "System Configuration"
         fView.RefreshProc = "LoadConfiguration"
         fView.RefreshProcParams = OrgList.Text
'         LoadConfiguration (OrgList.Text)
        
      Case "Connections"
         fView.RefreshProc = "LoadConnections"
         fView.RefreshProcParams = OrgList.Text
'         LoadConnections (OrgList.Text)
        
      Case "Maps"
         fView.RefreshProc = "LoadMaps"
         fView.RefreshProcParams = OrgList.Text
'         LoadMaps (OrgList.Text)
        
      Case "Modules"
         fView.RefreshProc = "LoadModules"
         fView.RefreshProcParams = OrgList.Text
'         LoadModules (OrgList.Text)
        
      Case "System Activity"
         fView.RefreshProc = "LoadMonitors"
         fView.RefreshProcParams = OrgList.Text
'         LoadMonitors (OrgList.Text)
        
      Case "READ Codes"
         fView.RefreshProc = "LoadReadCodes"
'         objTView.LoadReadCodes
        
      Case "My Settings"
         fView.RefreshProc = "LoadMySettings"
         fView.RefreshProcParams = OrgList.Text
         'objTView.LoadMySettings (OrgList.Text)
        
      Case "EDI Recipients"
         fView.RefreshProc = "LoadEDIRecipients"
         fView.RefreshProcParams = OrgList.Text
         currentAccount = ""
        
      Case "Result Mapping"
         fView.RefreshProc = "LoadResultMapping"
         fView.RefreshProcParams = OrgList.Text
'         objTView.LoadResultMapping (OrgList.Text)
        
      Case "UOM Mapping"
         fView.RefreshProc = "LoadUOMMap"
         fView.RefreshProcParams = OrgList.Text
'         objTView.LoadUOMMap (OrgList.Text)
        
      Case "Specimen Types"
         fView.RefreshProc = "LoadSpecimenCodes"
'         objTView.LoadSpecimenCodes
        
      Case "Körner Medical Specialties"
         fView.RefreshProc = "LoadKornerCodes"
'         objTView.LoadKornerCodes
        
      Case "Logs"
         fView.Show Fra_INFO, "Please wait whilst Log details for " & OrgList.Text & " are loaded..."
         fView.RefreshProc = "LoadLogs"
         fView.RefreshProcParams = "log"
'         objTView.LoadLogs ("log")
         
      Case "Search"
         fView.RefreshProc = ""
         txtSrchSurname.Text = ""
         txtSrchForename.Text = ""
         frmMain.MousePointer = vbNormal
         objTView.GetPractices
         fView.Show Fra_LOGSEARCH, "0"
      Case "About"
        frmAbout.Show 1
        Exit Sub
      Case "Key Generation"
        Shell App.Path + "\ICEKeyGen.EXE", vbNormalFocus
        Exit Sub
      Case Else
        MsgBox "This section has not yet been implemented.  Please check with your supplier for a later version.", vbInformation + vbOKOnly, "ICE...Configuration"
    End Select
   fView.RefreshDisplay
   Exit Sub
   
ProcEH:
   eClass.Show
   
End Sub


Private Sub Load_Report(OrgID As String, ServiceId As String)
   
   Dim RepId As Long
   Dim maxLen As Integer
   Dim TestLen As Integer
   Dim K As Integer
   Dim MaxRES As Integer
   Dim MaxUOM As Integer
   Dim TabStr As String
    
    Set Reports = New AHSLReporting.Reports
    Reports.SetParameters OrgID, ""
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open "Select service_report_index from service_reports where service_report_index = '" & ServiceId & "'", ICECon, adOpenKeyset, adLockReadOnly
    RepId = 0
    If RS.RecordCount > 0 Then
      RepId = RS!Service_Report_Index
      Report = Reports.GetReportByID(RepId)
      With Report
        LogText.Visible = False
        LogText.Text = Trim(Report.Patient.Surname) & ", " & Trim(Report.Patient.Forename) & vbTab & Report.Patient.dob & "  " & Report.Patient.Sex
        'LogText.Text = LogText.Text & vbCrLf & "Forename:" + vbTab + Report.Patient.Forename
        'LogText.Text = LogText.Text & vbCrLf & "Gender: " + vbTab + Report.Patient.Sex
        'LogText.Text = LogText.Text & vbCrLf & "D.o.B.: " + vbTab + Report.Patient.dob
        LogText.Text = LogText.Text & vbCrLf & "Hosp. No.: " + vbTab + Report.Patient.HospNo
        LogText.Text = LogText.Text & vbTab & "Service ID:  " + Report.LabNo
        LogText.Text = LogText.Text & vbCrLf & "Clinician/Specialty: " + Trim(Report.Clinician) & "  " & Report.ClinSpecialty
        'LogText.Text = LogText.Text & vbCrLf & "Clin. Specialty: " + Report.ClinSpecialty
        LogText.Text = LogText.Text & vbCrLf & "Destination: " + Report.Destination
        LogText.Text = LogText.Text & vbCrLf & "Report Date: " + Report.ReportProducedDate
        LogText.Text = LogText.Text & vbTab & "Type: " + Report.RepSpecialty
        Dim RS2 As ADODB.Recordset
        Set RS2 = New ADODB.Recordset
        RS2.Open "select colour_code from service_tubes_colours where colour_name in (select report_colour from service_reports_colours where report_type in (select specialty_code from specialty where specialty like '" & Report.RepSpecialty & "'))", ICECon, adOpenKeyset, adLockReadOnly
        If RS2.RecordCount > 0 Then
            LogBackColour = RS2!Colour_Code
'            LogText.BackColor = RS2!Colour_Code
        End If
        RS2.Close
        Set RS2 = Nothing
        LogText.Text = LogText.Text & vbCrLf & "Sample Type: " + Report.Samples(0).SampleType & vbCrLf
        LogText.Text = LogText.Text & vbTab & "Collected: " + Report.Samples(0).CollectionDateTime
        LogText.Text = LogText.Text & vbTab & "Received: " + Report.Samples(0).CollectionDateTime_Received
        If Report.ContainsAbnormal Then
            LogText.Text = LogText.Text & vbCrLf & "CONTAINS OUT OF RANGE RESULTS" & vbCrLf
        'Else
        '    LogText.Text = LogText.Text & vbCrLf & "No abnormal results"
        End If
        If Reports.HasEntries(Report.Comments) Then
            For i = 0 To UBound(Report.Comments)
                LogText.Text = LogText.Text & vbCrLf & vbTab & Report.Comments(i)
            Next i
        End If
        If Reports.HasEntries(Report.Investigations) Then
            maxLen = 0
            MaxUOM = 0
            MaxRES = 0
            For i = 0 To UBound(Report.Investigations)
                LogText.Text = LogText.Text & vbCrLf & vbCrLf & Report.Investigations(i).Investigation_Requested
                If Reports.HasEntries(Report.Investigations(i).Comment) Then
                    For j = 0 To UBound(Report.Investigations(i).Comment)
                        LogText.Text = LogText.Text & vbCrLf & vbTab & Report.Investigations(i).Comment(j)
                    Next j
                End If
                If Reports.HasEntries(Report.Investigations(i).Results) Then
                  For j = 0 To UBound(Report.Investigations(i).Results)
                      If Len(Report.Investigations(i).Results(j).Test) > maxLen Then
                          maxLen = Len(Report.Investigations(i).Results(j).Test)
                      End If
                      If Len(Report.Investigations(i).Results(j).Units) > MaxUOM Then
                          MaxUOM = Len(Report.Investigations(i).Results(j).Units)
                      End If
                      If Len(Report.Investigations(i).Results(j).Result) > MaxRES Then
                          MaxRES = Len(Report.Investigations(i).Results(j).Result)
                      End If
                  Next j
                  
                  For j = 0 To UBound(Report.Investigations(i).Results)
                    TabStr = ""
                    For K = Len(Report.Investigations(i).Results(j).Test) To maxLen
                        TabStr = TabStr & " "
                    Next K
                    LogText.Text = LogText.Text & vbCrLf & "  " & Report.Investigations(i).Results(j).Test
                    'If Report.Investigations(I).Results(J).Abnormal Then TreeView1.Nodes(TreeView1.Nodes.Count).ForeColor = vbRed
                    If Report.Investigations(i).Results(j).Abnormal Then
                        'LogText.Font.Bold
                        LogText.Text = LogText.Text & TabStr & " *" + Report.Investigations(i).Results(j).Result
                        'LogText.Font.Bold
                        '.ForeColor = vbRed
                    Else
                        LogText.Text = LogText.Text & TabStr & "  " & Report.Investigations(i).Results(j).Result
                    End If
                    TabStr = ""
                    For K = Len(Report.Investigations(i).Results(j).Result) To MaxRES
                        TabStr = TabStr & " "
                    Next K
                    LogText.Text = LogText.Text & TabStr & Report.Investigations(i).Results(j).Units
                    TabStr = ""
                    For K = Len(Report.Investigations(i).Results(j).Units) To MaxUOM
                        TabStr = TabStr & " "
                    Next K
                    LogText.Text = LogText.Text & TabStr & Report.Investigations(i).Results(j).Range
                    If Reports.HasEntries(Report.Investigations(i).Results(j).Comment) Then
                        For K = 0 To UBound(Report.Investigations(i).Results(j).Comment)
                            LogText.Text = LogText.Text & vbCrLf & "   " + Report.Investigations(i).Results(j).Comment(K)
                        Next K
                    End If
                Next j
               End If
            Next i
        End If
        LogText.Visible = True
      End With
    End If
    RS.Close
End Sub

Private Sub TD1_AfterEdit(PropertyItem As PropertiesListCtl.PropertyItem, NewValue As Variant, Cancel As Boolean)
    If PropertyItem.Key = "SCREEN_PANEL" And PropertyItem.value <> NewValue Then TD1.PropertyItems("PANEL_PAGE").value = ""
End Sub

Private Sub TD1_BeforeEdit(PropertyItem As PropertiesListCtl.PropertyItem, Cancel As Boolean)
    Command1.Enabled = True
    Command2.Enabled = True
    SSListBar1.Enabled = False
    TreeView1.Enabled = False
    OrgList.Enabled = False
    If PropertyItem.Style = plpsColor Then
        If PropertyItem.Key <> "TUBE_CODE" And PropertyItem.Key <> "PAED_TUBE_CODE" Then
            frmColour.Show 1
            If PickedColIndex > 0 Then
                PropertyItem.value = PickedCol
                PropertyItem.Tag = Format(PickedColIndex) + "-" + PickedColName
            End If
            'PropertyItem.Description = PropertyItem.Tag
            Cancel = True
        Else
            frmTube.Show 1
            If PickedTubeIndex > 0 Then
                PropertyItem.value = PickedTubeCol
                PropertyItem.Tag = Format(PickedTubeIndex) + "-" + PickedTube
            End If
            Cancel = True
        End If
    End If
    If PropertyItem.Key = "SCREEN_POSN" And Not GVSP Then
'        frmWait.Show
'        frmWait.Refresh
        If Format(TD1.PropertyItems("SCREEN_PANEL").value) = "" Or Format(TD1.PropertyItems("PANEL_PAGE").value) = "" Then
            MsgBox "Both the Panel and Page properties must be set prior to selecting a screen position in order for vacant screen positions to be determined", vbInformation + vbOKOnly, "Change Screen Position"
            Cancel = True
            Exit Sub
        End If
        GetVacantScreenPositions
        'GVSP = True
'        Unload frmWait
    End If
    If PropertyItem.Key = "PANEL_PAGE" Then
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        TD1.PropertyItems("PANEL_PAGE").ListItems.Clear
        RS.Open "Select PageName From Request_Panels_Pages Where PanelID=" & TD1.PropertyItems("SCREEN_PANEL").value & "Order By PageName", ICECon, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            TD1.PropertyItems("PANEL_PAGE").ListItems.Add RS!PageName, RS!PageName
            RS.MoveNext
        Loop
        End If
        RS.Close
    End If
End Sub

Private Sub TD1_PropertyBrowseClick(PropertyItem As PropertiesListCtl.PropertyItem)
    Select Case PropertyItem.Key
        Case "PROV_ID"
            frmProvider.Show 1
            If PickedProvIndex <> 0 Then
                PropertyItem.value = PickedProv
                PropertyItem.Tag = PickedProvIndex
            End If
    End Select
End Sub

Private Sub TD1_RequestDisplayValue(PropertyItem As PropertiesListCtl.PropertyItem, DisplayValue As String)
    If PropertyItem.Tag = "" Then Exit Sub
    Select Case PropertyItem.Key
        Case "SCREEN_COLOUR"
            DisplayValue = Mid$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") + 1, Len(PropertyItem.Tag) - InStr(PropertyItem.Tag, "-") + 1) + " (" + Left$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") - 1) + ")"
        Case "HELP_COLOUR"
            DisplayValue = Mid$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") + 1, Len(PropertyItem.Tag) - InStr(PropertyItem.Tag, "-") + 1) + " (" + Left$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") - 1) + ")"
        Case "TUBE_CODE"
            DisplayValue = Mid$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") + 1, Len(PropertyItem.Tag) - InStr(PropertyItem.Tag, "-") + 1) + " (" + Left$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") - 1) + ")"
        Case "PAED_TUBE_CODE"
            DisplayValue = Mid$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") + 1, Len(PropertyItem.Tag) - InStr(PropertyItem.Tag, "-") + 1) + " (" + Left$(PropertyItem.Tag, InStr(PropertyItem.Tag, "-") - 1) + ")"
        Case "PROV_ID"
            DisplayValue = Format(PropertyItem.value) + " (" + Format(PropertyItem.Tag) + ")"
    End Select
End Sub

Private Sub Text1_Change()
    Command7.Enabled = True
    Command8.Enabled = True
    Command5.Enabled = False
    Command6.Enabled = False
    TreeView1.Enabled = False
    SSListBar1.Enabled = False
    OrgList.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Set TreeView1.DropHighlight = TreeView1.HitTest(mfX, mfY)
    If m_iScrollDir = -1 Then 'Scroll Up
    ' Send a WM_VSCROLL message 0 is up and 1 is down
      SendMessage TreeView1.hwnd, 277&, 0&, vbNull
    Else 'Scroll Down
      SendMessage TreeView1.hwnd, 277&, 1&, vbNull
    End If
End Sub

Private Sub Timer2_Timer()
Dim i As Integer
Dim FL As Integer
FL = 0
    For i = 0 To UBound(ActivityDirs)
       If (ActivityDirs(i) <> "") Then
        FLActivity(FL).Refresh
        FL = FL + 1
       End If
    Next i
End Sub

Private Sub TreeView1_DblClick()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If Left(TreeView1.SelectedItem.Key, 1) <> "H" Then Exit Sub
    RI = Mid$(TreeView1.SelectedItem.Key, 2, InStr(1, TreeView1.SelectedItem.Key, "-") - 2)
    Load frmBehaviourEditor
    frmBehaviourEditor.Label1.Caption = RI
    frmBehaviourEditor.Show 1
End Sub

Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Not TreeView1.DropHighlight Is Nothing Then
        BN1 = Mid(TreeView1.DropHighlight.Key, InStr(TreeView1.DropHighlight.Key, "-") + 1, Len(TreeView1.DropHighlight.Key) - InStr(TreeView1.DropHighlight.Key, "-"))
        BN2 = Mid(TreeView1.SelectedItem.Key, InStr(TreeView1.SelectedItem.Key, "-") + 1, Len(TreeView1.SelectedItem.Key) - InStr(TreeView1.SelectedItem.Key, "-"))
        If BN2 = BN1 And TreeView1.DropHighlight.Key <> TreeView1.SelectedItem.Key Then
            TK = TreeView1.SelectedItem.Key
            tt = TreeView1.SelectedItem.Text
            If Not TreeView1.DropHighlight Is Nothing Then
                If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' below '" & TreeView1.DropHighlight.Text & "'?", vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
                    TreeView1.Nodes.Remove (TK)
                    TreeView1.Nodes.Add TreeView1.DropHighlight, tvwNext, TK, tt, 1, 1
                    WriteTestRules TreeView1.DropHighlight.Parent
                End If
            End If
            'MsgBox "Create node below " & TreeView1.DropHighlight.Text
        End If
        If "B" + BN2 = BN1 And TreeView1.DropHighlight.Key <> TreeView1.SelectedItem.Key Then
            TK = TreeView1.SelectedItem.Key
            tt = TreeView1.SelectedItem.Text
            If Not TreeView1.DropHighlight Is Nothing Then
                If MsgBox("Are you sure you wish to move '" & TreeView1.SelectedItem.Text & "' above '" & TreeView1.DropHighlight.Child.Text & "'?", vbYesNo + vbQuestion, "Re-order Rules") = vbYes Then
                    TreeView1.Nodes.Remove (TK)
                    TreeView1.Nodes.Add TreeView1.DropHighlight.Child, tvwFirst, TK, tt, 1, 1
                    WriteTestRules TreeView1.DropHighlight
                End If
            End If
        End If
  End If
  Set TreeView1.DropHighlight = Nothing
  Set moNode = Nothing
  Timer1.Enabled = False
End Sub

Private Sub TreeView1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
    mfX = x
    mfY = y
    If y > 0 And y < 100 Then 'scroll up
      m_iScrollDir = -1
      Timer1.Enabled = True
    ElseIf y > (TreeView1.Height - 200) And y < TreeView1.Height Then
    'scroll down
      m_iScrollDir = 1
      Timer1.Enabled = True
    Else
      Timer1.Enabled = False
    End If
    If TreeView1.DropHighlight Is Nothing Then
        TreeView1.DragIcon = ImageList1.ListImages(3).Picture
        Exit Sub
    End If
    If Left(TreeView1.DropHighlight.Key, 1) <> "H" And Left(TreeView1.DropHighlight.Key, 1) <> "B" Then
        TreeView1.DragIcon = ImageList1.ListImages(3).Picture
        Exit Sub
    End If
    BN1 = Mid(TreeView1.DropHighlight.Key, InStr(TreeView1.DropHighlight.Key, "-") + 1, Len(TreeView1.DropHighlight.Key) - InStr(TreeView1.DropHighlight.Key, "-"))
    BN2 = Mid(TreeView1.SelectedItem.Key, InStr(TreeView1.SelectedItem.Key, "-") + 1, Len(TreeView1.SelectedItem.Key) - InStr(TreeView1.SelectedItem.Key, "-"))
    If (BN2 <> BN1) And ("B" + BN2 <> BN1) Then
        TreeView1.DragIcon = ImageList1.ListImages(3).Picture
        Exit Sub
    End If
    TreeView1.DragIcon = ImageList1.ListImages(1).Picture
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    TreeView1.DropHighlight = TreeView1.HitTest(x, y)
    'Make sure we are over a Node
    If Not TreeView1.DropHighlight Is Nothing Then
       TreeView1.SelectedItem = TreeView1.HitTest(x, y)
       Set moNode = TreeView1.SelectedItem ' Set the item being dragged.
    End If
    Set TreeView1.DropHighlight = Nothing
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If Not TreeView1.SelectedItem Is Nothing Then
            If Left(TreeView1.SelectedItem.Key, 1) <> "H" Then
                Exit Sub
            Else
            End If
        End If
        If Not TreeView1.SelectedItem Is Nothing Then
            TreeView1.DragIcon = TreeView1.SelectedItem.CreateDragImage
            TreeView1.Drag vbBeginDrag
        End If
    End If
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Timer1.Enabled = False
    
    If Button = 2 Then
      If objTView.MenuStatus <> 0 Then
         PopupMenu item
         If nOrigin <> "T" Then
            PopulatePropertyList objctrl.ClassCurrency
         End If
      End If
   End If
    
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
   On Local Error GoTo ProcEH
   Dim Tstr As String
   Dim OrgStr As String
   Dim IDStr As String
   Dim natCode As String
   Dim blnfurtherChecks As Boolean
   Dim itemData As String
   Dim itemQualifier As String
   Dim fData As String
   Dim itemPos As Integer
   Dim kPos As Integer
   Dim P As Integer
   Dim E As Integer
   Dim i As Integer
   Dim TVN As Node
   Dim TVN2 As Node
   Dim RS As ADODB.Recordset
   Dim RS1 As New ADODB.Recordset
   Dim locCode As String
   Dim rSex As String
   Dim tBuf As New StringBuffer
   Dim strArray() As String
   Dim tIndex As String
   Dim errRet As Integer

   objctrl.DataGroup = "" '  Ensure the new flag is reset
   Set RS = New ADODB.Recordset
   frmMain.MousePointer = 11
   Timer2.Enabled = False
   objTView.MenuStatus = 0
   RequeueStr = ""
   nOrigin = objTView.NodeOrigin(Node.Key)
   strArray = Split(objTView.NodeLevel(Node.Key), ":")
   
   Set tPList = objTView.PreparePropertiesList(nOrigin)
   blnAddStatus = False
   objTView.NodeId = Node
   Select Case nOrigin
'    Select Case Left$(Node.Key, 1)
         Case "T"
            If objTView.NodeKey(Node.Key) = "New" Then
               tIndex = strArray(2)
               Select Case tIndex
                  Case "T"
                     SetupTestDetailsList
                     fView.Show Fra_TESTDETAILS, "0"
                  
                  Case "O"
                     SetupProfileList
                     fView.Show Fra_PROFILE
                     
                  Case "P"
                     Text1.Text = ""
                     Check1.value = vbUnchecked
                     Command7.Enabled = False
                     Command8.Enabled = False
                     OrgList.Enabled = True
                     Command5.Enabled = False
                     Command6.Enabled = False
                     fView.Show Fra_PICKLIST
                  
               End Select
               '  Add new
            Else
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               tIndex = strArray(2)
               AddMode.Caption = strArray(3)
               itemId = "orgTest"
'              Debug.Print strArray(0) & "-" & strArray(1) & "-" & strArray(2) & "-" & strArray(3) & "-" & objTView.NodeKey(Node.Key)
'             Set the popup menu status
               objTView.DataGroup = strArray(1) & ":" & strArray(2)
               objTView.curStore = strArray(3)
               objTView.Level = 2
               If IsNumeric(objTView.NodeKey(Node.Key)) Then
                  objTView.MenuStatus = ms_DELETE
'              ElseIf strArray(0) = "2" Then
'                 objTView.MenuStatus = ms_ADD
               Else
                  objTView.MenuStatus = ms_BOTH
               End If
               Select Case strArray(3)
                  Case "T"
                     SetupTestDetailsList
                     LoadTestDetails tIndex
                     fView.Show Fra_TESTDETAILS, "0"
                  
                  Case "I"
                     fView.Show Fra_INCEX, "0"
                     
                  Case "X"
                     fView.Show Fra_INCEX, "1"
                  
                  Case "R"
                     fView.Show Fra_INCEX, "2"
                  
                  Case "O"
                     strArray = Split(objTView.NodeLevel(Node.Key), ":")
                     If strArray(0) <> nl_FIRST Then
'                     If InStr(1, Node.Key, "-") = 0 Then
'                       Profile.Caption = " " + TreeView1.SelectedItem.Text + " - Profile Editor"
                        SetupProfileList
                        If tIndex = "New" Then
                           Command13.Enabled = False
                           Command14.Enabled = False
                        Else
                           LoadProfileDetails tIndex
                        End If
                        fView.Show Fra_PROFILE
'                       Profile.Visible = True
                      End If
                      'Proflile Entry
                  
                  Case "P"
                      'Picklist Entry
'                     Picklist.Caption = " " + Node.Text + " - Picklist Editor"
                      If tIndex <> "P9999" Then
                          LoadPicklistDetails tIndex
                      Else
                          Text1.Text = ""
                          Check1.value = vbUnchecked
                          Command7.Enabled = False
                          Command8.Enabled = False
                          OrgList.Enabled = True
                          Command5.Enabled = False
                          Command6.Enabled = False
                          TreeView1.Enabled = True
                          SSListBar1.Enabled = True
                      End If
                      fView.Show Fra_PICKLIST
'                     Picklist.Visible = True
                  End Select
               End If
               
            Case "A"
               eClass.FurtherInfo = "Adding/amending Read Code Mapping"

'              Set the properties list properties
               Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               nlevel = strArray(0)
               objctrl.PracticeId = strArray(1)

'              Determine the key which correlates the treeview, The properties list and data objects
               itemId = objTView.NodeKey(Node.Key)
               If itemId = "New" Then
                  itemId = tPList(1).Key
               End If
               
'              Set the popup menu status
               If itemId = "LR" Then
                  objTView.MenuStatus = ms_DELETE
               End If
               
'               Do we need to change the display for this item?
               If itemId <> "None" Then
'                 Now populate the propertis list from the relevant data object
                  PopulatePropertyList objctrl.ClassCurrency(strArray(2))
'                 Highlight the selected item in the properties list
                  tPList.Visible = False
                  Set tPList.SelectedItem = tPList(itemId)
                  tPList(itemId).EnsureVisible
                  tPList.Pages(ediPr(itemId).PageKeys).Selected = True
                  tPList.Visible = True
               End If
               
'              Now show the frame
               fView.Show Fra_EDI, "Unit of Measure Mapping"
            
        Case "C"
            'Configuration Entry
            eClass.FurtherInfo = "Amending Configuration Entries"
            If objTView.NodeKey(Node.Key) = "New" Then
               SetupConfigList
            Else
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               nlevel = strArray(0)
               CfgProgID = strArray(2)
               If nlevel > 0 Then
                  CfgCfgID = objTView.GetParentData(Node.Key, nl_FIRST, True)
               End If
               strArray = Split(objTView.NodeKey(Node.Key), ":")
               If Left(strArray(0), 4) = "Ward" Then
                  CfgWardID = strArray(1)
                  CfgUserID = ""
               Else
                  CfgUserID = strArray(1)
                  CfgWardID = ""
               End If
               If Right(strArray(0), 1) = "*" Then
                  SetupConfigList
               End If
            End If
            fView.Show Fra_CONFIG
            
         Case "D"
            eClass.FurtherInfo = "Adding/amending new Modules"

'              Set the properties list properties
            Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
            strArray = Split(objTView.NodeLevel(Node.Key), ":")
            nlevel = strArray(0)
            objctrl.PracticeId = strArray(1)
            objctrl.GroupKey = strArray(2)

'              Determine the key which correlates the treeview, The properties list and data objects
            itemId = objTView.NodeKey(Node.Key)
            If itemId = "New" Then
'                  objctrl.DataGroup = "EDI_InvTest_Codes"
               itemId = tPList(2).Key
            End If
            
'              Set the popup menu status
            If itemId = "CO" Then
               objTView.MenuStatus = ms_DELETE
            End If
                  
            If itemId <> "None" Then
'                 Now populate the propertis list from the relevant data object
               PopulatePropertyList objctrl.ClassCurrency(strArray(2))
            
'                 Highlight the selected item in the properties list
               tPList.Visible = False
               Set tPList.SelectedItem = tPList(itemId)
               tPList(itemId).EnsureVisible
               tPList.Pages(ediPr(itemId).PageKeys).Selected = True
               tPList.Visible = True
            End If
            
'              Now show the frame
            fView.Show Fra_EDI, "Configured Modules"
         
         Case "E"
            eClass.FurtherInfo = "Adding/amending new Result Mapping"

'              Set the properties list properties
            Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
            strArray = Split(objTView.NodeLevel(Node.Key), ":")
            nlevel = strArray(0)
            objctrl.PracticeId = strArray(1)
            objctrl.GroupKey = strArray(2)

'              Determine the key which correlates the treeview, The properties list and data objects
            itemId = objTView.NodeKey(Node.Key)
            If itemId = "New" Then
'                  objctrl.DataGroup = "EDI_InvTest_Codes"
               itemId = tPList(2).Key
            End If
            
'              Set the popup menu status
            If itemId = "RA" Then
               objTView.MenuStatus = ms_ADD
            ElseIf itemId = "IN1" Or itemId = "LC" Then
               objTView.MenuStatus = ms_DELETE
            End If
                  
            If itemId <> "None" Then
'                 Now populate the propertis list from the relevant data object
               PopulatePropertyList objctrl.ClassCurrency(strArray(2))
            
'                 Highlight the selected item in the properties list
               tPList.Visible = False
               Set tPList.SelectedItem = tPList(itemId)
               tPList(itemId).EnsureVisible
               tPList.Pages(ediPr(itemId).PageKeys).Selected = True
               tPList.Visible = True
            End If
            
'              Now show the frame
            fView.Show Fra_EDI, "Unit of Measure Mapping"
         
         Case "J"
            eClass.FurtherInfo = "Adding/amending new Result Mapping"

'              Set the properties list properties
            Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
            strArray = Split(objTView.NodeLevel(Node.Key), ":")
            nlevel = strArray(0)
            objctrl.PracticeId = strArray(1)
            objctrl.GroupKey = strArray(2)

'              Determine the key which correlates the treeview, The properties list and data objects
            itemId = objTView.NodeKey(Node.Key)
            If itemId = "New" Then
'                  objctrl.DataGroup = "EDI_InvTest_Codes"
               itemId = tPList(2).Key
            End If
            
'              Set the popup menu status
            If itemId = "CN" Then
               objTView.MenuStatus = ms_DELETE
            End If
            
            If itemId <> "None" Then
'                 Now populate the propertis list from the relevant data object
               PopulatePropertyList objctrl.ClassCurrency(strArray(2))
            
'                 Highlight the selected item in the properties list
               tPList.Visible = False
               Set tPList.SelectedItem = tPList(itemId)
               tPList(itemId).EnsureVisible
               tPList.Pages(ediPr(itemId).PageKeys).Selected = True
               tPList.Visible = True
            End If
            
'              Now show the frame
            fView.Show Fra_EDI, "Connection Details"
            
        Case "1J"
            'connections
'            Label6.Caption = "The treeview opposite shows all configured connections and their current status. Double click to 'New Connection' to add a new connection. Double click an entry to edit the settings. Click on the + to display/close the connection info"
            Command20.Visible = False
            Command21.Visible = False
            Command22.Visible = False
            Command23.Visible = False
            Command24.Visible = False
            CfgProgID = ""
            CfgUserID = ""
            CfgWardID = ""
            CfgCfgID = ""
            CfgName.Caption = ""
            Frame2.Visible = False
            If InStr(1, Node.Key, "-New") > 0 Then
                Frame2.Height = 5535
                Command18.Top = (Frame2.Height - Command19.Height - 60)
                Command19.Top = (Frame2.Height - Command19.Height - 60)
                Frame2.Visible = True
                ConfigList.PropertyItems.Clear
                ConfigList.PropertyItems.Add "Description", "Description", plpsString, "", "How you want the Connection to be known"
                ConfigList.PropertyItems.Add "Dir", "Direction", plpsList, "", "Data inbound or outbound"
                ConfigList.PropertyItems("Dir").ListItems.Add "Inbound", "I"
                ConfigList.PropertyItems("Dir").ListItems.Add "Outbound", "O"
                ConfigList.PropertyItems.Add "Frq", "Frequency", plpsString, "", "How often, in minutee, the connection should Poll - P = Permanent"
                ConfigList.PropertyItems.Add "SourceMethod", "Source Method", plpsList, "", "How is data collected"
                ConfigList.PropertyItems("SourceMethod").ListItems.Add "TCPIP", "T"
                ConfigList.PropertyItems("SourceMethod").ListItems.Add "FTP", "F"
                ConfigList.PropertyItems("SourceMethod").ListItems.Add "FOLDER", "D"
                ConfigList.PropertyItems("SourceMethod").ListItems.Add "PRINTER", "P"
                ConfigList.PropertyItems("SourceMethod").ListItems.Add "TERMINAL", "E"
                ConfigList.PropertyItems.Add "MethodParams", "Method Params", plpsString, "", "Scripts etc"
                ConfigList.PropertyItems.Add "ValType", "Validation Map", plpsList, "", "Map description to use to validate data"
                ConfigList.PropertyItems("ValType").ListItems.Add "HL7V2.3", "7"
                ConfigList.PropertyItems("ValType").ListItems.Add "NHS002", "2"
                ConfigList.PropertyItems("ValType").ListItems.Add "NHS003", "3"
                ConfigList.PropertyItems("ValType").ListItems.Add "NONE", "N"
                ConfigList.PropertyItems.Add "ACK", "Acknowledgements", plpsList, "", "Use an Acknowledgement method"
                ConfigList.PropertyItems("ACK").ListItems.Add "FILESCRIPT", "S"
                ConfigList.PropertyItems("ACK").ListItems.Add "HL7ACK", "A"
                ConfigList.PropertyItems("ACK").ListItems.Add "NONE", "N"
                ConfigList.PropertyItems.Add "InMap", "Inflight Mapping", plpsList, "", "Map Method to apply to incoming data"
                ConfigList.PropertyItems("InMap").ListItems.Add "HL7ASC", "7"
                ConfigList.PropertyItems("InMap").ListItems.Add "TP1PMIP", "T"
                ConfigList.PropertyItems("InMap").ListItems.Add "TP1toICE", "I"
                ConfigList.PropertyItems("InMap").ListItems.Add "NHS002", "2"
                ConfigList.PropertyItems("InMap").ListItems.Add "NHS003", "3"
                ConfigList.PropertyItems("InMap").ListItems.Add "ASTM1238", "A"
                ConfigList.PropertyItems("InMap").ListItems.Add "NHSPKISERVER", "K"
                ConfigList.PropertyItems.Add "TF", "Target Folder", plpsString, "", "Where to place the output data"
                ConfigList.PropertyItems.Add "TN", "Target Filemap", plpsString, "", "How to name the output files. Wildcards are acceptable"
                ConfigList.PropertyItems.Add "HF", "History Folder", plpsString, "", "Where to store archive data"
                ConfigList.PropertyItems.Add "PF", "Pending Folder", plpsString, "", "Where to hold data temporarily if problems with system"
                ConfigList.PropertyItems.Add "EF", "Error Folder", plpsString, "", "Where to hold files with data errors"
                ConfigList.PropertyItems.Add "LF", "Log Folder", plpsString, "", "Where to hold ASCI Representation of Audit Logs, for other systems"
                ConfigList.PropertyItems.Add "SF", "Special Filters", plpsString, "", "Special filter rules to apply to data processing"
                ConfigList.PropertyItems.Add "AC", "Active", plpsBoolean, "", "Whether connection is Active or not"
                ConfigList.Height = 4695
                CfgName.Caption = "New"
                ConfigPanel.Visible = True
            End If
            
            Case "K"
               eClass.FurtherInfo = "Adding/amending My Settings"

'              Set the properties list properties
               Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               nlevel = strArray(0)
               objctrl.PracticeId = strArray(1)

'              Determine the key which correlates the treeview, The properties list and data objects
               itemId = objTView.NodeKey(Node.Key)
               If itemId = "New" Then
                  itemId = tPList(1).Key
               End If
               
'              Now populate the propertis list from the relevant data object
               PopulatePropertyList objctrl.ClassCurrency(strArray(2))
               
'              Highlight the selected item in the properties list
               tPList.Visible = False
               Set tPList.SelectedItem = tPList(itemId)
               tPList(itemId).EnsureVisible
               tPList.Pages(ediPr(itemId).PageKeys).Selected = True
               tPList.Visible = True
               
'              Now show the frame
               fView.Show Fra_EDI, "Unit of Measure Mapping"

            Case "L"
               eClass.FurtherInfo = "Preparing log View "
                'Log entry
               LogBackColour = 12640511
               fView.WorkingFrame = Fra_LOGVIEW
               LogText.Text = ""
            
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               ntype = strArray(2)
               Select Case ntype
'                  NodeLevel 0 = Report
'                  NodeLevel 1 = Date
'                  NodeLevel 2 = File
'                  NodeLevel 3 = Tracking Entry or Report Entry
                  
                  Case "File" '  File Entry Clicked
                     Tstr = Node.Text
                     eClass.FurtherInfo = Tstr
                     RequeueStr = objTView.NodeKey(Node.Key)
                     Open Tstr For Input As #1
                     fData = Tstr
                     While Not EOF(1)
                        Line Input #1, Tstr
                        tBuf.Append Tstr
                        If Len(Tstr) > 0 Then
                           tBuf.Append vbCrLf
                        End If
                     Wend
                     LogText.Text = tBuf.value
                     Close (1)
                     If objTView.GetParentData(Node.Key, nl_ROOT, True) = "EDI Reports" Then
                        MailBtn.Caption = "RESEND EDI FILE"
                     Else
                        MailBtn.Caption = "REQUEUE ALL REPORTS"
                     End If
                     MailBtn.Visible = True
                     
                  Case "Tracking"  '  Tracking or report
                     Tstr = objTView.NodeKey(Node.Key)
                     RS.Open "Select * From Service_ImpExp_Comments Where Service_ImpExp_Id = " & Val(Tstr) & " order by service_impexp_comment_id", ICECon, adOpenKeyset, adLockReadOnly
                     If RS.RecordCount > 0 Then
                        Do While Not RS.EOF
                           LogText.Text = LogText.Text & vbCrLf & Trim(RS!Service_ImpExp_comment)
                           RS.MoveNext
                        Loop
                     End If
                     RS.Close
                     MailBtn.Caption = ""
                     MailBtn.Visible = False
                        
                  Case "Report"
                     MailBtn.Caption = "REQUEUE THIS REPORT"
                     MailBtn.Visible = True
                     strArray = Split(objTView.NodeKey(Node.Key), "+")
                     OrgStr = strArray(0)
                     IDStr = strArray(1)
                     RequeueStr = strArray(2)
                     Call Load_Report(OrgStr, IDStr)
                     
                  Case Else
                     fView.Show Fra_HELP, ""
                     
               End Select
                
                If fView.WorkingFrame <> 10 Then
                  fView.Show Fra_LOGVIEW, fData
               End If
               
            
            Case "M"
               eClass.FurtherInfo = "Adding/amending new Unit of Measure"

'              Set the properties list properties
               Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               nlevel = strArray(0)
               objctrl.PracticeId = strArray(1)

'              Determine the key which correlates the treeview, The properties list and data objects
               itemId = objTView.NodeKey(Node.Key)
               If itemId = "New" Then
                  itemId = tPList(1).Key
               End If
               
'              Set the popup menu status
               If itemId = "LocUOM" Then
                  objTView.MenuStatus = ms_DELETE
               End If
               
'              Now populate the propertis list from the relevant data object
               PopulatePropertyList objctrl.ClassCurrency(strArray(2))
               
'              Highlight the selected item in the properties list
               tPList.Visible = False
               Set tPList.SelectedItem = tPList(itemId)
               tPList(itemId).EnsureVisible
               tPList.Pages(ediPr(itemId).PageKeys).Selected = True
               tPList.Visible = True
               
'              Now show the frame
               fView.Show Fra_EDI, "Unit of Measure Mapping"
            
            Case "N"
               eClass.FurtherInfo = "Adding/amending Korner Codes"

'              Set the properties list properties
               Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               nlevel = strArray(0)
               objctrl.PracticeId = strArray(1)

'              Determine the key which correlates the treeview, The properties list and data objects
               itemId = objTView.NodeKey(Node.Key)
               If itemId = "New" Then
                  itemId = tPList(1).Key
               End If
               
               If itemId <> "None" Then
'                 Now populate the propertis list from the relevant data object
                  PopulatePropertyList objctrl.ClassCurrency(strArray(2))
               
'                 Highlight the selected item in the properties list
                  tPList.Visible = False
                  Set tPList.SelectedItem = tPList(itemId)
                  tPList(itemId).EnsureVisible
                  tPList.Pages(ediPr(itemId).PageKeys).Selected = True
                  tPList.Visible = True
               End If
'              Now show the frame
               fView.Show Fra_EDI, "Specimen Code Rubrics"
            
         Case "P"
            eClass.FurtherInfo = "Adding/amending new Result Mapping"

'           Set the properties list properties
            Set tPList = objTView.PreparePropertiesList(nOrigin)

'           Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
            strArray = Split(objTView.NodeLevel(Node.Key), ":")
            nlevel = strArray(0)
            objctrl.PracticeId = strArray(1)
            objctrl.GroupKey = strArray(2)

'           Determine the key which correlates the treeview, The properties list and data objects
            itemId = objTView.NodeKey(Node.Key)
            If itemId = "New" Then
'                  objctrl.DataGroup = "EDI_InvTest_Codes"
               itemId = tPList(2).Key
            End If
            
'           Set the popup menu status
            If itemId = "MN" Then
               objTView.MenuStatus = ms_DELETE
            End If
                  
            If itemId <> "None" Then
'                 Now populate the propertis list from the relevant data object
               PopulatePropertyList objctrl.ClassCurrency(strArray(2))
            
'                 Highlight the selected item in the properties list
               tPList.Visible = False
               Set tPList.SelectedItem = tPList(itemId)
               tPList(itemId).EnsureVisible
               tPList.Pages(ediPr(itemId).PageKeys).Selected = True
               tPList.Visible = True
            End If
            
'              Now show the frame
            fView.Show Fra_EDI, "Connection Details"
            
            Case "S"
               eClass.FurtherInfo = "Adding/amending Sample Code Mapping"

'              Set the properties list properties
               Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               nlevel = strArray(0)
               objctrl.PracticeId = strArray(1)

'              Determine the key which correlates the treeview, The properties list and data objects
               itemId = objTView.NodeKey(Node.Key)
               If itemId = "New" Then
                  itemId = tPList(1).Key
               End If
               
'              Set the popup menu status
               If itemId = "SA" Or itemId = "SC" Or itemId = "ST" Then
                  objTView.MenuStatus = ms_DELETE
               End If
               
               If itemId <> "None" Then
'                 Now populate the propertis list from the relevant data object
                  PopulatePropertyList objctrl.ClassCurrency(strArray(2))
               
'                Highlight the selected item in the properties list
                  tPList.Visible = False
                  Set tPList.SelectedItem = tPList(itemId)
                  tPList(itemId).EnsureVisible
                  tPList.Pages(ediPr(itemId).PageKeys).Selected = True
                  tPList.Visible = True
               End If
'              Now show the frame
               fView.Show Fra_EDI, "Specimen Code Rubrics"
            
            Case "U"
               eClass.FurtherInfo = "Adding/amending Recipients"

'              Set the properties list properties
               Set tPList = objTView.PreparePropertiesList(nOrigin)

'              Ascertain the Tree Level (strArray(0)), the DataGroup (strArray(1)) and the DataItems (strArray(2))
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               nlevel = strArray(0)
               objctrl.PracticeId = strArray(1)

'              Determine the key which correlates the treeview, The properties list and data objects
               itemId = objTView.NodeKey(Node.Key)
               If itemId = "New" Then
                  itemId = tPList(1).Key
                  objctrl.GroupKey = "New"
               End If
               
'              Set the popup menu status
               If itemId = "Spec" Or itemId = "Msg" Or itemId = "Ind" Then
                  objTView.MenuStatus = ms_ADD
               ElseIf itemId = "SP" Or itemId = "MS" Or itemId = "IN" Or itemId = "NatId" Then
                  objTView.MenuStatus = ms_DELETE
               End If
               
'               Do we need to change the display for this item?
               If itemId <> "None" Then
'                 Populate the propertis list from the relevant data object
                  PopulatePropertyList objctrl.ClassCurrency(strArray(2))
               
'                 Highlight the selected item in the properties list
                  tPList.Visible = False
                  Set tPList.SelectedItem = tPList(itemId)
                  tPList(itemId).EnsureVisible
                  tPList.Pages(ediPr(itemId).PageKeys).Selected = True
                  tPList.Visible = True
               End If
               
'              Now show the frame
               fView.Show Fra_EDI, "EDi Recipients"

            Case "Z"
               eClass.FurtherInfo = "System Monitor"
               ReDim Activity(0 To 0)
               ReDim ActivityDirs(0 To 0)
               ReDim ActivityDirPattern(0 To 0)
               For i = 0 To FLActivity.UBound
                  FLActivity(i).Visible = False
                  If i <= FLlbl.UBound Then
                     FLlbl(i).Visible = False
                  End If
               Next i
               strArray = Split(objTView.NodeLevel(Node.Key), ":")
               IDStr = strArray(2)
               itemId = objTView.NodeKey(Node.Key)
               
               If IDStr = "All" Then   '  All categories
                  Set TVN = Node.Child.FirstSibling
                  For i = 1 To Node.Children
                     Set TVN2 = TVN.Child.FirstSibling
                     For P = 1 To TVN.Children
                        If objTView.NodeKey(TVN2.Key) = "FLD" Then
                           Tstr = Trim(TVN2.Text)
                           For E = Len(Tstr) To 1 Step -1
                              If Mid$(Tstr, E, 1) = "\" Then
                                 ReDim Preserve Activity(0 To UBound(Activity) + 1)
                                 ReDim Preserve ActivityDirs(0 To UBound(ActivityDirs) + 1)
                                 ReDim Preserve ActivityDirPattern(0 To UBound(ActivityDirPattern) + 1)
                                 Activity(UBound(Activity)) = Trim(TVN.Text) ' & Str(P)
                                 ActivityDirPattern(UBound(ActivityDirPattern)) = Mid$(Tstr, E + 1, Len(Tstr) - E)
                                 ActivityDirs(UBound(ActivityDirs)) = Mid$(Tstr, 1, E - 1)
                                 Exit For
                              End If
                           Next E
                        End If
                        Set TVN2 = TVN2.Next
                     Next P
                     Set TVN = TVN.Next
                  Next i
                  Display_DirActivity
               
               Else  '  Determine which node has been clicked
                  Select Case itemId
                     Case "FLD"  '  Connection folder
                        Tstr = Node.Text
                        For E = Len(Tstr) To 1 Step -1
                           If Mid$(Tstr, E, 1) = "\" Then
                               ReDim Preserve Activity(0 To UBound(Activity) + 1)
                               ReDim Preserve ActivityDirs(0 To UBound(ActivityDirs) + 1)
                               ReDim Preserve ActivityDirPattern(0 To UBound(ActivityDirPattern) + 1)
                               Activity(UBound(Activity)) = IDStr
                               ActivityDirPattern(UBound(ActivityDirPattern)) = Mid$(Tstr, E + 1, Len(Tstr) - E)
                               ActivityDirs(UBound(ActivityDirs)) = Mid$(Tstr, 1, E - 1)
                               E = 1
                           End If
                        Next E
                        Display_DirActivity
                     
                     Case "CON"  '  Connection type node
                        Set TVN = Node.Child.FirstSibling
                        For i = 1 To Node.Children
                           Tstr = Trim(TVN.Text)
                           If objTView.NodeKey(TVN.Key) = "FLD" Then
                              For E = Len(Tstr) To 1 Step -1
                                 If Mid$(Tstr, E, 1) = "\" Then
                                    ReDim Preserve Activity(0 To UBound(Activity) + 1)
                                    ReDim Preserve ActivityDirs(0 To UBound(ActivityDirs) + 1)
                                    ReDim Preserve ActivityDirPattern(0 To UBound(ActivityDirPattern) + 1)
                                    Activity(UBound(Activity)) = Trim(Node.Text) & UBound(Activity)
                                    ActivityDirPattern(UBound(ActivityDirPattern)) = Mid$(Tstr, E + 1, Len(Tstr) - E)
                                    ActivityDirs(UBound(ActivityDirs)) = Mid$(Tstr, 1, E - 1)
                                    E = 1
                                 End If
                              Next E
                           End If
                           Set TVN = TVN.Next
                        Next i
                        Display_DirActivity
                     
                     Case Else   '  Connection Port
                  End Select
               
               End If
               
               fView.Show Fra_FILEDETAILS, "System Activity"
                  
    End Select
    frmMain.MousePointer = 0
    Exit Sub
    
ProcEH:
   eClass.CurrentProcedure = "frmMain.NodeClick"
   errRet = eClass.Add(Err.Number, Err.Description, Err.Source)
   If errRet = 0 Then
      frmMain.MousePointer = vbNormal
      Exit Sub
   ElseIf errRet = -1 Then
      Stop
      Resume
   Else
      eClass.Show
   End If
'
'ErrorHandler:
'
'   If eClass.LogError(Err.Number, Err.Description, "frmMain.Treeview1_NodeClick", "") = -1 Then
'      MsgBox "File: " & Tstr & vbCrLf & vbCrLf & "not found on this system. Please check the configuration.", vbExclamation, "ICEConfig - Non critical eror"
'      frmMain.MousePointer = vbNormal
'      Exit Sub
'   Else
'      End
'   End If
   
End Sub

Private Sub DebugPL()
   Dim i As Integer
   For i = 1 To ediPr.PropertyItems.Count
      Debug.Print "[" & ediPr(i).Caption & "] has value of [" & ediPr(i).value & "] of type [" & ediPr(i).Style & "]"
   Next i
End Sub

Public Sub PopulatePropertyList(DataKey As GenericCollClass)

   Dim vData As Variant
   Dim i As Integer
'   Dim objctrl As Class1
   
   vData = objctrl.ReadData(DataKey)
'   vData = objctrl.ReadClassData
   For i = 0 To UBound(vData) - 1 Step 2
      If Left(vData(i), 2) <> "pv" And Left(vData(i), 2) <> "hv" Then
         If tPList(vData(i)).Style = plpsBoolean Then
            tPList(vData(i)).value = CBool(vData(i + 1))
         Else
            tPList(vData(i)).value = vData(i + 1)
         End If
      End If
   Next i
   

End Sub

Private Sub SetUpResultMapping(LocTestCode As String, sexData As String)

   Dim RS As New ADODB.Recordset
   
   If sexData = "*" Then
      RS.Open "SELECT * FROM EDI_InvTest_Codes WHERE Organisation LIKE '" & OrgList.Text & "' AND EDI_Local_Test_Code LIKE '" & LocTestCode & "'", ICECon, adOpenKeyset, adLockReadOnly
'      ediPr("Code").Value = LocTestCode
      ediPr("LC").value = RS("EDI_Local_Test_Code")
      ediPr("LD").value = RS("EDI_Local_Rubric")
      ediPr("RC").value = RS("EDI_READ_Code")
      ediPr("SC").value = RS("EDI_Sample_TypeCode")
      ediPr("SQ").value = RS("EDI_OP_Seq")
      ediPr("UM").value = RS("EDI_OP_UOM")
   Else
      RS.Open "SELECT * FROM EDI_InvTest_Ranges WHERE EDI_InvTest_Code = '" & LocTestCode & "' AND EDI_Range_Sex LIKE '" & sexData & "'", ICECon, adOpenKeyset, adLockReadOnly
      If RS.BOF = False And RS.EOF = False Then
         ediPr("RA").value = LocTestCode
         ediPr("RA-IN1").value = RS("EDI_Range_Sex")
         ediPr("RA-IN2").value = RS("EDI_Range_MinAge")
         ediPr("RA-IN3").value = RS("EDI_Range_MaxAge")
         ediPr("RA-IN4").value = RS("EDI_Range_Lo")
         ediPr("RA-IN5").value = RS("EDI_Range_Hi")
         ediPr("RA-IN6").value = RS("EDI_Range_UOM")
         ediPr("RA-IN7").value = RS("EDI_Range_Comment")
      End If
   End If
   RS.Close
   Set RS = Nothing
'   edipr.Visible = True
   
End Sub
Private Sub SetUpEDIRecipients(NationalCode As String, nodeData As String, QueryData As String)

   Dim RS As New ADODB.Recordset
   Dim RS1 As New ADODB.Recordset
   Dim idx As String
   Dim pageId As String
   
   
   
   Select Case Left(nodeData, 2)
      Case "SP"
         RS.Open "SELECT * FROM EDI_Loc_Specialties " & _
                        "WHERE EDI_Nat_Code = '" & NationalCode & "' AND EDI_Korner_Code= '" & QueryData & "'", _
                        ICECon, adOpenKeyset, adLockReadOnly
         idx = "SP"
         If RS.BOF And RS.EOF Then
            ediPr(idx).Caption = "No Specialty selected"
            ediPr(idx).value = ""
         Else
'           Set up the header to allow for the title node being clicked
            ediPr(idx).Caption = RS("EDI_Korner_Code")
            ediPr(idx).value = RS("EDI_Specialty")
            ediPr(idx).Bold = True
            ediPr(idx).ReadOnly = True
      
'           Now set up the rest of the data for this page
            ediPr("SP+MS1").value = uCtrl.ReadClassData("SP+MS4")  '   Trim(RS("EDI_Korner_Code"))
            ediPr("SP+MS2").value = uCtrl.ReadClassData("SP+MS4")  '   Trim(RS("EDI_Specialty"))
            ediPr("SP+MS3").value = uCtrl.ReadClassData("SP+MS4")  '   Trim(RS("EDI_Msg_Format"))
            ediPr("SP+MS4").value = uCtrl.ReadClassData("SP+MS4")  '   RS("EDI_Msg_Active")
         End If
   
      Case "MS"
'         Add message type page(s)
         RS.Open "SELECT * FROM EDI_Msg_Types WHERE EDI_Org_NatCode = '" & NationalCode & "' AND  EDI_Msg_Format = '" & QueryData & "'", ICECon, adOpenKeyset, adLockReadOnly
         idx = "MS"
         If RS.BOF And RS.EOF Then
            ediPr(idx).Caption = "No Message Type Selected"
         Else
'           Set up the header to allow for the title node being clicked
            ediPr(idx).Caption = "Message Type"
            ediPr(idx).value = RS("EDI_Msg_Format")
            ediPr(idx).Bold = True
            ediPr(idx).ReadOnly = True
   
   '        Now set up the rest of the data for this page
            ediPr("MS+MS1").value = Trim(RS("EDI_Msg_Format"))
            ediPr("MS+MS4").value = Trim(RS("EDI_Delivery_Method"))
            ediPr("MS+MS6").value = RS("EDI_Encrypt_Enabled")
            ediPr("MS+MS7").value = RS("EDI_Acks_Active")
            ediPr("MS+MS8").value = RS("EDI_Msg_Test")
            ediPr("MS+MS9").value = RS("EDI_Msg_Active")
         End If
      
      Case "IN"
'         Add Individual page(s)
         RS.Open "SELECT * FROM EDI_Recipient_Individuals WHERE EDI_Org_NatCode = '" & NationalCode & "' AND EDI_Local_Key3 = '" & QueryData & "'", ICECon, adOpenKeyset, adLockReadOnly
         idx = "IN"
         If RS.BOF And RS.EOF Then
            ediPr("IN").Caption = "No Individual Selected"
            ediPr(idx).value = ""
         Else
'           Set up the header to allow for the title node being clicked
            ediPr(idx).Caption = RS("EDI_Local_Key3")
            ediPr(idx).value = RS("EDI_OP_Name")
            ediPr(idx).Bold = True
            ediPr(idx).ReadOnly = True
         
   '        Now set up the rest of the data for this page
            ediPr("IN+IN1").value = Trim(RS("EDI_Local_Key1"))
            ediPr("IN+IN2").value = Trim(RS("EDI_Local_Key2"))
            ediPr("IN+IN3").value = RS("EDI_Local_Key3")
            ediPr("IN+IN4").value = Trim(RS("EDI_OP_Name"))
            ediPr("IN+IN5").value = Trim(RS("EDI_NatCode"))
            ediPr("IN+IN6").value = RS("EDI_Active")
         End If
         
      Case Else
   
      '  General details
         RS.Open "SELECT * FROM EDI_Recipients WHERE EDI_NatCode = '" & NationalCode & "'", ICECon, adOpenKeyset, adLockReadOnly
                  
'        Set up the header to allow for the title node being clicked
         idx = "NatId"
         ediPr(idx).value = Trim(RS("EDI_NatCode")) & " - " & Trim(RS("Edi_Name"))
         ediPr(idx).Bold = True
         ediPr(idx).ReadOnly = True
               
'        Now set up the rest of the data for this page
         ediPr("NC").value = Trim(RS("EDI_NatCode"))
         ediPr("LC").value = Trim(RS("EDI_LocalCode"))
         ediPr("ER").value = Trim(RS("EDI_Trader_Account"))
         ediPr("FP").value = Trim(RS("EDI_Free_Part"))
         ediPr("EN").value = RS("EDI_Encryption")
         ediPr("NA").value = Trim(RS("EDI_Name"))
         ediPr("AD").value = Trim(RS("EDI_Address"))
         ediPr("SM").value = Trim(RS("EDI_SMTP_Mail"))
         ediPr("AC").value = RS("EDI_Active")
         
      '  Add X400 details
         
      '     Set up the header to allow for the title node being clicked
         idx = "X4"
         ediPr(idx).value = RS("EDI_X400_GivenName")
         ediPr(idx).Bold = True
         ediPr(idx).ReadOnly = True
               
      '      Now set up the rest of the data for this page
         ediPr("X4A").value = Trim(RS("EDI_X400_GivenName"))
         ediPr("X4B").value = Trim(RS("EDI_X400_Surname"))
         ediPr("X4C").value = Trim(RS("EDI_X400_Initials"))
         ediPr("X4D").value = Trim(RS("EDI_X400_Generation"))
         ediPr("X4E").value = Trim(RS("EDI_X400_Common"))
         ediPr("X41").value = Trim(RS("EDI_X400_Org"))
         ediPr("X42").value = Trim(RS("EDI_X400_OU1"))
         ediPr("X43").value = Trim(RS("EDI_X400_OU2"))
         ediPr("X44").value = Trim(RS("EDI_X400_OU3"))
         ediPr("X45").value = Trim(RS("EDI_X400_OU4"))
         ediPr("X46").value = Trim(RS("EDI_X400_PRD"))
         ediPr("X47").value = Trim(RS("EDI_X400_Adm"))
         ediPr("X48").value = Trim(RS("EDI_X400_c"))
   
   End Select
   
   RS.Close
   Set RS = Nothing

End Sub

Private Sub SetupConfigList()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    ConfigList.PropertyItems.Clear
    If CfgWardID = "" And CfgUserID = "" Then
        RS.Open "Select CfgType,CfgValue,CfgNotes From Configuration Where Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "'", ICECon, adOpenKeyset, adLockReadOnly
        ConfigList.PropertyItems.Add "Value", "Value", Val(RS!CfgType), RS!CfgValue & "", RS!CfgNotes & ""
    ElseIf CfgWardID <> "" Then
        RS.Open "Select CfgType,CfgValue,CfgNotes From Configuration Where Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "' And Location='" & CfgWardID & "'", ICECon, adOpenKeyset, adLockReadOnly
        ConfigList.PropertyItems.Add "Value", "Value", Val(RS!CfgType), RS!CfgValue & "", RS!CfgNotes & ""
    ElseIf CfgUserID <> "" Then
        RS.Open "Select CfgType,CfgValue,CfgNotes From Configuration Where Organisation='" & OrgList.Text & "' And CfgID='" & CfgCfgID & "' And ProgramID='" & CfgProgID & "' And Username='" & CfgUserID & "'", ICECon, adOpenKeyset, adLockReadOnly
        ConfigList.PropertyItems.Add "Value", "Value", Val(RS!CfgType), RS!CfgValue & "", RS!CfgNotes & ""
    End If
End Sub
Private Sub SetupTestDetailsList()
    TD1.View = plvCategorised
    TD1.PropertyItems.Clear
    TD1.PropertyItems.Add "GENERAL", "General", , , , , , pldsHeader
    TD1.PropertyItems.Add "TEST_CODE", "Test Code", plpsString, "", "Test code as required by laboratory system", , , pldsNormal
    TD1.PropertyItems("TEST_CODE").Max = 20
    TD1.PropertyItems.Add "DEPT", "Department", plpsString, "", "Laboratory department which handles test", , , pldsNormal
    TD1.PropertyItems("DEPT").Max = 65
    TD1.PropertyItems.Add "SCREEN_PANEL", "Screen Panel", plpsList, "", "Panel on which test appears"
    TD1.PropertyItems.Add "PANEL_PAGE", "Panel Page", plpsList, "", "Page on panel on which test appears"
    TD1.PropertyItems.Add "SCREEN_POSN", "Screen Position", plpsList, 1, "Position in which test is displayed on screen. ", , , pldsNormal
    'GetVacantScreenPositions
    'TD1.PropertyItems.Add "SCREEN_POSN", "Screen Position", plpsNumber, 1, "Position in which test is displayed on screen", , , pldsNormal
    'TD1.PropertyItems("SCREEN_POSN").Increment = 1
    'TD1.PropertyItems("SCREEN_POSN").Min = 1
    'TD1.PropertyItems("SCREEN_POSN").Max = 450
    TD1.PropertyItems.Add "SCREEN_CAPTION", "Screen Caption", plpsString, , "Caption displayed on screen for user selection"
    TD1.PropertyItems("SCREEN_CAPTION").Max = 35
    TD1.PropertyItems.Add "SCREEN_HELP", "Screen Help", plpsString, , "Help message displayed on screen upon user selection of test"
    TD1.PropertyItems("SCREEN_HELP").Max = 140
    TD1.PropertyItems.Add "SCREEN_COLOUR", "Screen Colour", plpsColor, , "Colour of test text on screen"
    TD1.PropertyItems.Add "HELP_COLOUR", "Help Colour", plpsColor, , "Colour of background behind help message"
    TD1.PropertyItems.Add "INFO", "Information", plpsString, , "Information about the test"
    TD1.PropertyItems("INFO").Max = 50
    TD1.PropertyItems.Add "PROV_ID", "Provider ID", plpsCustom, , "Which lab handles this test"
    TD1.PropertyItems.Add "READ_CODE", "Read Code", plpsString, , "READ Code assigned to this test"
    TD1.PropertyItems("READ_CODE").Max = 5
    TD1.PropertyItems.Add "TUBE_CODE", "Sample Tube", plpsColor, , "Sample container required for this test"
    TD1.PropertyItems.Add "PAED_TUBE_CODE", "Paediatric Tube", plpsColor, , "Sample container required for this test when carried out on paediatric patient"
    TD1.PropertyItems.Add "ENABLED", "Enabled", plpsBoolean, True, "Can users request this test"
    TD1.PropertyItems.Add "VOL", "Sample Volume Required", plpsNumber, 0, "Amount of sample required to perform test"
    TD1.PropertyItems("VOL").Max = 5000
    TD1.PropertyItems("VOL").Min = 0
    TD1.PropertyItems("VOL").Increment = 5
    TD1.PropertyItems.Add "PAED_VOL", "Paediatric Volume Required", plpsNumber, 0, "Amount of sample required to perform test on a paediatric patient"
    TD1.PropertyItems("PAED_VOL").Max = 5000
    TD1.PropertyItems("PAED_VOL").Min = 0
    TD1.PropertyItems("PAED_VOL").Increment = 5
    TD1.PropertyItems.Add "SA", "Special Attributes", , , , , , pldsHeader
    TD1.PropertyItems("SA").Expanded = False
    TD1.PropertyItems.Add "DYNAMIC", "Dynamic", plpsBoolean, False, "Dynamic"
    TD1.PropertyItems.Add "SENSITIVE", "Sensitive", plpsBoolean, False, "Sensitive"
    TD1.PropertyItems.Add "WORKLIST", "Worklist Enabled", plpsBoolean, False, "Worklist enabled"
    TD1.PropertyItems.Add "DPR", "Display Previous Results", , , , , , pldsHeader
    TD1.PropertyItems("DPR").Expanded = False
    TD1.PropertyItems.Add "HIST_ST", "History Search Type", plpsList, , "Select type of search you wish to do to show previous results prior to requesting this test"
    TD1.PropertyItems("HIST_ST").ListItems.Add "None", "NO"
    TD1.PropertyItems("HIST_ST").ListItems.Add "Results", "SR"
    TD1.PropertyItems("HIST_ST").ListItems.Add "Investigations", "SI"
    TD1.PropertyItems.Add "HIST_SS", "History Search String", plpsString, , "Enter string you wish to search for (only valid if above option not set to None)"
    TD1.PropertyItems("HIST_SS").Max = 60
    TD1.PropertyItems.Add "HIST_MSG", "History Message", plpsString, , "Enter message to display to users when performing search"
    TD1.PropertyItems("HIST_MSG").Max = 255
End Sub

Private Sub LoadTestDetails(idx)
    Dim RS As ADODB.Recordset
    Dim RS2 As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS2 = New ADODB.Recordset
    RS.Open "Select * From Request_Tests Where Test_Index=" & idx, ICECon, adOpenKeyset, adLockReadOnly
    TD1.PropertyItems("TEST_CODE").value = Mid$(RS!Test_Code & "", 7, Len(RS!Test_Code) - 6)
    TD1.PropertyItems("DEPT").value = RS!Department & ""
    RS2.Open "Select PanelID,PanelName From Request_Panels Where PanelType=1 Order By PanelID", ICECon, adOpenKeyset, adLockReadOnly
    If RS2.RecordCount > 0 Then
        Do While Not RS2.EOF
            TD1.PropertyItems("SCREEN_PANEL").ListItems.Add RS2!PanelName, Format(RS2!PanelID)
            RS2.MoveNext
        Loop
    End If
    RS2.Close
    TD1.PropertyItems("SCREEN_PANEL").value = Format(RS!Screen_Panel)
    RS2.Open "Select PageName From Request_Panels_Pages Where PanelID=" & RS!Screen_Panel & "Order By PageName", ICECon, adOpenKeyset, adLockReadOnly
    If RS2.RecordCount > 0 Then
        Do While Not RS2.EOF
            TD1.PropertyItems("PANEL_PAGE").ListItems.Add RS2!PageName, RS2!PageName
            RS2.MoveNext
        Loop
    End If
    RS2.Close
    TD1.PropertyItems("PANEL_PAGE").value = Format(RS!Screen_Panel_Page)
    TD1.ShowListValues = False
    TD1.PropertyItems("SCREEN_POSN").ListItems.Add Format(RS!Screen_Position), Format(RS!Screen_Position)
    TD1.PropertyItems("SCREEN_POSN").value = Format(RS!Screen_Position)
    TD1.PropertyItems("SCREEN_CAPTION").value = RS!Screen_Caption
    TD1.PropertyItems("SCREEN_HELP").value = RS!Screen_Help
    If Format(RS!Screen_Colour) & "" <> "" Then
        RS2.Open "Select Colour_Code,Colour_Name From Colours Where Colour_Index=" & RS!Screen_Colour, ICECon, adOpenKeyset, adLockReadOnly
        TD1.PropertyItems("SCREEN_COLOUR").value = Val(RS2!Colour_Code)
        TD1.PropertyItems("SCREEN_COLOUR").Tag = Format(RS!Screen_Colour) + "-" + RS2!Colour_Name
        RS2.Close
    End If
    If Format(RS!Screen_Help_Backcolour) & "" <> "" Then
        RS2.Open "Select Colour_Code,Colour_Name From Colours Where Colour_Index=" & RS!Screen_Help_Backcolour, ICECon, adOpenKeyset, adLockReadOnly
        TD1.PropertyItems("HELP_COLOUR").value = Val(RS2!Colour_Code)
        TD1.PropertyItems("HELP_COLOUR").Tag = Format(RS!Screen_Help_Backcolour) + "-" + RS2!Colour_Name
        RS2.Close
    End If
    TD1.PropertyItems("INFO").value = RS!Information
    If Format(RS!Provider_ID) & "" <> "" Then
        RS2.Open "Select Provider_Name From Service_Providers Where Provider_ID=" & RS!Provider_ID, ICECon, adOpenKeyset, adLockReadOnly
        TD1.PropertyItems("PROV_ID").value = RS2!Provider_Name & ""
        TD1.PropertyItems("PROV_ID").Tag = Format(RS!Provider_ID)
        RS2.Close
    End If
    TD1.PropertyItems("READ_CODE").value = RS!Read_Code
    If Format(RS!Tube_Code) & "" <> "" Then
        RS2.Open "Select Name,Colour,Description From Request_Tubes Where Tube_Index=" & RS!Tube_Code, ICECon, adOpenKeyset, adLockReadOnly
        TC = RS2!Colour
        TD1.PropertyItems("TUBE_CODE").Tag = Format(RS!Tube_Code) + "-" + RS2!Name + " - " + Format(RS2!Description)
        RS2.Close
        If Format(TC) & "" <> "" Then
            RS2.Open "Select Colour_Code From Colours Where Colour_Index=" & TC, ICECon, adOpenKeyset, adLockReadOnly
            TD1.PropertyItems("TUBE_CODE").value = Val(RS2!Colour_Code)
            RS2.Close
        End If
    End If
    If Format(RS!PaedTube_Code) & "" <> "" Then
        RS2.Open "Select Name,Colour,Description From Request_Tubes Where Tube_Index=" & RS!PaedTube_Code, ICECon, adOpenKeyset, adLockReadOnly
        TC = RS2!Colour
        TD1.PropertyItems("PAED_TUBE_CODE").Tag = Format(RS!PaedTube_Code) + "-" + RS2!Name + " - " + Format(RS2!Description)
        RS2.Close
        If Format(TC) & "" <> "" Then
            RS2.Open "Select Colour_Code From Colours Where Colour_Index=" & TC, ICECon, adOpenKeyset, adLockReadOnly
            TD1.PropertyItems("PAED_TUBE_CODE").value = Val(RS2!Colour_Code)
            RS2.Close
        End If
    End If
    TD1.PropertyItems("ENABLED").value = RS!Enabled
    If Format(RS!Test_Volume) & "" <> "" Then
        TD1.PropertyItems("VOL").value = RS!Test_Volume
    Else
        TD1.PropertyItems("VOL").value = 0
    End If
    If Format(RS!PaedTube_Test_Volume) & "" <> "" Then
        TD1.PropertyItems("PAED_VOL").value = RS!PaedTube_Test_Volume
    Else
        TD1.PropertyItems("PAED_VOL").value = 0
    End If
    TD1.PropertyItems("DYNAMIC").value = RS!Dynamic
    TD1.PropertyItems("SENSITIVE").value = RS!Sensitive
    TD1.PropertyItems("WORKLIST").value = RS!Worklist_Enabled
    If Format(RS!ResHistory_SearchType) & "" <> "" Then
        TD1.PropertyItems("HIST_ST").value = Format(RS!ResHistory_SearchType)
    Else
        TD1.PropertyItems("HIST_ST").value = "NO"
    End If
    TD1.PropertyItems("HIST_SS").value = RS!ResHistory_SearchString
    TD1.PropertyItems("HIST_MSG").value = RS!ResHistory_Message
End Sub

Private Sub WriteTestRules(N)
    If N.Children = 0 Then Exit Sub
    Dim TVN As Node
    Dim SeqNo As Integer
    Set TVN = N.Child
    SeqNo = 1
    TempStr = "Insert Into Request_Test_Prompts (Test_Index,Prompt_Index,Sequence) Values ("
    TempStr2 = Mid$(TVN.Key, 2, Len(TVN.Key) - 1)
    PI = Left$(TempStr2, InStr(1, TempStr2, "-") - 1)
    TI = Mid$(TempStr2, InStr(1, TempStr2, "-") + 1, Len(TempStr2) - InStr(1, TempStr2, "-") + 1)
    ICECon.Execute "Delete From Request_Test_Prompts Where Test_Index=" & TI
    ICECon.Execute TempStr + TI + "," + PI + "," + Format(SeqNo) + ")"
    For i = 1 To N.Children - 1
        SeqNo = SeqNo + 1
        Set TVN = TVN.Next
        TempStr2 = Mid$(TVN.Key, 2, Len(TVN.Key) - 1)
        PI = Left$(TempStr2, InStr(1, TempStr2, "-") - 1)
        TI = Mid$(TempStr2, InStr(1, TempStr2, "-") + 1, Len(TempStr2) - InStr(1, TempStr2, "-") + 1)
        ICECon.Execute TempStr + TI + "," + PI + "," + Format(SeqNo) + ")"
    Next i
End Sub


Private Sub SetupProfileList()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    PR1.PropertyItems.Clear
    PR1.PropertyItems.Add "WARD", "Ward", plpsList, , "Select the ward to which this test profile is available"
    RS.Open "Select Clinic_Name,Localation_Code From Location Where Localation_Code Like '" & OrgList.Text & "%' Order By Clinic_Name", ICECon, adOpenKeyset, adLockReadOnly
    Do While Not RS.EOF
        PR1.PropertyItems("WARD").ListItems.Add RS!Clinic_Name, RS!Localation_Code
        RS.MoveNext
    Loop
    RS.Close
    PR1.PropertyItems.Add "ENABLED", "Enabled", plpsBoolean, False, "Is this profile enabled"
    PR1.PropertyItems.Add "PROFILE_COLOUR", "Profile Colour", plpsColor, , "Select the colour of the text on the profile button"
    PR1.PropertyItems.Add "POSITION", "Position", plpsNumber, 1, "Enter the position that the profile appears in"
    PR1.PropertyItems("POSITION").Max = 10
    PR1.PropertyItems("POSITION").Min = 1
    PR1.PropertyItems("POSITION").Increment = 1
    PR1.PropertyItems.Add "CAPTION", "Caption", plpsString, , "Enter the caption that is displayed to the user"
    PR1.PropertyItems("CAPTION").Max = 25
    PR1.PropertyItems.Add "HELP", "Help Text", plpsString, , "Enter the help detail that is displayed on screen when the user selects the profile"
    PR1.PropertyItems("HELP").Max = 65
    PR1.PropertyItems.Add "HELP_COLOUR", "Help Colour", plpsColor, , "Select the background colour of the help string"
    PR1.PropertyItems.Add "TEXT_STRING", "Text String", plpsString, , "Enter the text that appears on the printed request to show this profile has been selected"
    PR1.PropertyItems("TEXT_STRING").Max = 50
End Sub

Private Sub SetupEdiDetails()

   Dim RS As ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "SELECT * from EDI_Recupients"
   RS.Open "SELECT * FROM EDI_Recipients", ICECon, adOpenKeyset, adLockReadOnly
   
End Sub

Private Sub LoadPicklistDetails(idx)
   
   Dim RS As New ADODB.Recordset
   
   RS.Open "Select * From Request_Picklist Where Picklist_Index=" & Format(idx), ICECon, adOpenKeyset, adLockReadOnly
   Text1.Text = RS!PickList_Name
   If RS!Multichoice Then
      Check1.value = vbChecked
   Else
      Check1.value = vbUnchecked
   End If
   RS.Close
   Command7.Enabled = False
   Command8.Enabled = False
   OrgList.Enabled = True
   Command5.Enabled = True
   Command6.Enabled = True
   TreeView1.Enabled = True
   SSListBar1.Enabled = True
End Sub

Private Sub LoadProfileDetails(idx)
    Dim RS As New ADODB.Recordset
    Dim RS2 As New ADODB.Recordset
    
    RS.Open "Select * from Request_Profiles Where Profile_Index=" & idx, ICECon, adOpenKeyset, adLockReadOnly
    PR1.PropertyItems("Enabled").value = RS!Enabled
    PR1.PropertyItems("WARD").value = RS!Profile_Location_Code & ""
    PR1.PropertyItems("POSITION").value = RS!Profile_Position
    PR1.PropertyItems("CAPTION").value = RS!Profile_Caption & ""
    If Format(RS!Profile_Colour) & "" <> "" Then
        RS2.Open "Select Colour_Code,Colour_Name From Colours Where Colour_Index=" & RS!Profile_Colour, ICECon, adOpenKeyset, adLockReadOnly
        PR1.PropertyItems("PROFILE_COLOUR").value = Val(RS2!Colour_Code)
        PR1.PropertyItems("PROFILE_COLOUR").Tag = Format(RS!Profile_Colour) + "-" + RS2!Colour_Name
        RS2.Close
    End If
    PR1.PropertyItems("HELP").value = RS!Profile_Help
    If Format(RS!Profile_Help_Backcolour) & "" <> "" Then
        RS2.Open "Select Colour_Code,Colour_Name From Colours Where Colour_Index=" & RS!Profile_Help_Backcolour, ICECon, adOpenKeyset, adLockReadOnly
        PR1.PropertyItems("HELP_COLOUR").value = Val(RS2!Colour_Code)
        PR1.PropertyItems("HELP_COLOUR").Tag = Format(RS!Profile_Help_Backcolour) + "-" + RS2!Colour_Name
        RS2.Close
    End If
    PR1.PropertyItems("TEXT_STRING").value = RS!Profile_TextString & ""
    RS.Close
End Sub

Private Sub WriteLog(LogString)
    'Open App.Path + "\ICEConfig.LOG" For Append As #1
    'Print #1, Format(Now, "YYYY/MM/DD HH:MM:SS") + " - " + LogString
    'Close #1
End Sub

Private Sub GetVacantScreenPositions()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    TD1.PropertyItems("SCREEN_POSN").ListItems.Clear
    RS.Open "Select Screen_Position From Request_Tests Where Screen_Panel=" & TD1.PropertyItems("SCREEN_PANEL").value & " And Screen_Panel_Page='" & TD1.PropertyItems("PANEL_PAGE").value & "' Order By Screen_Position", ICECon, adOpenKeyset, adLockReadOnly
    ProgressBar1.value = 0
    ProgressBar1.Max = 45
    TD1.PropertyItems("SCREEN_POSN").ListItems.Add TD1.PropertyItems("SCREEN_POSN").value, TD1.PropertyItems("SCREEN_POSN").value
    For i = 1 To 45
        If Not RS.EOF Then
            If i < RS!Screen_Position Then
                If Format(i) <> TD1.PropertyItems("SCREEN_POSN").value Then TD1.PropertyItems("SCREEN_POSN").ListItems.Add Format(i), Format(i)
            ElseIf i = RS!Screen_Position Then
                RS.MoveNext
            End If
        Else
            If Format(i) <> TD1.PropertyItems("SCREEN_POSN").value Then TD1.PropertyItems("SCREEN_POSN").ListItems.Add Format(i), Format(i)
        End If
    Next i
    RS.Close
    ProgressBar1.value = 0
End Sub


