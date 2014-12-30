VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "frameDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum fNames_Enum
   Fra_FILEDETAILS = 0
   Fra_BEHAVIOUR = 1
   Fra_CONFIG = 2
   Fra_EDI = 3
   Fra_INCEX = 4
   Fra_LOGVIEW = 5
   Fra_LOGSEARCH = 6
   Fra_PICKLIST = 7
   Fra_PROFILE = 8
   Fra_TESTDETAILS = 9
   Fra_HELP = 10
   Fra_INFO = 11
   Fra_ReadCode = 12
   Fra_COLOURS = 13
   fra_None = 14
End Enum

Private curFrame As fNames_Enum
Private workFrame As fNames_Enum
Private intPData(13, 4) As Integer
Private fColl As New Collection

Private tviewProc As String
Private tviewParams As String
Private ShowFrame As Long
Private DisplayRCErrors As Boolean
Private defRCStat As Long
Private logFile As String

Private Sub Class_Initialize()
   Dim i As Integer
   
   intPData(0, 0) = 7925   '  File Details Height
   intPData(0, 1) = 8200   '  File Details Width
   intPData(1, 0) = 4335   '  Behaviour
   intPData(1, 1) = 7515
   intPData(2, 0) = 4455   '  Config
   intPData(2, 1) = 8115
   intPData(3, 0) = 5850   '  EDI
   intPData(3, 1) = 7485
   intPData(4, 0) = 2585   '  IncEx
   intPData(4, 1) = 7455
   intPData(5, 0) = 8640   '  LogView
   intPData(5, 1) = 8900
   intPData(6, 0) = 3525   '  Log Search
   intPData(6, 1) = 5000
   intPData(7, 0) = 4605   '  PickList
   intPData(7, 1) = 8055
   intPData(8, 0) = 7005   '  Profile
   intPData(8, 1) = 8025
   intPData(9, 0) = 1740   '  Test Details/Monitors
   intPData(9, 1) = 8130
   intPData(10, 0) = 3735  '  Help Screen
   intPData(10, 1) = 8095
   intPData(11, 0) = 1725  '  Info screen
   intPData(11, 1) = 3765
   intPData(12, 0) = 2525  '  Read Code Errors
   intPData(12, 1) = 7000
   intPData(13, 0) = 1100  '  Colours
   intPData(13, 1) = 3800
      
   For i = 0 To UBound(intPData)
      frmMain.frapanel(i).Visible = False
   Next i
   ShowFrame = 3
   DisplayRCErrors = False
End Sub

Public Property Let OutputFileName(strNewValue As String)
   logFile = strNewValue
End Property

Public Property Let RefreshProc(strNewValue As String)
   tviewProc = strNewValue
   tviewParams = ""
End Property

Public Property Get RefreshProc() As String
   RefreshProc = tviewProc
End Property

Public Property Let RefreshProcParams(strNewValue As String)
   tviewParams = strNewValue
End Property

Public Property Get RefreshProcParams() As String
   RefreshProcParams = tviewParams
End Property

Public Property Let WorkingFrame(fNewValue As fNames_Enum)
   workFrame = fNewValue
End Property

Public Property Get WorkingFrame() As fNames_Enum
   WorkingFrame = workFrame
End Property

Public Property Get CurrentFrame() As String
   CurrentFrame = curFrame '  frmmain.fraPanel(curFrame).Name
End Property

Public Property Let LogViewStatus(lngNewValue As Long)
   defRCStat = lngNewValue
End Property

Friend Sub ShowReadCodes()
'   Dim RS As New ADODB.Recordset
'   Dim RS2 As New ADODB.Recordset
'   Dim strSQL As String
'   Dim i As Integer
'   Dim strArray() As String
'   Dim maxColWidth As Long
'   Dim thisWidth As Long
'   Dim iceCmd As New ADODB.Command
'
'   With frmMain.vbgReadCodes
'      .Clear True
'      .AddColumn "1", "Local Code", ecgHdrTextALignLeft, , 80
'      .AddColumn "2", "Read", ecgHdrTextALignLeft, , 40
'      .AddColumn "3", "Comment", ecgHdrTextALignLeft, , 280
'      .RowMode = True
'   End With
'
'   With iceCmd
'      .ActiveConnection = iceCon
'      .CommandType = adCmdStoredProc
'      .CommandText = "IceConfig_ReadCodeErrors"
'      .Parameters.Append .CreateParameter("Stat", adInteger, adParamInput, , defRCStat)
'      Set RS = .Execute
'   End With
'
'   With frmMain.vbgReadCodes
'      i = 1
'      Do Until RS.EOF
'         .AddRow
'         .Cell(i, 1).Text = RS(0)
'         .Cell(i, 2).Text = RS(1)
'         .Cell(i, 3).Text = RS(2)
'         RS.MoveNext
'         i = i + 1
'      Loop
'      .AutoWidthColumn ("3")
'   End With
'
'   With frmMain.frapanel(12)
'      .Height = intPData(12, 0)
'      .Width = intPData(12, 1)
'      .Top = 7000
'      .Left = 900
'      .Visible = (RS.RecordCount > 0)
'      DisplayRCErrors = .Visible
'   End With
'   RS.Close
'   Set RS = Nothing
End Sub

Public Sub Show(Optional thisframe As fNames_Enum = -1, _
                Optional dData As String = "")
   Dim frameId As Long
   
   If thisframe = -1 Then
      frameId = ShowFrame
   Else
      frameId = thisframe
   End If
   
   If frameId = fra_None Then
      Hide
'      With frmMain.wb
'         .Top = 400
'         .Left = 200
'         .Height = 6000
'         .Width = 8000
'         If thisframe > -1 Then
'            .NavigateTo ""
'         End If
'         .Visible = True
'      End With
   Else
      SetUpPanel frameId, dData
'      If frameId <> curFrame Then   'Or dData <> "" Then
      If frmMain.frapanel(frameId).Visible = False Then
         'frmMain.wb.Visible = False
         With frmMain.frapanel(frameId)
            Select Case frameId
               Case Fra_INCEX
                  .Top = 7000
                  .Left = 600
               
               Case Fra_COLOURS
                  .Top = 6800
                  .Left = 2500
               
               Case Fra_TESTDETAILS
                  Hide
   '               frmMain.wb.Visible = Tr.ue
                  .Top = 150
                  .Left = 195
                  With frmMain.wb
                     .Top = 2600
                     .Left = 200
                     .Height = 6000
                     .Width = 8100
'                     .NavigateTo ""
                     .Visible = True
                  End With
                  
               Case Else
                  Hide
                  .Top = 500
                  .Left = 350
            End Select
         
'         If frameId = Fra_INCEX Then
'            .Top = 7000
'            .Left = 600
'
'         ElseIf frameId = Fra_COLOURS Then
'            .Top = 6800
'            .Left = 2500
'
'         Else
'            Hide
'            .Top = 500
'            .Left = 350
'         End If
            .Height = intPData(frameId, 0)
            .Width = intPData(frameId, 1)
            .Refresh
            .Visible = True
         End With
         frmMain.frapanel(12).Visible = DisplayRCErrors
      End If
      curFrame = frameId
      frmMain.Refresh
   End If
'   ShowFrame = 3
End Sub

Public Sub Hide()
   Dim i As Integer
   
   For i = 0 To 13
      frmMain.frapanel(i).Visible = False
   Next i
'   frmMain.mainPanel.Caption = "ICE Configuration...  Please select an option from the bar on the left"
   frmMain.Refresh
   
End Sub

Public Sub SetUpPanel(whatFrame As fNames_Enum, Extra As String)

   Dim i As Integer
   
   With frmMain
      Select Case whatFrame
         Case Fra_FILEDETAILS
            .mainPanel.Caption = Extra
            
         Case Fra_BEHAVIOUR
            
         Case Fra_CONFIG
            .mainPanel.Caption = "Configuration Editor"
            
         Case Fra_EDI
            .mainPanel.Caption = Extra
            .ediPr.Visible = True
            
         Case Fra_INCEX
            .mainPanel.Caption = "Include/Exclude/Reflex Tests"
            .txtTestDisplay.Text = frmMain.txtTestDesc(Val(Extra)).Text
            
         Case Fra_LOGVIEW
            .mainPanel.Caption = "Log Information"
            .frapanel(whatFrame).Caption = logFile
'            .LogText.BackColor = LogBackColour
   
         Case Fra_LOGSEARCH
            .mainPanel.Caption = "Search Log Information"
            For i = 1 To 4
               .fraSCrit(i).Visible = False
            Next
            If Extra = 0 Then
               .cmdSrchOk.Top = 2960
               .cmdSrchCancel.Top = 2960
               .frapanel(6).Height = 3525
            Else
               .TreeView1.Visible = False
               .Refresh
               .frapanel(6).Height = 4705
               .cmdSrchOk.Top = 4160
               .cmdSrchCancel.Top = 4160
               With .fraSCrit(Val(Extra))
                  .Top = 2920
                  .Left = 550
                  .Height = 975
                  .Width = 4065
                  .Visible = True
               End With
               If Extra = 1 Then
                  If .ComboSrchPractice.ListCount > 0 Then
                     .ComboSrchPractice.ListIndex = 0
                  End If
               End If
               .optLogSearch(Val(Extra)).value = True
            End If
            
            If Val(Extra) < 2 Then
               frmMain.fraShowErr.Visible = True
            Else
               frmMain.optShowErr(0).value = True
               frmMain.fraShowErr.Visible = False
            End If
            
         Case Fra_PICKLIST
            frmMain.mainPanel.Caption = "Picklist Editor"
            
         Case Fra_PROFILE
            frmMain.mainPanel.Caption = "Profile Editor"
            
         Case Fra_TESTDETAILS
            frmMain.mainPanel.Caption = ""
            'frmMain.optStatPeriod(1).value = True
            
         Case Fra_HELP
            frmMain.mainPanel.Caption = "Instructions"
            frmMain.txtHelp(0).Visible = False
            frmMain.txtHelp(1).Visible = False
            frmMain.txtHelp(Val(Extra)).Visible = True
            frmMain.txtHelp(1).Visible = True
            frmMain.txtHelp(Val(Extra)).Left = 100
            frmMain.txtHelp(Val(Extra)).Top = 200
            frmMain.txtHelp(Val(Extra)).Width = intPData(10, 1) - 200
            frmMain.txtHelp(Val(Extra)).Height = intPData(10, 0) - 400
            workFrame = Fra_HELP
   
         Case Fra_INFO
            frmMain.mainPanel.Caption = "Working... Please wait"
            frmMain.frapanel(whatFrame).Top = 3000
            frmMain.frapanel(whatFrame).Left = 3500
            frmMain.lblComment.Caption = Extra
            
      End Select
   End With
End Sub

Public Sub RefreshDisplay(Optional ClassId As String = "")
   On Error GoTo procEH
   Dim tClass As Object
   
   If tviewProc = "" Then
      frmMain.frapanel(12).Visible = False
   Else
      frmMain.MousePointer = vbHourglass
      frmMain.SSPanel2.Caption = "Searching..."
      
      If tviewParams = "" Then
         CallByName loadCtrl, tviewProc, VbMethod
      Else
         CallByName loadCtrl, tviewProc, VbMethod, tviewParams
      End If
'      Show Fra_HELP, "1"
      
      frmMain.TreeView1.Visible = True
      frmMain.MousePointer = vbNormal
      Set tClass = Nothing
   End If
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.FrameDataClass.RefreshDisplay"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let ShowRCFrame(blnNewValue As Boolean)
   DisplayRCErrors = blnNewValue
End Property

Public Property Get FrameToShow() As fNames_Enum
   FrameToShow = ShowFrame
End Property

Public Property Let FrameToShow(lngNewValue As fNames_Enum)
   ShowFrame = lngNewValue
End Property
