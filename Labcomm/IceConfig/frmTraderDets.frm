VERSION 5.00
Begin VB.Form frmTraderDets 
   Caption         =   "EDI Trader Details"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraIC 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   600
      TabIndex        =   7
      Top             =   2880
      Width           =   2535
      Begin VB.TextBox txtIC 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Interchange"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame fraUnlinked 
      Caption         =   "Message Types"
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
      Begin VB.ComboBox cboMsgIC 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox txtFreePart 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtTraderCode 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox chkLink 
      Caption         =   "Single Interchange number"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Label lblFP 
      Caption         =   "Free Part"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblTC 
      Caption         =   "Trader Code"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmTraderDets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public recipientStatus As Integer
'  -1 = New & Saved - awaiting update; 0 = New - not saved; anything else = existing recipient
'  Flag is set in Proprty Let for EDI_Ref_Index and changed in "cmdApply_Click" when a new record
'  has been created. It is checked in the form Unload event procedure

Private natCode As String
Private TraderCode As String
Private FreePart As String
Private blnICWarning As Boolean
Private blnLinkIC As Boolean
Private refId As Long
Private sharedPractices As Integer

Public Property Let EDI_NatCode(strNewValue As String)
   natCode = strNewValue
End Property

Public Property Get EDI_RefIndex() As Long
   EDI_RefIndex = refId
End Property

Public Property Let EDI_RefIndex(lngNewValue As Long)
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   refId = lngNewValue
   recipientStatus = refId
   
   If refId = 0 Then
      chkLink.value = 1
      chkLink_Click
      blnLinkIC = True
      chkLink.Enabled = False
      txtIC.Text = 0
   Else
      chkLink.Enabled = True
      strSQL = "SELECT Count(*) FROM EDI_Recipients WHERE Ref_Index = " & refId
      RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
      sharedPractices = RS(0)
      
      If sharedPractices = 0 Then ' No EDI_Recipient record - user must have come back in after initial set-up. Reset status
         recipientStatus = -1
      End If
      
      If RS(0) <= 1 Then
         lblTC.ForeColor = &H80000012
         lblFP.ForeColor = &H80000012
      Else
         lblTC.ForeColor = vbRed
         lblFP.ForeColor = vbRed
      End If
      
      RS.Close
      
      strSQL = "SELECT * FROM EDI_Recipient_Ref WHERE Ref_Index = " & refId
      RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
      
      TraderCode = RS!EDI_Trader_Account
      FreePart = RS!EDI_Free_Part
      blnLinkIC = RS!Link_Interchange_Nos
      
      RS.Close
      
      strSQL = "SELECT * FROM EDI_Interchange_No WHERE Ref_Index = " & refId
      RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
      
      If blnLinkIC = False Then
         cboMsgIC.Clear
         i = 0
         
         Do Until RS.EOF
            cboMsgIC.AddItem RS!EDI_Msg_Format
            cboMsgIC.ItemData(i) = RS!EDI_Last_Interchange
            i = i + 1
            RS.MoveNext
         Loop
         
         cboMsgIC.ListIndex = 0
         
      Else
         txtIC.Text = RS!EDI_Last_Interchange
         txtIC.Tag = txtIC.Text
         chkLink_Click
      End If
      
      'Status blnLinkIC
      
      txtTraderCode.Text = TraderCode
      txtFreePart.Text = FreePart
      chkLink.value = Abs(CInt(blnLinkIC))
      
      RS.Close
   End If
   
   Set RS = Nothing
End Property

Private Sub status(Linked As Boolean)
   If Linked Then
      fraUnlinked.Visible = False
      fraIC.Top = 1680
      Me.Height = 3570
   Else
      fraIC.Top = 2940
      fraUnlinked.Visible = True
      Me.Height = 4755
   End If
End Sub

Private Sub cboMsgIC_Click()
   txtIC.Text = cboMsgIC.ItemData(cboMsgIC.ListIndex)
   blnICWarning = False
End Sub

Private Sub cboMsgIC_DropDown()
   cboMsgIC.ItemData(cboMsgIC.ListIndex) = txtIC.Text
End Sub

Private Sub chkLink_Click()
   Dim strSQL As String
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   
   If chkLink.value = 0 Then
      '  We are turning the link off
      If cboMsgIC.ListCount = 0 Then
         strSQL = "SELECT EDI_Msg_Format " & _
                  "FROM EDI_Msg_Types " & _
                  "WHERE EDI_Org_NatCode = '" & natCode & "'"
                     '"AND EDI_Msg_Active = 1"
         RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
         Do Until RS.EOF
            cboMsgIC.AddItem RS!EDI_Msg_Format
            cboMsgIC.ItemData(i) = txtIC.Text
            i = i + 1
            RS.MoveNext
         Loop
         
         If RS.RecordCount > 0 Then
            cboMsgIC.ListIndex = 0
         End If
      End If
      
      status False
   Else
      '  We are turning the link on
      'If MsgBox("Linking the interchange nos. will maintain a single sequence regardless of the message type" & _
                vbCrLf & "Are you sure you wish to abandon seperate interchange sequences", _
                vbExclamation Or vbSystemModal Or vbYesNo, "Confirm Interchange sequence change") = vbYes Then
         
         status True
      'End If
   End If
End Sub

Private Sub chkLink_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Integer
   
   If chkLink.value = 1 Then
      If cboMsgIC.ListCount > 0 Then
         If MsgBox("This will maintain a single sequence number regardless of the message type." & _
                   "The interchange will be set to the highest value of those present. " & vbCrLf & _
                   "Are you sure you wish to abandon seperate interchange sequences", _
                   vbExclamation Or vbSystemModal Or vbYesNo, "Confirm Interchange sequence change") = vbNo Then
            chkLink.value = 0
         Else
            For i = 0 To cboMsgIC.ListCount - 1
               If Val(cboMsgIC.ItemData(i)) > Val(txtIC.Text) Then
                  txtIC.Text = cboMsgIC.ItemData(i)
               End If
            Next i
         End If
      Else
         MsgBox "No Message Types are associated with this practice. Recording Interchange by message type " & _
                "will be impossible.", vbCritical Or vbSystemModal, "Unlink failed"
         chkLink.value = 1
      End If
   End If
   
   'blnLinkIC = (chkLink.value = 1)
End Sub

Private Sub cmdApply_Click()
   On Error GoTo procEH
   Dim iceCmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim refIndex As Long

   '  Does this trader code/Free part already exist? (Phoenix)
   strSQL = "SELECT Ref_Index " & _
            "FROM EDI_Recipient_Ref " & _
            "WHERE EDI_Trader_Account = '" & txtTraderCode.Text & "' " & _
               "AND EDI_Free_Part = '" & txtFreePart.Text & "'"
   RS.Open strSQL, iceCon, adOpenStatic, adLockReadOnly
   
   iceCon.BeginTrans
      
   If RS.EOF Then
   '  Set up a new EDI_Recipient_Ref record
      With iceCmd
         .ActiveConnection = iceCon
         .CommandType = adCmdStoredProc

         .CommandText = "ICECONFIG_New_RecipientRef"

         .Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
         .Parameters.Append .CreateParameter("R_TradAcc", adVarChar, adParamInput, 10, txtTraderCode.Text)
         .Parameters.Append .CreateParameter("R_FreePart", adVarChar, adParamInput, 5, txtFreePart.Text)
         .Parameters.Append .CreateParameter("R_Link", adBoolean, adParamInput, , (chkLink.value = 1))
         .Execute

         refIndex = .Parameters("Return").value

         .Parameters.Delete "R_TradAcc"
         .Parameters.Delete "R_FreePart"
         .Parameters.Delete ("R_Link")

         .CommandText = "ICECONFIG_New_InterchangeRecord"
         .Parameters.Append .CreateParameter("RefId", adInteger, adParamInput, , refIndex)
         .Parameters.Append .CreateParameter("ICVal", adInteger, adParamInput)
         .Parameters.Append .CreateParameter("MsgFormat", adVarChar, adParamInput, 16)

         If chkLink.value = 0 Then

            For i = 0 To cboMsgIC.ListCount - 1
               .Parameters("MsgFormat").value = cboMsgIC.List(i)
               .Parameters("ICVal").value = cboMsgIC.ItemData(i)
               .Execute
            Next i

         Else

            .Parameters("MsgFormat").value = "None"
            .Parameters("ICVal").value = txtIC.Text
            .Execute
         End If

         '  TraderCode variable is set when setting the EDI_Ref_Index property. If it is a null string, no
         '  recipient record was found. If this has been amended, we could have orphaned records.

         If TraderCode <> "" Then
            RS.Close
            strSQL = "SELECT EDI_NatCode " & _
                     "FROM EDI_Recipients " & _
                     "WHERE Ref_Index = " & refId
            RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly

            If RS.EOF Then
               '  No EDI_Recipients point to this record, so delete it
               strSQL = "DELETE FROM EDI_Recipient_Ref " & _
                        "WHERE Ref_Index = " & refId

               iceCon.Execute strSQL

               strSQL = "DELETE FROM EDI_Interchange_No WHERE Ref_Index = " & refId

               iceCon.Execute strSQL
            End If
         End If

'         If recipientStatus = 0 Then
'            '  EDI_Recipient Record does not yet exist. Flag to indicate delayed uodate
'            recipientStatus = -1
'         Else
            '  This recipient already exists - so point to new record.
            strSQL = "UPDATE EDI_Recipients SET Ref_Index = " & refIndex & _
                     " WHERE EDI_NatCode = '" & natCode & "'"

            iceCon.Execute strSQL

            'recipientStatus = 0

         'End If
      End With

   Else
      '***************************************************
      '  Recipient Ref record already exists
      '***************************************************
   
      refIndex = RS!Ref_Index
      
      If recipientStatus > 0 Then
         '  A status of 0 indicates this is a new recipient pointing to an existing record (e.g. TPP practice)
            
         If txtTraderCode.Text = TraderCode And txtFreePart.Text = FreePart Then
            '**************************************************************************
            '  We haven't changed the Trader ref, so amend the link param as necessary
            '  If the trader code has changed - we do NOT want to make any amandments
            '  to the existing link and interchange values!
            '**************************************************************************
            Set iceCmd = Nothing
            Set iceCmd = New ADODB.Command
            
            With iceCmd
               .ActiveConnection = iceCon
               .CommandType = adCmdStoredProc
                        
               .CommandText = "ICECONFIG_Update_RecipientRef"
               
               .Parameters.Append .CreateParameter("RefId", adInteger, adParamInput, , refIndex)
               .Parameters.Append .CreateParameter("Link", adBoolean, adParamInput, , (chkLink.value = 1))
               .Execute
            End With
         
            If chkLink.value <> Abs(CInt(blnLinkIC)) Then
               '***************************************************
               '  The Link Interchange flag has been amended...
               '***************************************************
               Set iceCmd = Nothing
               Set iceCmd = New ADODB.Command
               
               With iceCmd
                  .ActiveConnection = iceCon
                  .CommandType = adCmdStoredProc
                  
                  .CommandText = "ICECONFIG_New_InterchangeRecord"
                  
                  .Parameters.Append .CreateParameter("RefId", adInteger, adParamInput, , refIndex)
                  .Parameters.Append .CreateParameter("ICVal", adInteger, adParamInput)
                  .Parameters.Append .CreateParameter("MsgFormat", adVarChar, adParamInput, 16)
                  
                  strSQL = "DELETE FROM EDI_Interchange_No WHERE Ref_Index = " & refIndex
                  iceCon.Execute (strSQL)
               
                  If chkLink.value = 0 Then
                  
                     For i = 0 To cboMsgIC.ListCount - 1
                        .Parameters("MsgFormat").value = cboMsgIC.List(i)
                        .Parameters("ICVal").value = CInt(cboMsgIC.ItemData(i))
                        .Execute
                     Next i
                  
                  Else
                     
                     .Parameters("MsgFormat").value = "None"
                     .Parameters("ICVal").value = txtIC.Text
                     .Execute
                     
                  End If
               End With
            End If
                  
            '***************************************************
            '  Update interchange values
            '***************************************************
                  
            Set iceCmd = Nothing
            Set iceCmd = New ADODB.Command
            
            With iceCmd
               .ActiveConnection = iceCon
               .CommandType = adCmdStoredProc
               
               .Parameters.Append .CreateParameter("RefId", adInteger, adParamInput, , refIndex)
               .Parameters.Append .CreateParameter("ICVal", adInteger, adParamInput)
               
               If chkLink.value = 0 Then
                  .CommandText = "ICEMSG_Update_Interchange_ByMsgType"
               
                  .Parameters.Append .CreateParameter("MsgFormat", adVarChar, adParamInput, 16)
                  
                  For i = 0 To cboMsgIC.ListCount - 1
                     .Parameters("ICVal").value = Val(cboMsgIC.ItemData(i))
                     .Parameters("MsgFormat").value = cboMsgIC.List(i)
                     .Execute
                  Next i
               
               Else
                  .CommandText = "ICEMSG_Update_Interchange"
                  .Parameters("ICVal").value = Val(txtIC.Text)
                  .Execute
               End If
            End With
         Else
         
            '**********************************************************
            '  Trader code has been changed on an existing recipient.
            '**********************************************************
            RS.Close
            
            strSQL = "UPDATE EDI_Recipients SET Ref_Index = " & refIndex & _
                     " WHERE EDI_NatCode = '" & natCode & "'"
         
            iceCon.Execute strSQL
                        
            strSQL = "SELECT EDI_NatCode " & _
                     "FROM EDI_Recipients " & _
                     "WHERE Ref_Index = " & refId
            RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
            
            '  Any recipients still pointing to the old record?
            
            If RS.EOF Then
               '  No, so delete it
               strSQL = "DELETE FROM EDI_Recipient_Ref " & _
                        "WHERE Ref_Index = " & refId
               
               iceCon.Execute strSQL
               
               strSQL = "DELETE FROM EDI_Interchange_No WHERE Ref_Index = " & refId
               
               iceCon.Execute strSQL
            End If
                        
            'recipientStatus = 0
            
         End If
      Else
         '******************************************************************************
         '  A New recipient - we are unable to update the Ref_Index on EDI_Recipient
         '  because the being a new practice, the record does not yet exist!
         '******************************************************************************
         recipientStatus = -1
      End If
   End If
      
   refId = refIndex
   
   frmMain.ediPr("REFID").value = refIndex
   
   iceCon.CommitTrans
   Set iceCmd = Nothing
   'Set iceCmd = New ADODB.Command
   RS.Close
   Set RS = Nothing
   
   Me.Hide
   
   Exit Sub
   
procEH:
   iceCon.RollbackTrans
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "frmTraderDets.cmdApply_Click"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdCancel_Click()
   ' A status of -1 indicates the Recipient_Ref record has been created - threfeore we will need to tidy if cancelled.
   If recipientStatus > 0 Then
      recipientStatus = 0
   End If
   
   Unload Me
End Sub

Public Sub ConfirmUpdate()
   Dim strSQL As String
   
   strSQL = "UPDATE EDI_Recipients SET Ref_Index = " & refId & _
            " WHERE EDI_NatCode = '" & natCode & "'"
   
   iceCon.Execute strSQL
   
   '  Unset the tidy flag. If not, the Recipient_Ref records will be deleted in Form_Unload
   recipientStatus = refId
End Sub

Private Sub Form_Load()
'   chkLink.value = 1
'   chkLink_Click
'   blnLinkIC = True
   blnICWarning = False
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'   Dim strSQL As String
'   '**************************************************************************************************
'   '  The Unload event is utilised to see if the database needs tidying.
'   '
'   '**************************************************************************************************
'   If recipientStatus = -1 Then
'      '  The EDI_Recipients record has not been saved, so remove linked records from the database.
'      strSQL = "DELETE FROM EDI_Recipient_Ref WHERE Ref_Index = " & refId & "; " & vbCrLf & _
'               "DELETE FROM EDI_Interchange_No WHERE Ref_Index = " & refId
'      iceCon.Execute strSQL
''   ElseIf recipientStatus = 0 Then
''      strSQL = "UPDATE EDI_Recipients SET Ref_Index = " & refId & _
''               " WHERE EDI_NatCode = '" & natCode & "'"
'   End If
'
''   If recipientStatus <> 0 Then
''      iceCon.Execute strSQL
''   End If
'
'End Sub

Private Sub txtFreePart_GotFocus()
   txtFreePart.SelStart = 0
   txtFreePart.SelLength = Len(txtFreePart.Text)
End Sub

Private Sub txtFreePart_Validate(Cancel As Boolean)
   If Len(txtFreePart.Text) > 0 Then
      If IsNumeric(txtFreePart.Text) Then
         If Len(txtFreePart.Text) < 5 Then
            txtFreePart.Text = Format(txtFreePart.Text, "0000#")
         End If
         
         If Len(txtFreePart.Text) > 5 Then
            MsgBox "Max 5 digits only", vbExclamation, "Trader Code Input Error"
            Cancel = True
         End If
      Else
         MsgBox "5 digit numeric only", vbExclamation, "Trader Code Error"
         Cancel = True
      End If
   End If
End Sub

Private Sub txtIC_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim mbReply As VbMsgBoxResult
   
   If blnICWarning = False Then
      mbReply = MsgBox("Amending this value may cause messages to be rejected by the practice." & vbCrLf & _
                "Are you sure you wish to amend the interchange sequence?", _
                vbExclamation Or vbSystemModal Or vbYesNo, "Integrity warning")
      
      If mbReply = vbNo Then
         '  Reset to existing value
         txtIC.Text = cboMsgIC.ItemData(cboMsgIC.ListIndex)
         txtIC.SelStart = Len(txtIC.Text)
      Else
         blnICWarning = True
      End If
   End If
   
   If KeyCode <> 8 Then
      If Not IsNumeric(Chr(KeyCode)) Then
         MsgBox "Only numeric values allowed", vbExclamation, "Invalid interchange sequence"
         txtIC.Text = Left(txtIC.Text, Len(txtIC.Text) - 1)
         txtIC.SelStart = Len(txtIC.Text)
      End If
   End If
End Sub

Private Sub txtIC_Validate(Cancel As Boolean)
   If IsNumeric(txtIC.Text) Then
      If refId = 0 Then
         txtIC.Tag = txtIC.Text
      Else
'         If MsgBox("Amending this value may cause messages to be rejected by the practice." & vbCrLf & _
'                   "Are you sure you wish to amend the interchange Id?", _
'                   vbExclamation Or vbSystemModal Or vbYesNo, "Integrity warning") = vbNo Then
'            '  Reset to existing value
'            txtIC.Text = txtIC.Tag
'         Else
            txtIC.Tag = txtIC.Text
            
            If blnLinkIC = False Then
               cboMsgIC.ItemData(cboMsgIC.ListIndex) = txtIC.Text
            End If
'         End If
      End If
'   Else
'      MsgBox "Numeric values only", vbExclamation, "Invalid interchange"
'      Cancel = True
'      txtIC.Text = txtIC.Tag
'      txtIC.SelStart = 0
'      txtIC.SelLength = Len(txtIC.Tag)
'      txtIC.SetFocus
   End If
End Sub

Private Sub txtTraderCode_GotFocus()
   txtTraderCode.SelStart = 0
   txtTraderCode.SelLength = Len(txtTraderCode.Text)
End Sub

Private Sub txtTraderCode_Validate(Cancel As Boolean)
   If Len(txtTraderCode.Text) > 0 Then
      If IsNumeric(txtTraderCode.Text) Then
         If Len(txtTraderCode.Text) < 10 Then
            txtTraderCode.Text = Format(txtTraderCode.Text, "000000000#")
         End If
         
         If Len(txtTraderCode.Text) > 10 Then
            MsgBox "Max 10 digits only", vbExclamation, "Trader Code Input Error"
            Cancel = True
         End If
      Else
         MsgBox "10 digit numeric only", vbExclamation, "Trader Code Error"
         Cancel = True
      End If
   End If
End Sub
