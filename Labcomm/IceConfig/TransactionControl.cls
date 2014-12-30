VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransactionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TransActive As Boolean
Private dbCon As ADODB.Connection
Private TimeoutAfter As Integer
Private TimeRemaining As Integer
Private ControlledForm As Form
Private WarningDisplayed As Boolean

Private Sub Class_Initialize()
   TimeoutAfter = 20
End Sub

Public Property Let DBConnection(ByRef objNewValue As ADODB.Connection)
   Set dbCon = iceCon
End Property

Public Property Let SetForm(objNewValue As Form)
   Set ControlledForm = objNewValue
End Property

Public Property Let TimeoutValue(intNewValue As Integer)
   TimeoutAfter = intNewValue
End Property

Public Sub StartTransaction(Optional FormToControl As Form = Nothing, _
                            Optional TimeoutIn As Integer = 0)
   If TransActive = False Then
      iceCon.BeginTrans
      TransActive = True
      If TimeoutIn > 0 Then
         If FormToControl Is Nothing Then
            MsgBox "Specify form to monitor if timeout is required", vbInformation, "Transaction Timeout"
         Else
            TimeoutAfter = TimeoutIn
            frmTimerClose.Timer1.Interval = 1000
            frmTimerClose.Timer1.Enabled = True
            Set ControlledForm = FormToControl
            frmTimerClose.TimeoutValue = TimeoutIn
            Load frmTimerClose
         End If
      End If
   End If
End Sub

Public Sub EndTransaction()
   If TransActive Then
      iceCon.CommitTrans
      TransActive = False
   End If
End Sub

Public Sub AbandonTransaction()
   If TransActive Then
      iceCon.RollbackTrans
      TransActive = False
   End If
End Sub

Public Sub StartTimer()
   frmTimerClose.Timer1.Interval = 1000
   frmTimerClose.Timer1.Enabled = True
   Load frmTimerClose
End Sub

Public Sub ResetTimer()
   frmTimerClose.Visible = False
   frmTimerClose.TimeoutValue = TimeoutAfter
   ControlledForm.Visible = True
End Sub

Public Sub TimeoutWarning()
   ControlledForm.Hide
   frmTimerClose.Visible = True
End Sub

Public Sub TimeOutExpired()
   AbandonTransaction
   If ControlledForm.Name <> frmMain.Name Then
      Unload ControlledForm
   End If
   Unload frmTimerClose
   MsgBox "To release the locks on the database, the transaction has been abandoned", vbInformation, "Timeout"
   frmMain.Visible = True
End Sub

Public Sub TimeoutAbandoned()
   AbandonTransaction
   frmTimerClose.Timer1.Enabled = False
   Unload frmTimerClose
End Sub
