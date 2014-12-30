VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private AssociatedForm As Form
Private dLabel As Label
Private ShowWarning As Boolean
Private TimeOut As Integer
Private TimeLeft As Integer

Private Sub Class_Initialize()
   Set dLabel = Nothing
End Sub

Public Property Let SetLabel(lblNewValue As Label)
   Set dLabel = lblNewValue
End Property

Public Property Let SetForm(objNewValue As Form)
   Set AssociatedForm = objNewValue
End Property

Public Property Let TimeoutValue(intNewValue As Integer)
   TimeOut = intNewValue
   TimeLeft = TimeOut
End Property

Public Property Get Warning() As Boolean
   Warning = ShowWarning
End Property

Private Sub cmdContinue_Click()
   Unload AssociatedForm
   If transCount > 0 Then
      ICECon.RollbackTrans
      transCount = 0
   End If
End Sub

Private Sub cmdReset_Click()
   TimeLeft = TimeOut
   blnWarn = False
   Me.Visible = False
End Sub

Public Sub CloseForm()
   ICECon.RollbackTrans
   transCount = 0
   Unload AssociatedForm
End Sub

Public Sub Timer()
   If TimeLeft > 0 Then
      If TimeLeft < 15 Then
         If ShowWarning = False Then
               AssociatedForm.Hide
               frmTimerClose.Visible = True
               frmTimerClose.Show
               ShowWarning = True
         End If
      End If
   Else
      CloseForm
   End If
   
   TimeLeft = TimeLeft - 1
   If Not dLabel Is Nothing Then
      dLabel.Caption = TimeLeft
   End If
End Sub
