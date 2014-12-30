VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDtsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dtsURL As String
Private dtsPass As String
Private dtsUser As String
Private fromDTS As String
Private fs As New FileSystemObject
Private xmlSavePath As String
Private blnSaveToFile As Boolean
Private totReqWaiting As Integer
Private CurrentQuery As String
Private WithEvents evSignal As clsDTSSignal
Attribute evSignal.VB_VarHelpID = -1

Private colRequest As New Collection

Private Sub Class_Initialize()
'   dtsURL = "http://nww.reg.nhs.uk/dts/message_tracking_api.asp"
   Set evSignal = New clsDTSSignal
End Sub

Public Property Let DTSName(strNewValue As String)
   fromDTS = strNewValue
End Property

Public Property Get DTSName() As String
   DTSName = fromDTS
End Property

Public Property Let DTSServer(strNewValue As String)
   dtsURL = strNewValue
End Property

Public Function FindSiteDTS(Filename As String) As Boolean
   On Error GoTo procEH
   Dim xmlDoc As New MSXML2.DOMDocument
   
   If xmlDoc.Load(Filename) Then
      fromDTS = xmlDoc.selectSingleNode("DTSConfig/Client/ClientIdentity").firstChild.Text
      FindSiteDTS = True
   Else
      MsgBox "Unable to read DTSClient.cfg - DTS Information will be unavailable", vbExclamation, Filename
      FindSiteDTS = False
   End If
   
   Exit Function
   
procEH:
   MsgBox "Unable to parse the DTS client configuration" & _
          "- it's format may have changed or it may be corrupt" & vbCrLf & _
          "DTS functionality has been disabled", vbExclamation, Filename
   FindSiteDTS = False
End Function

Public Property Let Password(strNewValue As String)
   dtsPass = strNewValue
End Property

Public Property Get ReadDTSResponse(LocalId As String)
   On Error GoTo procEH
   Dim rData As clsDTSRequest
   
'  This is used to check whether or not to update the form with the returned data
   CurrentQuery = LocalId
   
   Set rData = colRequest(LocalId)
   
   If rData.RequestStatus = "Prepared" Then
      If totReqWaiting < 5 Then
         rData.SendDTSRequest
         totReqWaiting = totReqWaiting + 1
      End If
   End If
   
   Set ReadDTSResponse = rData
   Exit Sub
   
procEH:
   If Err.Number = 5 Then
      Set rData = New clsDTSRequest
      rData.SiteDTS = fromDTS
      rData.DTSServer = dtsURL
      rData.SignalClass = evSignal
      rData.PrepareEnquiry = LocalId
      colRequest.Add rData, LocalId
      Resume
   Else
      eClass.CurrentProcedure = "clsDtsAPI.ReadDTSResponse"
      eClass.Add Err.Number, Err.Description, Err.Source
   End If
End Property

Public Property Let SavePath(strNewValue As String)
   If fs.FolderExists(strNewValue) Then
      xmlSavePath = strNewValue
   Else
      xmlSavePath = App.Path
   End If
End Property

Public Property Let SaveOutput(blnNewValue As Boolean)
   blnSaveToFile = blnNewValue
   If xmlSavePath = "" Then
      xmlSavePath = App.Path
   End If
End Property

Public Property Let Username(strNewValue)
   dtsUser = strNewValue
End Property

Private Sub Class_Terminate()
   Dim i As Integer
   Dim rData As clsDTSRequest
   
   For i = 1 To colRequest.Count
      Set rData = colRequest(i)
      rData.Abort
      Set rData = Nothing
   Next i
End Sub

Private Sub evSignal_ResponseReceived(LocalId As String)
   Dim rData As clsDTSRequest
   
   Set rData = colRequest(LocalId)
   If rData.LocalId = CurrentQuery Then
      With rData
         frmNewLogView.wb.DTSData .MsgId, _
                                  .LocalId, _
                                  .Recipient, _
                                  .FromSMTPAddress, _
                                  .ToSMTPAddress, _
                                  .TransferredOn, _
                                  .CurrentStatus, _
                                  .SentOn
      End With
   End If
   Set rData = Nothing
   totReqWaiting = totReqWaiting - 1
End Sub
