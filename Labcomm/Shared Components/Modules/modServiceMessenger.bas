Attribute VB_Name = "modServiceMessenger"
Public MessengerControl As VBControlExtender
Public Sub CreateMessenger(Owner As VB.Form, Host As String, Port As Integer)
    On Local Error GoTo WSError
    Licenses.Add "MSWinsock.Winsock", "2c49f800-c2dd-11cf-9ad6-0080c7e7b78d"
    Set MessengerControl = Owner.Controls.Add("MSWinsock.Winsock", "ServiceMessengerControl")
    MessengerControl.object.RemoteHost = Host
    MessengerControl.object.RemotePort = Port
    MessengerControl.object.protocol = 1
    Randomize
    LP = Int(999 * Rnd) + 9000
    MessengerControl.object.Bind LP
    On Local Error GoTo 0
    Exit Sub
WSError:
    Select Case Err.Number
        Case 10048
            LP = Int(999 * Rnd) + 9000
            Resume
        Case 733
            Resume Next
    End Select
End Sub

Public Sub SendMessage(AppName As String, Message As String)
    On Local Error GoTo WSError
    MessengerControl.object.SendData AppName + " - " + Message
    Exit Sub
WSError:
    On Local Error Resume Next
    Select Case Err.Number
        Case 40006
            MessengerControl.object.SendData AppName + " - " + Message
    End Select
End Sub

Public Sub DestroyMessenger(Owner As VB.Form)
    Owner.Controls.Remove "ServiceMessengerControl"
End Sub

