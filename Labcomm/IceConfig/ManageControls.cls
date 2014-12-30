VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ManageControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
   Option Explicit

Private strSQL As String
Private uniqueKey As Long
Private blnUpdateTreeView
Private strOrigin

Public Function NewNodeKey(NodeId As String, _
                           DataKey As String, _
                           Optional callMethod As String = "", _
                           Optional Extra As String = "", _
                           Optional MenuStatus As ENUM_MenuStatus = ms_DISABLED, _
                           Optional newNode As Boolean = False)
   Dim nKey As String
   Dim strLevel As String
   Dim mStat As ENUM_MenuStatus
   Dim i As Integer
   Dim strNew As String
   
   If newNode Then
      strNew = "Y"
   Else
      strNew = "N"
   End If
   
      
   nKey = "!*" & strNew & "*!_" & _
          "!#" & NodeId & "#!_" & _
          "!@" & objTV.NodeOrigin & "@!_" & _
          "!$" & callMethod & "$!_" & _
          "!+" & Extra & "+!_" & _
          uniqueKey & "_" & _
          "!<" & DataKey & ">!" & _
          MenuStatus
   
   uniqueKey = uniqueKey + 1
   NewNodeKey = nKey
End Function

Public Sub SetDbStructures(ByRef propList As PropertiesList)
   On Error GoTo procEH
   Dim fLen As String
   Dim tableId As String
   Dim RS As New ADODB.Recordset
   Dim i As Integer
   Dim j As Integer
   
   With propList
      For i = 1 To .Pages.Count
         tableId = .Pages(i).Caption
         strSQL = "SELECT syscolumns.name as Name,syscolumns.length as Length,systypes.name as Type " & _
                  "FROM dbo.syscolumns, dbo.systypes " & _
                  "WHERE id = (select id from dbo.sysobjects where sysobjects.name='" & tableId & "') " & _
                     "AND dbo.syscolumns.xtype=dbo.systypes.xtype"
         
         RS.Open strSQL, iceCon, adOpenForwardOnly, adLockReadOnly
         If RS.RecordCount > 0 Then
'         For j = 0 To RS.RecordCount - 1
'            Debug.Print RS!Name
'         Next j
            For j = 1 To .PropertyItems.Count
               With .PropertyItems(j)
                  If .defaultValue = "" Or IsNull(.defaultValue) Then
                     If .Style = plpsBoolean Then
                        .defaultValue = False
                     ElseIf .Style = plpsNumber Then
                        .defaultValue = 0
                     Else
                        .defaultValue = ""
                     End If
                  End If
               
                  RS.Find "Name = '" & .Tag & "'"
                  If RS.EOF = False Then
                     If .PageKeys = "" Then
                        .PageKeys = propList.Pages(i).Key
                     End If
                     
                     If .PageKeys = "PMIP" Then
                        .PageKeys = "EDI" 'propList.Pages(i).Key
                     End If
                     
                     If RS!Type = "varchar" Then
                        If .max = 0 Then
                           .max = RS!Length
                        End If
                     End If
                  End If
                  RS.MoveFirst
               End With
            Next j
         End If
         RS.Close
      Next i
   End With
   Set RS = Nothing
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "IceConfig.ManageControls.SetDbStructures"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Let TreeViewUpdate(blnNewValue As Boolean)
   blnUpdateTreeView = blnNewValue
End Property

Public Property Get TreeViewUpdate() As Boolean
   TreeViewUpdate = blnUpdateTreeView
End Property
