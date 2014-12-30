VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIceMsgLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EnmuLogTypes
   LT_Process = 0
   LT_Practice = 1
   LT_Log = 2
   LT_Error = 3
End Enum

Private sessionRoot As MSXML2.IXMLDOMNode
Private logRoot As MSXML2.IXMLDOMElement
Private curNode As MSXML2.IXMLDOMNode
Private ProcessNode As MSXML2.IXMLDOMElement
Private PracticeNode As MSXML2.IXMLDOMElement

Public Property Let SourceNode(ndNewvALUE As MSXML2.IXMLDOMNode)
   Set sessionRoot = ndNewvALUE
End Property

Public Property Get SourceNode() As MSXML2.IXMLDOMNode
   Set SourceNode = logRoot
End Property

Private Function ErrMsg(Number As String, Src As String, Desc As String) As MSXML2.IXMLDOMElement
   Dim logNode As MSXML2.IXMLDOMElement
   Dim nAtt As MSXML2.IXMLDOMAttribute
   Dim nTxt As MSXML2.IXMLDOMText

   With xmlLog
      Set logNode = .createElement("Error")
      
      Set nAtt = .createAttribute("TimeStamp")
      nAtt.Value = Format(Now(), "hh:mm:ss")
      logNode.setAttributeNode nAtt
      
      Set nAtt = .createAttribute("Number")
      nAtt.Value = Number
      logNode.setAttributeNode nAtt
      
      Set nAtt = .createAttribute("Source")
      nAtt.Value = Src
      logNode.setAttributeNode nAtt
      
      Set nTxt = .createTextNode(Desc)
      logNode.appendChild nTxt
   End With
   
   Set ErrMsg = logNode
End Function

Private Function Message(Data As String) As MSXML2.IXMLDOMElement
   Dim logNode As MSXML2.IXMLDOMElement
   Dim nAtt As MSXML2.IXMLDOMAttribute
   Dim nTxt As MSXML2.IXMLDOMText
   
   With xmlLog
      Set logNode = .createElement("Log")
      
      Set nAtt = .createAttribute("TimeStamp")
      nAtt.Value = Format(Now(), "hh:mm:ss")
      logNode.setAttributeNode nAtt
      
      Set nTxt = .createTextNode(Data)
      logNode.appendChild nTxt
   End With
   
   Set Message = logNode
End Function

Private Function Practice(NatCode As String) As MSXML2.IXMLDOMElement
   Dim logNode As MSXML2.IXMLDOMElement
   Dim nAtt As MSXML2.IXMLDOMAttribute
      
   With xmlLog
      Set logNode = .createElement("Practice")
      
      Set nAtt = .createAttribute("NatCode")
      nAtt.Value = NatCode
      logNode.setAttributeNode nAtt
   End With

   Set Practice = logNode
End Function

Private Function Process(RepCount As String, ModuleData As String) As MSXML2.IXMLDOMElement
   Dim logNode As MSXML2.IXMLDOMElement
   Dim nAtt As MSXML2.IXMLDOMAttribute
   
   With xmlLog
      Set logNode = .createElement("Process")
      
      Set nAtt = .createAttribute("StartTime")
      nAtt.Value = Format(Now(), "hh:mm:ss")
      logNode.setAttributeNode nAtt
      
      Set nAtt = .createAttribute("ReportCount")
      nAtt.Value = RepCount
      logNode.setAttributeNode nAtt
      
      Set nAtt = .createAttribute("Module")
      nAtt.Value = ModuleData
      logNode.setAttributeNode nAtt
      
   End With
   
   Set Process = logNode
End Function

Private Sub xLogMsg(Data As String, _
                  Optional NodeName As String = "", _
                  Optional Additional As String = "")
   Dim nArr() As String
   Dim logNode As MSXML2.IXMLDOMElement
   Dim nAtt As MSXML2.IXMLDOMAttribute
   
   nArr = Split(Additional, "~")
   
   If NodeName = "" Then
      Set logNode = Message(Data)
   Else
      If NodeName = curNode.NodeName Then
      
      Else
      End If
   End If
   
End Sub

Public Sub LogMsg(mType As EnmuLogTypes, logData As String, Optional logInfo As String = "")
   Dim tNode As MSXML2.IXMLDOMNode
   Dim nArr() As String
   
   Select Case mType
      Case 0   '  Process
         CheckSaved
         Set ProcessNode = Process(logInfo, logData)
         logRoot.appendChild ProcessNode
         Set curNode = ProcessNode
         
      Case 1   '  Practice
         CheckSaved
'         If Not PracticeNode Is Nothing Then
'            If logData <> PracticeNode.Attributes(0).nodeValue Then
'               ProcessNode.appendChild curNode
'            End If
'         End If
            
         Set PracticeNode = Practice(logData)
         Set curNode = PracticeNode
         
      Case 2   '  Log
         curNode.appendChild Message(logData)
      
      Case 3   '  Error
         nArr = Split(logInfo, "~")
         curNode.appendChild ErrMsg(nArr(0), nArr(1), logData)
         
   End Select
End Sub

Public Sub SaveLog()
   If curNode.NodeName = "Practice" Then
      ProcessNode.appendChild curNode
   End If
   
   'sessionRoot.appendChild logRoot
   
   .xmlLog.Save fs.BuildPath(App.Path, "testLog.xml")
End Sub

Private Sub DecodeAttributes(ByRef nodeId As MSXML2.IXMLDOMElement, AttList As String)
   Dim nArr() As String
   Dim attPair() As String
   Dim nAtt As MSXML2.IXMLDOMAttribute
   
   nArr = Split(AttList, "~")
   For i = 0 To UBound(nArr)
      attPair = Split(nArr(i), "=")
      
      If UBound(attPair) = 1 Then
         Set nAtt = xmlLog.createAttribute("Attribute_" & i)
         nAtt.Value = attPair(0)
      Else
         Set nAtt = xmlLog.createAttribute(attPair(0))
         nAtt.Value = attPair(1)
      End If
      
      nodeId.setAttributeNode nAtt
   Next i
   
End Sub

Public Sub CheckSaved()
   Dim tNode As MSXML2.IXMLDOMNode
   
   If Not curNode Is Nothing Then
      Set tNode = curNode
      Do Until tNode.parentNode Is Nothing
         Set tNode = tNode.parentNode
      Loop
      
      If tNode.NodeName <> logRoot.NodeName Then
         If tNode.NodeName = "Practice" Then
            ProcessNode.appendChild PracticeNode
         End If
               
         logRoot.appendChild ProcessNode
      End If
      
      Set PracticeNode = Nothing
      Set curNode = Nothing
   End If
   
End Sub

Private Sub Class_Initialize()
   Dim nAtt As MSXML2.IXMLDOMAttribute
   
   Set xmlLog = New MSXML2.DOMDocument
   
   Set logRoot = xmlLog.createElement("IceMsgLog")
   Set nAtt = xmlLog.createAttribute("LogDate")
   nAtt.Value = Format(Now(), "ddd dd mmm yyyy")
   
   logRoot.setAttributeNode nAtt
   
   xmlLog.appendChild logRoot
   xmlLog.Save fs.BuildPath(App.Path, "testLog.xml")
End Sub
