VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDTSRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************
'  Responses from the DTS Server
'***********************************

Private fromDTS As String
Private dtsURL As String
Private reqStatus As String
Private dtsStatus As String
Private fromSMTP As String
Private toSMTP As String
Private sentTo As String
Private DTSMsgId As String
Private DTSProcess As String
Private DTSWorkflow As String
Private transferDate As String
Private transferStatus As String
Private sentDate As String
Private FileId As String
Private DTSLocalId As String
Private timeStamp As Date

Private rsc As clsDTSEvents
Private evSignal As clsDTSSignal

Private DOMOutput As String
Private xmlDoc As MSXML2.DOMDocument
Private xmlHTTP As MSXML2.XMLHTTP40

Public Sub Abort()
   Set rsc = Nothing
End Sub

Public Property Let CurrentStatus(strNewValue As String)
   transferStatus = strNewValue
End Property

Public Property Get CurrentStatus() As String
   CurrentStatus = transferStatus
End Property

Public Property Get DTSResponse() As String
   DTSResponse = dtsStatus
End Property

Public Property Let DTSServer(strNewValue As String)
   dtsURL = strNewValue
End Property

Private Sub EvaluateResponse(xmlData As String)
   On Error GoTo procEH
   Dim blnLoadOK As Boolean
   Dim cNode As MSXML2.IXMLDOMNode
   Dim i As Integer
   
   If fs.FileExists(xmlData) Then
      blnLoadOK = xmlDoc.Load(xmlData)
   Else
      blnLoadOK = xmlDoc.loadXML(xmlData)
   End If
   
   If blnLoadOK Then
      For Each cNode In xmlDoc.selectNodes("/DTSResponse/Response")
         dtsStatus = cNode.selectSingleNode("ResponseStatus").firstChild.Text
         
         If cNode.selectSingleNode("ResponseStatus").firstChild.Text = "OK" Then
            DTSMsgId = cNode.selectSingleNode("Details/MsgID").firstChild.Text
            sentTo = cNode.selectSingleNode("Details/ToDTSName").firstChild.Text
            fromSMTP = cNode.selectSingleNode("Details/FromSMTPAddr").firstChild.Text
            toSMTP = cNode.selectSingleNode("Details/ToSMTPAddr").firstChild.Text
            transferDate = Replace(cNode.selectSingleNode("Details/TransferDateTime").firstChild.Text, "T", " @ ")
            
            transferStatus = cNode.selectSingleNode("Details/Status/StatusText").firstChild.Text
            
            If Not cNode.selectSingleNode("Details/SentDateTime") Is Nothing Then
               sentDate = Replace(cNode.selectSingleNode("Details/SentDateTime").firstChild.Text, "T", " @ ")
            Else
               sentDate = "Not collected"
            End If
            
            FileId = cNode.selectSingleNode("Details/Subject").firstChild.Text
            
            If Not cNode.selectSingleNode("Details/ProcessID").firstChild Is Nothing Then
               DTSProcess = cNode.selectSingleNode("Details/ProcessID").firstChild.Text
            End If
            
            If Not cNode.selectSingleNode("Details/WorkflowID").firstChild Is Nothing Then
               DTSWorkflow = cNode.selectSingleNode("Details/WorkflowID").firstChild.Text
            End If
         End If
      Next
   Else
      MsgBox "Invalid XML from DTS Server"
   End If
   Exit Sub
   
procEH:
   Stop
   Resume
End Sub

Public Property Get FileData() As String
   FileData = FileId
End Property

Public Property Get FromSMTPAddress() As String
   FromSMTPAddress = fromSMTP
End Property

Public Property Get HTTPStatus() As Long
   HTTPStatus = xmlHTTP.ReadyState
End Property

Public Property Let LocalId(strNewValue As String)
   DTSLocalId = strNewValue
End Property

Public Property Get LocalId() As String
   LocalId = DTSLocalId
End Property

Public Property Get MsgId() As String
   MsgId = DTSMsgId
End Property
                          
Public Property Let PrepareEnquiry(LocalId As String)
   On Error GoTo procEH
   Dim domRoot As MSXML2.IXMLDOMElement
   Dim cNode As MSXML2.IXMLDOMNode
   Dim pNode As MSXML2.IXMLDOMNode
   Dim NPI As MSXML2.IXMLDOMProcessingInstruction
   
   Dim nAtt As MSXML2.IXMLDOMAttribute
   Dim nTxt As MSXML2.IXMLDOMText
   
   Set xmlDoc = Nothing
   Set xmlDoc = New MSXML2.DOMDocument
   
   DTSLocalId = LocalId
   
   With xmlDoc
      Set NPI = .createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
      .appendChild NPI
   
      Set domRoot = .createNode(NODE_ELEMENT, "DTSRequest", "")
      .appendChild domRoot
   
      Set nAtt = .createAttribute("Version")
      nAtt.value = "1.0"
      domRoot.setAttributeNode nAtt
   
      Set nAtt = .createAttribute("xmlns:xsi")
      nAtt.value = "http://www.w3.org/2001/XMLSchema-instance"
      domRoot.setAttributeNode nAtt
   
      Set nAtt = .createAttribute("xsi:noNamespaceSchemaLocation")
      nAtt.value = "D:\NHSRegSys\Backend\www\Schema\dts request.xsd"
      domRoot.setAttributeNode nAtt
   
      Set pNode = .createNode(NODE_ELEMENT, "Message", "")
      
      Set cNode = .createNode(NODE_ELEMENT, "FromDTSName", "")
      Set nTxt = .createTextNode(fromDTS)
      cNode.appendChild nTxt
      
      pNode.appendChild cNode
      
      Set cNode = .createNode(NODE_ELEMENT, "LocalID", "")
      Set nTxt = .createTextNode(LocalId)
      cNode.appendChild nTxt
      
      pNode.appendChild cNode
      
      domRoot.appendChild pNode
      
   End With
   
   serializeDOM xmlDoc.documentElement
   reqStatus = "Prepared"
   
   transferStatus = "Queued - too many outstanding requests"
   Exit Property
   
procEH:
   Stop
   Resume
End Property

Public Property Get ProcessId() As String
   ProcessId = DTSProcess
End Property

Public Sub ProcessResponse()
   With xmlHTTP
      reqStatus = "Received"
      
      Select Case .status
         Case Is < 200
            transferStatus = .status & " (Redirection/Continue from server)"

         Case Is < 301
            EvaluateResponse .responseText

         Case Is < 400
            transferStatus = .status & " (" & .responseText & ")"

         Case Is < 500
            transferStatus = .status & ": " & .StatusText

         Case Else
            transferStatus = "Server error (" & .status & "; " & .StatusText & ")"

      End Select
   End With
   
   timeStamp = Now()
   Set xmlHTTP = Nothing
   evSignal.Notify DTSLocalId
   DoEvents
End Sub

Public Property Get Recipient() As String
   Recipient = sentTo
End Property

Public Property Get RequestStatus() As String
   If DateDiff("N", timeStamp, Now()) < 10 Then
      RequestStatus = reqStatus
   Else
      RequestStatus = "Prepared"
   End If
End Property

Public Sub SendDTSRequest()
   On Error GoTo procEH
   
   eClass.FurtherInfo = "Preparing HTTP Request object"
   Set xmlHTTP = New MSXML2.ServerXMLHTTP40
   
   DoEvents
   
'  Set up the event handler
   eClass.FurtherInfo = "Setting up event handler"
   Set rsc = New clsDTSEvents
   rsc.CallingClass = Me
   xmlHTTP.OnReadyStateChange = rsc
   
   If dtsURL <> "" Then
      With xmlHTTP
'        Create an asynchronous connection to the server
         .Open "post", dtsURL, True ', dtsUser, dtsPass
        
'        Amend the XML output. The server expects every line to terminate with a cr/lf
         DoEvents
         
         .Send DOMOutput
      End With
      
      transferStatus = "Data requested - awaiting response from DTS Server"
   Else
      transferStatus = "The DTS Server address is not set up in config.ini"
   End If
   
   reqStatus = "Awaiting response"
   
   DoEvents
   Exit Sub
   
procEH:
   If eClass.Behaviour = -1 Then
      Stop
      Resume
   End If
   eClass.CurrentProcedure = "clsDTSRequest.SendDTSRequest"
   eClass.Add Err.Number, Err.Description, Err.Source
End Sub

Public Property Get SentOn() As String
   SentOn = sentDate
End Property

Private Sub serializeDOM(Node As IXMLDOMNode, _
                         Optional ByRef iLastNodeType As MSXML2.DOMNodeType = NODE_DOCUMENT, _
                         Optional ByVal iLevel As Integer = 0)
    
   Dim nAtt As MSXML2.IXMLDOMAttribute
   Dim attMap As MSXML2.IXMLDOMNamedNodeMap
   Dim child As IXMLDOMNode
   
'  process current node here
'  use a select statement on node.nodeType
'  for each node type, serialize the node to domOutput appropriately
   
   If TypeName(Node.ParentNode) = "DOMDocument" Then
      
      DOMOutput = "<?" & Node.ParentNode.firstChild.nodeName & " " & Node.ParentNode.firstChild.nodeValue & "?>" + vbCrLf
   End If
   
   Select Case Node.NodeType
    
'     for example, for NODE_ELEMENTs serialize the
'     begin tag including all attributes
      Case NODE_ELEMENT
         If (iLastNodeType = NODE_ELEMENT) Then
            DOMOutput = DOMOutput + vbCrLf
         End If
         
         DOMOutput = DOMOutput & Space(iLevel * 2) & "<" & Node.nodeName
         iLevel = iLevel + 1
         
         Set attMap = Node.Attributes
            
         If Not attMap Is Nothing Then
            For Each nAtt In attMap
               DOMOutput = DOMOutput & " " & nAtt.nodeName & "='" & nAtt.nodeValue & "'"
            Next
         End If
            
         DOMOutput = DOMOutput & ">"
                        
      Case NODE_PROCESSING_INSTRUCTION
         If (Len(DOMOutput) > 0) Then
            DOMOutput = DOMOutput + vbCrLf
         End If
         
         DOMOutput = DOMOutput & Space(iLevel * 2) & "<?" & Node.nodeName & " " & Node.nodeValue & "?>" + vbCrLf
    
      Case NODE_CDATA_SECTION
         DOMOutput = DOMOutput & "<![CDATA[" & Node.nodeValue & "]]>"
            
      Case NODE_COMMENT
         DOMOutput = DOMOutput & Space(iLevel * 2) & "<!--" & Node.nodeValue & "-->" + vbCrLf
            
      Case NODE_TEXT
         DOMOutput = DOMOutput & Node.nodeValue
            
   End Select
   iLastNodeType = Node.NodeType
    
' recursively process children
' call serializeDOM again for each child
    
   If Node.hasChildNodes Then
      For Each child In Node.childNodes
         serializeDOM child, iLastNodeType, iLevel
      Next
   End If
    
' for NODE_ELEMENTs, serialize the end tag here
   If (iLastNodeType = NODE_ELEMENT) Then
      DOMOutput = DOMOutput & vbCrLf + Space((iLevel - 1) * 2)
   End If
    
   If Node.NodeType = NODE_ELEMENT Then
      DOMOutput = DOMOutput & "</" & Node.nodeName & ">"
      iLevel = iLevel - 1
      iLastNodeType = NODE_ELEMENT
   End If
End Sub

Public Property Let SignalClass(oClass As clsDTSSignal)
   Set evSignal = oClass
End Property

Public Property Let SiteDTS(strNewValue As String)
   fromDTS = strNewValue
End Property

Public Property Get ToSMTPAddress() As String
   ToSMTPAddress = toSMTP
End Property

Public Property Get TransferredOn() As String
   TransferredOn = transferDate
End Property

Public Property Get WorkFlowId() As String
   WorkFlowId = DTSWorkflow
End Property
