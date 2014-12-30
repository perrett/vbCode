VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oElement(255) As IXMLDOMElement
Private oCDATAElement As IXMLDOMCDATASection
Private ArrayItem As Integer
Private ElementActive As Boolean

Public Sub Add(ByVal iLevel As Integer, ByVal sName As String, Optional ByVal sValue As String = "!NOVALUE!")

    If Not RootNodeLoaded Then
        Err.Raise vbObjectError + 32767, "AHSLXML", "Root Node not loaded into class"
    Else
        If iLevel = 0 Then
            Set oElement(iLevel) = oXMLMessage.documentElement.appendChild(oXMLMessage.createElement(sName))
        Else
            Set oElement(iLevel) = oElement(iLevel - 1).appendChild(oXMLMessage.createElement(sName))
        End If

        If sValue <> "!NOVALUE!" Then
            oElement(iLevel).Text = sValue
        End If
        
        ArrayItem = iLevel
        ElementActive = True
    End If

End Sub

Public Sub SetAttribute(ByVal sAName As String, Optional ByVal sAValue As String = "")

    If Not RootNodeLoaded Then
        Err.Raise vbObjectError + 32767, "AHSLXML", "Root Node not loaded into class"
    Else
        If Not ElementActive Then
            Err.Raise vbObjectError + 32768, "AHSLXML", "No Element Active in document"
        Else
            oElement(ArrayItem).SetAttribute sAName, sAValue
        End If
    End If

End Sub

Public Sub AddCDATA(ByVal iLevel As Integer, ByVal sName As String, ByVal sValue As String)

    If Not RootNodeLoaded Then
        Err.Raise vbObjectError + 32767, "AHSLXML", "Root Node not loaded into class"
    Else
        If iLevel = 0 Then
            Set oElement(iLevel) = oXMLMessage.documentElement.appendChild(oXMLMessage.createElement(sName))
            Set oCDATAElement = oElement(iLevel).appendChild(oXMLMessage.createCDATASection(sValue))
        Else
            Set oElement(iLevel) = oElement(iLevel - 1).appendChild(oXMLMessage.createElement(sName))
            Set oCDATAElement = oElement(iLevel).appendChild(oXMLMessage.createCDATASection(sValue))
        End If
        
        ArrayItem = iLevel
        ElementActive = True
    End If

End Sub
