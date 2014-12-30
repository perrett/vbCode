Attribute VB_Name = "modClinicalLetter"
Option Explicit

Public Function GetLetter(iLetterIndex As Long, sHospitalNumber As String) As String

    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXml As String
    Dim strParam As String

    If blnUseHTTPS Then
       strUrl = "https://" & sDOMAIN & "/icedesktop/dotnet/ws/ClinicalLetterWebService/LetterExport.asmx"
    Else
       strUrl = "http://" & sDOMAIN & "/icedesktop/dotnet/ws/ClinicalLetterWebService/LetterExport.asmx"
    End If
    
    strSoapAction = "Ahsl.Ice.ClinicalLetters.WebService.LetterExport/GetLetter"

    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<GetLetter xmlns=""Ahsl.Ice.ClinicalLetters.WebService.LetterExport"">" & _
             "<letterIndex>" & CStr(iLetterIndex) & "</letterIndex>" & _
             "<language>1</language>" & _
             "<userIndex>" & defUserIndex & "</userIndex>" & _
             "<hospitalNumber>" & sHospitalNumber & "</hospitalNumber>" & _
             "</GetLetter>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
    
        ' Call PostWebservice and put result in text box
    GetLetter = DecodeXML(PostWebservice(strUrl, strSoapAction, strXml))

End Function

Private Function PostWebservice(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
    
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim lngPos1 As Long
    Dim lngPos2 As Long

    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Load XML
    objDom.async = False
    objDom.loadXML XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    
    ' Send XML command
    objXmlHttp.Send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responseText
    
    ' Close object
    Set objXmlHttp = Nothing
    
    ' Extract result
    lngPos1 = InStr(strRet, "<GetLetterResult>") + 17
    lngPos2 = InStr(strRet, "</GetLetterResult>")
    If lngPos1 > 7 And lngPos2 > 0 Then
        strRet = Mid(strRet, lngPos1, lngPos2 - lngPos1)
    End If
    
    ' Return result
    PostWebservice = strRet
    
Exit Function

Err_PW:
    PostWebservice = "Error: " & Err.Number & " - " & Err.Description

End Function

Private Function DecodeXML(xml As String) As String

    xml = Replace(xml, "&lt;", "<")
    xml = Replace(xml, "&gt;", ">")
    xml = Replace(xml, "&amp;", "&")

    DecodeXML = xml

End Function

