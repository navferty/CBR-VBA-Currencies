Attribute VB_Name = "Tutorial"
Option Explicit

' Tutorial for Pikabu users =)

Private Function LoadUsdItems() As Collection

    Dim resultXmlDocument As DOMDocument60
    Dim recordItem As IXMLDOMElement
    Dim col As New Collection
    Dim resultItem As CurrencyRecord
    
    Set resultXmlDocument = RequestGetXml()
    
    For Each recordItem In resultXmlDocument.LastChild.ChildNodes
        Set resultItem = New CurrencyRecord
        With resultItem
            .CurrencyCode = recordItem.Attributes.getNamedItem("Id").NodeValue
            .CurrencyDate = CDate(recordItem.Attributes.getNamedItem("Date").NodeValue)
            .CurrencyValue = CDec(recordItem.ChildNodes.Item(1).nodeTypedValue)
        End With
        col.Add resultItem
    Next

End Function

Private Function RequestGetXml() As DOMDocument60
    Dim XMLHTTP As New MSXML2.XMLHTTP60
    Dim requestUrl As String
    
    requestUrl = "http://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=02/03/2001&date_req2=14/03/2001&VAL_NM_RQ=R01235"
    XMLHTTP.Open "GET", requestUrl, False
    XMLHTTP.send
    
    Dim resultXmlDocument As DOMDocument60
    Set resultXmlDocument = XMLHTTP.responseXML
    
    Stop
    
    Set RequestGetXml = resultXmlDocument
End Function
