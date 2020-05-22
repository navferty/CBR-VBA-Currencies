Attribute VB_Name = "LoadCurrencies"
Option Explicit


' returns Collection on CurrencyRecord
Public Function GetCurrency(currCode As String, startDate As Date, endDate As Date) As Collection

    ' TODO move to static variable
    Dim codeDict As Dictionary
    Set codeDict = New Dictionary
    codeDict.Add "USD", "R01235"
    ' наполнить другими значениями
    
    Dim currCbrCode As String
    currCbrCode = codeDict.Item(currCode)
    
    Dim requestUri As String
    requestUri = "http://www.cbr.ru/scripts/XML_dynamic.asp?" _
                + "date_req1=" + Format(startDate, "dd/mm/yyyy") _
                + "&date_req2=" + Format(endDate, "dd/mm/yyyy") _
                + "&VAL_NM_RQ=" + currCbrCode
    ' формат даты решает!
    ' http://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=02/03/2020&date_req2=14/03/2020&VAL_NM_RQ=R01235
    
    Dim resp As DOMDocument60
    Set resp = RequestGetXml(requestUri)
    
    Dim col As New Collection
    Dim recordItem As IXMLDOMElement
    Dim resultItem As CurrencyRecord
    For Each recordItem In resp.LastChild.ChildNodes
        Set resultItem = New CurrencyRecord
        With resultItem
            .CurrencyCode = currCode
            .CurrencyDate = CDate(recordItem.Attributes.getNamedItem("Date").NodeValue)
            .CurrencyValue = CDec(recordItem.ChildNodes.Item(1).nodeTypedValue)
        End With
        col.Add resultItem
    Next
    
    Set GetCurrency = col
End Function


Private Function RequestGetXml(requestUrl As String) As DOMDocument60
    'about requests in VBA: https://codingislove.com/http-requests-excel-vba/
    Dim XMLHTTP As New MSXML2.XMLHTTP60
    
    XMLHTTP.Open "GET", requestUrl, False
    XMLHTTP.send
    Set RequestGetXml = XMLHTTP.responseXML
End Function

'Sample:
' http://www.cbr.ru/scripts/XML_dynamic.asp?date_req1=02/03/2020&date_req2=14/03/2020&VAL_NM_RQ=R01235

'<?xml version="1.0" encoding="windows-1251"?>
'<ValCurs ID="R01235" DateRange1="17.04.2004" DateRange2="27.04.2004" name="Foreign Currency Market Dynamic">
'  <Record Date="17.04.2004" Id="R01235">
'    <Nominal>1</Nominal>
'    <Value>28,6223</Value>
'  </Record>
'  <Record Date="20.04.2004" Id="R01235">
'    <Nominal>1</Nominal>
'    <Value>28,6693</Value>
'  </Record>
'  <Record Date="21.04.2004" Id="R01235">
'    <Nominal>1</Nominal>
'    <Value>28,7662</Value>
'  </Record>
'  <Record Date="22.04.2004" Id="R01235">
'    <Nominal>1</Nominal>
'    <Value>28,9237</Value>
'  </Record>
'  <Record Date="23.04.2004" Id="R01235">
'    <Nominal>1</Nominal>
'    <Value>28,9800</Value>
'  </Record>
'  <Record Date="24.04.2004" Id="R01235">
'    <Nominal>1</Nominal>
'    <Value>28,9671</Value>
'  </Record>
'  <Record Date="27.04.2004" Id="R01235">
'    <Nominal>1</Nominal>
'    <Value>29,0033</Value>
'  </Record>
'</ValCurs>

