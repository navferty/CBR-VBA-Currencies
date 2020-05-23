Attribute VB_Name = "LoadCurrencies"
Option Explicit

Public Function GetTodayCurrency(currCode As String, volatileArg As Variant) As Variant 'decimal
    Dim currItem As CurrencyRecord
    Dim col As Collection
    Set col = GetCurrency(currCode, DateAdd("d", -1, Date), Now)
    
    If col.Count = 0 Then
        Exit Function
    End If
    
    Set currItem = col.Item(col.Count)
    GetTodayCurrency = currItem.CurrencyValue
    
    Debug.Print "Currency loaded, value is " + CStr(currItem.CurrencyValue)
End Function

Private Sub TestCurrLoad()
    Dim col As Collection
    Set col = GetCurrency("USD", #1/1/2020#, #2/2/2020#)
    Stop
End Sub

' returns Collection of CurrencyRecord's
Private Function GetCurrency(currCode As String, startDate As Date, endDate As Date) As Collection

    Dim codeDict As Dictionary
    Set codeDict = New Dictionary
    codeDict.Add "USD", "R01235"
    codeDict.Add "GBP", "R01035"
    codeDict.Add "BYN", "R01090B"
    ' наполнить другими значениями
    
    Dim currCbrCode As String
    currCbrCode = codeDict.Item(currCode)
    
    Dim requestUri As String
    requestUri = "http://www.cbr.ru/scripts/XML_dynamic.asp?" _
                + "date_req1=" + Format(startDate, "dd/mm/yyyy") _
                + "&date_req2=" + Format(endDate, "dd/mm/yyyy") _
                + "&VAL_NM_RQ=" + currCbrCode
    ' Date format is important!
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

