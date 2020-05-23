Attribute VB_Name = "LoadOrganization"
Option Explicit

Private Const AuthToken As String = "<HERE YOU PUT YOUR DADATA API KEY>"
Private Const DadataUrl As String = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"

Public Function LoadOrgName(query As String) As String
    Dim org As LoadedOrganization
    Set org = LoadOrganization(query)
    LoadOrgName = org.OrgName
End Function

Private Sub TestDadataApi()
      Dim org As LoadedOrganization
      Set org = LoadOrganization("7707083893")
      Stop
End Sub

Private Function LoadOrganization(query As String) As LoadedOrganization
    Dim doc As DOMDocument60
    Set doc = RequestPostJson(DadataUrl, "{ ""query"": """ + query + """ }", AuthToken)
    Set LoadOrganization = New LoadedOrganization
    With LoadOrganization
        .OrgName = doc.SelectSingleNode("//suggestions/value").Text
        .OrgTin = doc.SelectSingleNode("//data/inn").Text
    End With
    Debug.Print "Organization loaded: " + LoadOrganization.OrgName
End Function

Private Function RequestPostJson(reqUrl As String, reqBody As String, AuthToken As String) As DOMDocument60
    Dim XMLHTTP As New MSXML2.XMLHTTP60
    Dim requestUrl As String
    With XMLHTTP
        .Open "POST", reqUrl, False
        .setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        .setRequestHeader "Authorization", "Token " + AuthToken
        .setRequestHeader "Content-Type", "application/json; charset=UTF-8"
        .setRequestHeader "Accept", "application/xml; charset=UTF-8"
        .send reqBody
        Set RequestPostJson = .responseXML
    End With
End Function
