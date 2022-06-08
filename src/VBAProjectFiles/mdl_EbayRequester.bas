Attribute VB_Name = "mdl_EbayRequester"
Option Explicit




'
'
'   The first Get request returns HTML State 301 (site moved) with the new address location.
'   The Search parameters of the second Get Request will be changed by the server itself.
'
'
'   Example of First GET Request
'       "https://www.ebay-kleinanzeigen.de/s-suchanfrage.html?keywords=mc707 roland&categoryId=73&locationStr=65428 Rüsselsheim&locationId=4816&radius=30&sortingField=SORTING_DATE&adType=&posterType=&pageNum=1&action=find&maxPrice=&minPrice="
'
'       "https://www.ebay-kleinanzeigen.de/s-suchanfrage.html?
'                    keywords=mc707 roland
'                    &categoryId=73
'                    &locationStr=65428 Rüsselsheim
'                    &locationId=4816
'                    &radius=30
'                    &sortingField=SORTING_DATE
'                    &adType=
'                    &posterType=
'                    &pageNum=1
'                    &action=find
'                    &maxPrice=
'                    &minPrice=
'
'
'   Example of the 2. following request
'       https://www.ebay-kleinanzeigen.de/s-immobilien/65428/c195l4816r30



'Public Enum eCategory
'    all = 0
'    to_give_away = 192
'End Enum
Public Enum eRadius
    KM_5 = 5
    KM_10 = 10
    KM_20 = 20
    KM_30 = 30
    KM_50 = 50
    KM_100 = 100
    KM_150 = 150
    KM_200 = 200
End Enum



Public Function GetHTMLDocument(strAdress As String) As HTMLDocument
    Dim Html As HTMLDocument, hTable As HTMLTable
    Set Html = New HTMLDocument
        
    Dim strGetRequest As String
    strGetRequest = strAdress
    
    Dim oXMLHTTP As New MSXML2.XMLHTTP
    With oXMLHTTP
        .Open "GET", strGetRequest, False
        .send
    Html.body.innerHTML = .responseText
    End With
    Set GetHTMLDocument = Html
End Function


Public Function GetResultpage(ByVal strWebsiteAddress As String, SearchTerm As String, CategoryValue As Integer, Location As String, LocationSearchArea As eRadius) As HTMLDocument
        
    Dim Html As HTMLDocument, hTable As HTMLTable
    Set Html = New HTMLDocument
    
    SearchTerm = Replace(SearchTerm, " ", "+")
    
    Dim strGetRequest As String
    strGetRequest = strWebsiteAddress & "s-suchanfrage.html?keywords=" & SearchTerm & "&categoryId=" & CategoryValue & "&locationStr=" & Location & "&locationId=&radius=" & LocationSearchArea & "&sortingField=SORTING_DATE&adType=&posterType=&pageNum=1&action=find&maxPrice=&minPrice="
    
    Dim oXMLHTTP As New MSXML2.XMLHTTP
    'With CreateObject("MSXML2.XMLHTTP")
    With oXMLHTTP
        .Open "GET", strGetRequest, False
        '.Open "GET", strWebsiteAddress & "s-" & SearchTerm & "/k0", False
        '.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT" 'to deal with potential caching
        .send
    
    
    
    Html.body.innerHTML = .responseText
    End With
    Set GetResultpage = Html
End Function
