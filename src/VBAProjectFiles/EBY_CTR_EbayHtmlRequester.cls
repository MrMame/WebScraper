VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_CTR_EbayHtmlRequester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Implements EBY_DBG_LOG_ILoggable

Dim m_logger As EBY_DBG_LOG_ILogger

'
'Public Function GetHTMLDocument(strAdress As String) As HTMLDocument
'
'
'
'    Dim Html As HTMLDocument, hTable As HTMLTable
'    Set Html = New HTMLDocument
'
'    Dim strGetRequest As String
'    strGetRequest = strAdress
'
'    Dim oXMLHTTP As New MSXML2.XMLHTTP
'    With oXMLHTTP
'        .Open "GET", strGetRequest, False
'        .send
'    Html.body.innerHTML = .responseText
'    End With
'
'    EBY_DBG_LOG_ILoggable_WriteToLogger ("[EBY_CTR_EbayHtmlRequester] Requested HTMLDocument from address (" & strAdress & "). Length of Responsetext was (" & Len(Html.body.innerHTML) & ") Characters")
'
'    Set GetHTMLDocument = Html
'End Function


Public Function GetResultpage(ByVal strWebsiteAddress As String, SearchTerm As String, CategoryValue As Integer, Location As String, RadiusValue As Integer) As HTMLDocument
        
    Dim Html As HTMLDocument, hTable As HTMLTable
    Set Html = New HTMLDocument
    
    SearchTerm = Replace(SearchTerm, " ", "+")
    
    Dim strGetRequest As String
    strGetRequest = strWebsiteAddress & "s-suchanfrage.html?keywords=" & SearchTerm & "&categoryId=" & CategoryValue & "&locationStr=" & Location & "&locationId=&radius=" & RadiusValue & "&sortingField=SORTING_DATE&adType=&posterType=&pageNum=1&action=find&maxPrice=&minPrice="
    
    Dim oXMLHTTP As New MSXML2.XMLHTTP
    With oXMLHTTP
        .Open "GET", strGetRequest, False
        .send
        Html.body.innerHTML = .responseText
        EBY_DBG_LOG_ILoggable_WriteToLogger ("[EBY_CTR_EbayHtmlRequester] oXMLHTTP.status = (" & .Status & ") oXMLHTTP.statusText=(" & .statusText & ")")
    End With
    
    EBY_DBG_LOG_ILoggable_WriteToLogger ("[EBY_CTR_EbayHtmlRequester] Requested Resultpage HTMLDocument from address (" & strWebsiteAddress & "). Lentgh of Responsetext was (" & Len(Html.body.innerHTML) & ") characters. Get-Request text was (" & strGetRequest & ")")

    
    Set GetResultpage = Html
End Function

'   Interface: EBY_DBG_LOG_ILoggable
'
Public Sub EBY_DBG_LOG_ILoggable_SetLogger(Logger As EBY_DBG_LOG_ILogger)
    Set m_logger = Logger
End Sub
Private Sub EBY_DBG_LOG_ILoggable_WriteToLogger(Text As String)
    If (Not m_logger Is Nothing) Then m_logger.Log (Text)
End Sub
