VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_CTR_ResultPageHTMLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements EBY_DBG_LOG_ILoggable

Dim m_logger As EBY_DBG_LOG_ILogger


Private Const HTML_CLASSNAME_AD_COLLECTION As String = "aditem-main"
Private Const HTML_CLASSNAME_AD_LOCATION As String = "aditem-main--top--left"
Private Const HTML_CLASSNAME_AD_DATE As String = "aditem-main--top--right"
Private Const HTML_CLASSNAME_AD_NAME As String = "ellipsis"
Private Const HTML_CLASSNAME_AD_PRICE As String = "aditem-main--middle--price"
Private Const HTML_CLASSNAME_NEXT_RESULTPAGE As String = "pagination-next"
Private Const STRING_IDENTIFIER_DATE_TODAY As String = "Heute"
Private Const STRING_IDENTIFIER_DATE_YESTERDAY As String = "Gestern"
Private Const STRING_NEGOTIABLE_TEXT As String = "VB"
Private Const STRING_AD_LOCATION_NOT_AVAILABLE As String = "NichtVerfügbar"

Public Function ReadResults(strWebsiteAddress As String, Html As HTMLDocument) As Collection

    Dim ResultAds As New Collection
    
    Dim AdDivs As IHTMLElementCollection
    Set AdDivs = Html.getElementsByClassName(HTML_CLASSNAME_AD_COLLECTION)

    Dim AdDiv As HTMLDivElement
    For Each AdDiv In AdDivs
        Dim NewAd As EBY_DAT_Ad
        Set NewAd = ReadAd(strWebsiteAddress, AdDiv)
        Call ResultAds.Add(NewAd)
    Next
    
    Call EBY_DBG_LOG_ILoggable_WriteToLogger("(" & ResultAds.Count & ") Ad's are parsed from address (" & strWebsiteAddress & ")")
    
    Set ReadResults = ResultAds

End Function

Public Function ReadNextResultPage(strWebsiteAddress As String, Html As HTMLDocument) As EBY_DAT_PAG_ResultPage
    
    
    Dim retNextResultPage As EBY_DAT_PAG_ResultPage
    
    Dim colElements As IHTMLElementCollection
    Set colElements = Html.getElementsByClassName(HTML_CLASSNAME_NEXT_RESULTPAGE)
    
    
    If (colElements.Length < 1) Then
        Set retNextResultPage = Nothing
    Else
        
        Dim NextResultLinkElement As HTMLLinkElement
        Set NextResultLinkElement = colElements(0)
        
        Dim strUrl As String
        strUrl = Strings.Replace(NextResultLinkElement, "about:/", strWebsiteAddress)
        
        Dim NextResultHtml As HTMLDocument
        Dim theHttpRequester As New EBY_UTL_HttpRequester
        Call theHttpRequester.EBY_DBG_LOG_ILoggable_SetLogger(m_logger)
        Call theHttpRequester.ReadHTMLDocument(strUrl, NextResultHtml)
        
        Set retNextResultPage = New EBY_DAT_PAG_ResultPage
        Call retNextResultPage.LoadFromHTMLDocument(strWebsiteAddress, NextResultHtml, Me)
        
    End If
    
    
    Set ReadNextResultPage = retNextResultPage
    
End Function

Private Function ReadAd(strWebsiteAddress As String, hResultDiv As HTMLDivElement) As EBY_DAT_Ad

    Dim Ad As New EBY_DAT_Ad
    Dim colElements As IHTMLElementCollection

    Set colElements = hResultDiv.getElementsByClassName(HTML_CLASSNAME_AD_LOCATION)
    Ad.Location = ReadAdLocationFromElement(colElements)

    Set colElements = hResultDiv.getElementsByClassName(HTML_CLASSNAME_AD_DATE)
    Ad.AdDate = ReadAdDateFromElement(colElements)

    Set colElements = hResultDiv.getElementsByClassName(HTML_CLASSNAME_AD_NAME)
    Ad.AdName = ReadAdNameFromElement(colElements)
    Ad.LinkAddress = ReadLinkAddressFromElement(colElements, strWebsiteAddress)
    
    Set colElements = hResultDiv.getElementsByClassName(HTML_CLASSNAME_AD_PRICE)
    Ad.Price = ReadPriceFromElement(colElements)
    Ad.Negotiable = ReadNegotiableFromElement(colElements)
    
    Set ReadAd = Ad
      
End Function


Private Function ReadAdLocationFromElement(colElements As IHTMLElementCollection) As String
    Dim retLocation As String
    If (colElements.Length < 1) Then
        retLocation = STRING_AD_LOCATION_NOT_AVAILABLE
    Else
        retLocation = colElements(0).innerText
    End If
    ReadAdLocationFromElement = retLocation
End Function

Private Function ReadAdDateFromElement(colElements As IHTMLElementCollection) As Date
    Dim retDate As Date
    If (colElements.Length < 1) Then
        retDate = CDate("01.01.1900")
    Else
        retDate = CastToDate(colElements(0).innerText)
    End If
    ReadAdDateFromElement = retDate
End Function

Private Function ReadAdNameFromElement(colElements As IHTMLElementCollection) As String
    Dim strName As String
    If (colElements.Length < 1) Then
        strName = ""
    Else
        Dim hLink As HTMLLinkElement
        Set hLink = colElements(0)
        strName = hLink.innerText
    End If
    ReadAdNameFromElement = strName
End Function

Private Function ReadLinkAddressFromElement(colElements As IHTMLElementCollection, strWebsiteAddress As String) As String
    Dim strLink As String
    If (colElements.Length < 1) Then
        strLink = ""
    Else
        Dim hLink As HTMLLinkElement
        Set hLink = colElements(0)
        strLink = Strings.Replace(hLink, "about:/", strWebsiteAddress)   ' Adding the Domain Name because the GET delivers "about:/" instead f the website domain
    End If
    ReadLinkAddressFromElement = strLink
End Function


Private Function ReadPriceFromElement(colElements As IHTMLElementCollection) As Single
    Dim sngPrice As Single
    If (colElements.Length < 1) Then
        sngPrice = 0
    Else
        Dim strPrice As String
        strPrice = colElements(0).innerText
        sngPrice = CastToSingle(strPrice)
    End If
    ReadPriceFromElement = sngPrice
End Function

Private Function ReadNegotiableFromElement(colElements As IHTMLElementCollection) As Boolean
    Dim retBool As Boolean
    If (colElements.Length < 1) Then
       retBool = False
    Else
        If (InStr(1, colElements(0).innerText, STRING_NEGOTIABLE_TEXT) > 0) Then
            retBool = True
        Else
            retBool = False
        End If
    End If
   ReadNegotiableFromElement = retBool
End Function


' Removes all Characters.
Private Function CastToSingle(strValue As String) As Single
    
    Dim strTmpVal As String

    Dim char As String
    Dim i As Integer
    For i = 1 To Len(strValue)
        char = Mid(strValue, i, 1)
        If (IsNumeric(char) Or char = ",") Then
            strTmpVal = strTmpVal & Mid(strValue, i, 1)
        End If
    Next
    On Error GoTo fehler
    CastToSingle = CSng(strTmpVal)
    GoTo ende
fehler:
    CastToSingle = 0
ende:
    
End Function



Private Function CastToDate(strDate As String) As Date

    Dim retDate As Date

    If (strDate = "") Then
        retDate = CDate("1.1.1900")
    Else
    
        Dim uDate As String
        uDate = UCase(strDate)
        
        If InStr(1, uDate, UCase(STRING_IDENTIFIER_DATE_TODAY)) > 0 Then
            retDate = Now
        ElseIf InStr(1, uDate, UCase(STRING_IDENTIFIER_DATE_YESTERDAY)) > 0 Then
            retDate = DateAdd("d", -1, Now)
        Else
            retDate = CDate(strDate)
        End If
    End If
    CastToDate = retDate

End Function



'   Interface: EBY_DBG_LOG_ILoggable
'
Public Sub EBY_DBG_LOG_ILoggable_SetLogger(Logger As EBY_DBG_LOG_ILogger)
    Set m_logger = Logger
End Sub
Private Sub EBY_DBG_LOG_ILoggable_WriteToLogger(Text As String)
    If (Not m_logger Is Nothing) Then m_logger.Log (Text)
End Sub

