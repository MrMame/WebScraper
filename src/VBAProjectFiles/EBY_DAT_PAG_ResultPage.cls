VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_DAT_PAG_ResultPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_HtmlDocument As HTMLDocument
Dim m_NextResultPage As EBY_DAT_PAG_ResultPage
Dim m_Ads As New Collection ' of cls_Ad

Public Property Get NextResultPage() As EBY_DAT_PAG_ResultPage
    Set NextResultPage = m_NextResultPage
End Property

Public Property Get AdsCount() As Integer
    AdsCount = m_Ads.Count
End Property


Public Function GetAds() As Collection
    Set GetAds = m_Ads
End Function


Public Sub LoadFromHTMLDocument(ByVal strWebsiteAddress As String, Html As HTMLDocument)
    Set m_HtmlDocument = Html
    Dim theHtmlParser As New EBY_CTR_ResultPageHTMLParser
    Set m_Ads = theHtmlParser.ReadResults(strWebsiteAddress, Html)
    Set m_NextResultPage = theHtmlParser.ReadNextResultPage(strWebsiteAddress, Html)
    
    Debug.Print "EBY_DAT_PAG_ResultPage.LoadFromHTMLDocument() -> m_Ads.count=" & m_Ads.Count
    
    
End Sub
