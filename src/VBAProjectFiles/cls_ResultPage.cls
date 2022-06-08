VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ResultPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_HtmlDocument As HTMLDocument
Dim m_NextResultPage As cls_ResultPage
Dim m_Ads As New Collection ' of cls_Ad

Public Property Get NextResultPage() As cls_ResultPage
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
    Set m_Ads = mdl_HtmlParser.ReadResults(strWebsiteAddress, Html)
    Set m_NextResultPage = mdl_HtmlParser.ReadNextResultPage(strWebsiteAddress, Html)
    
    Debug.Print "cls_ResultPage.LoadFromHTMLDocument() -> m_Ads.count=" & m_Ads.Count
    
    
End Sub
