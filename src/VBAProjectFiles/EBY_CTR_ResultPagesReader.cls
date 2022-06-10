VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EBY_CTR_ResultPagesReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim m_colResultPages As Collection ' of ResultPages



Public Property Get ResultPages() As Collection ' of EBY_DAT_PAG_ResultPages
    Set ResultPages = m_colResultPages
End Property



Public Sub LoadResultPages(strWebsiteAddress As String, strSearchTerm As String, CategoryValue As Integer, strLocation As String, RadiusValue As Integer)

    Dim retColResultPages As New Collection

    Dim Html As HTMLDocument
    Set Html = EBY_CTR_EbayHtmlRequester.GetResultpage(strWebsiteAddress, strSearchTerm, CategoryValue, strLocation, RadiusValue)
    
    Dim ResultPage As New EBY_DAT_PAG_ResultPage
    Call ResultPage.LoadFromHTMLDocument(strWebsiteAddress, Html)
    
    Call retColResultPages.Add(ResultPage)
    
    Do While (Not ResultPage.NextResultPage Is Nothing)
        Set ResultPage = ResultPage.NextResultPage
        Call retColResultPages.Add(ResultPage)
    Loop
    
    Set m_colResultPages = retColResultPages
 
End Sub

