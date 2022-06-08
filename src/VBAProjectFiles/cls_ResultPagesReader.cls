VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ResultPagesReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim m_colResultPages As Collection ' of ResultPages



Public Property Get ResultPages() As Collection ' of cls_ResultPages
    Set ResultPages = m_colResultPages
End Property



Public Sub LoadResultPages(strWebsiteAddress As String, strSearchTerm As String, CategoryValue As Integer, strLocation As String, Radius As eRadius)

    Dim retColResultPages As New Collection

    Dim Html As HTMLDocument
    Set Html = mdl_EbayRequester.GetResultpage(strWebsiteAddress, strSearchTerm, CategoryValue, strLocation, Radius)
    
    Dim ResultPage As New cls_ResultPage
    Call ResultPage.LoadFromHTMLDocument(strWebsiteAddress, Html)
    
    Call retColResultPages.Add(ResultPage)
    
    Do While (Not ResultPage.NextResultPage Is Nothing)
        Set ResultPage = ResultPage.NextResultPage
        Call retColResultPages.Add(ResultPage)
    Loop
    
    Set m_colResultPages = retColResultPages
 
End Sub

