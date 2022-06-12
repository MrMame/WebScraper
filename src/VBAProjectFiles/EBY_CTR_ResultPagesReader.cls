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

Implements EBY_DBG_LOG_ILoggable


Dim m_logger As EBY_DBG_LOG_ILogger

Dim m_colResultPages As Collection ' of ResultPages

Public Property Get ResultPages() As Collection ' of EBY_DAT_PAG_ResultPages
    Set ResultPages = m_colResultPages
End Property



Public Sub LoadResultPages(strWebsiteAddress As String, strSearchTerm As String, CategoryValue As Integer, strLocation As String, RadiusValue As Integer)

    Call EBY_DBG_LOG_ILoggable_WriteToLogger("Load Resultpage from address (" & strWebsiteAddress & ")." & _
                                            "Searchterm was (" & strSearchTerm & ")" & _
                                            "CategoryValue was (" & CategoryValue & ")" & _
                                            "strLocation was (" & strLocation & ")" & _
                                            "RadiusValue was (" & RadiusValue & ")")

    Dim retColResultPages As New Collection

    Dim Html As HTMLDocument
    Dim theHtmlRequester As New EBY_CTR_EbayHtmlRequester
    Call theHtmlRequester.EBY_DBG_LOG_ILoggable_SetLogger(m_logger)
    Set Html = theHtmlRequester.GetResultpage(strWebsiteAddress, strSearchTerm, CategoryValue, strLocation, RadiusValue)
    
    Dim ResultPage As New EBY_DAT_PAG_ResultPage
    Call ResultPage.LoadFromHTMLDocument(strWebsiteAddress, Html)
    
    Call retColResultPages.Add(ResultPage)
    
    Do While (Not ResultPage.NextResultPage Is Nothing)
        Set ResultPage = ResultPage.NextResultPage
        Call retColResultPages.Add(ResultPage)
    Loop
    
    Call EBY_DBG_LOG_ILoggable_WriteToLogger("Number of ResultPages found (" & retColResultPages.Count & ")")
    
    Set m_colResultPages = retColResultPages
 
End Sub

'   Interface: EBY_DBG_LOG_ILoggable
'
Public Sub EBY_DBG_LOG_ILoggable_SetLogger(Logger As EBY_DBG_LOG_ILogger)
    Set m_logger = Logger
End Sub
Private Sub EBY_DBG_LOG_ILoggable_WriteToLogger(Text As String)
    If (Not m_logger Is Nothing) Then m_logger.Log (Text)
End Sub

