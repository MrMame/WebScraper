Attribute VB_Name = "EBY_CTR_Main"
Option Explicit


Public Function SearchEbay(strSearchTerm As String, strLocation As String, CategoryValue As Integer, RadiusValue As Integer)
    ' Create Logger and collect them inside the Multilogger
    Dim theLoggers As EBY_DBG_LOG_ILogger
    Set theLoggers = EBY_DBG_LOG_LoggerFactory.CreateFullMultiLogger

    Dim strWebsiteAddress As String
    strWebsiteAddress = "https://www.ebay-kleinanzeigen.de/"

    Dim AllAds As New Collection ' of cls_Ad
    
    Dim ResultPagesReader As New EBY_CTR_ResultPagesReader
    Call ResultPagesReader.EBY_DBG_LOG_ILoggable_SetLogger(theLoggers)
    
    Dim ResultPages As Collection ' of EBY_DAT_PAG_ResultPage
    Call ResultPagesReader.LoadResultPages(strWebsiteAddress, strSearchTerm, CategoryValue, strLocation, RadiusValue)
    Set ResultPages = ResultPagesReader.ResultPages

    Dim ResultPage As EBY_DAT_PAG_ResultPage
    For Each ResultPage In ResultPages

        Dim Ad As EBY_DAT_Ad
        For Each Ad In ResultPage.GetAds()
        
            AllAds.Add Ad
        
            Debug.Print "Ad.AdDate=" & Ad.AdDate
            Debug.Print "Ad.AdName=" & Ad.AdName
            Debug.Print "Ad.LinkAddress=" & Ad.LinkAddress
            Debug.Print "Ad.Location=" & Ad.Location
            Debug.Print "Ad.Price=" & Ad.Price
        Next
        Debug.Print "(" & ResultPage.AdsCount & ") Ads found."
      
    Next

    Dim theTableWriter As New EBY_CTR_TableWriter
    Call theTableWriter.EBY_DBG_LOG_ILoggable_SetLogger(theLoggers)
    Call theTableWriter.WriteAds(AllAds, "Data")

End Function



