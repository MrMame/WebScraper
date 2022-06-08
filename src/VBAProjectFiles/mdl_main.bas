Attribute VB_Name = "mdl_main"
Option Explicit


Public Function SearchEbay()
    
    Dim strSearchTerm As String
    Dim strLocation As String
    Dim Category As eCategory
    Dim Radius As eRadius
    
    strSearchTerm = "roland mc707"
    strLocation = "65428 Rüsselsheim"
    Category = eCategory.all
    Radius = eRadius.KM_200
    
    Dim strWebsiteAddress As String
    strWebsiteAddress = "https://www.ebay-kleinanzeigen.de/"
    


    Dim AllAds As New Collection ' of cls_Ad
    
    Dim ResultPagesReader As New cls_ResultPagesReader
    Dim ResultPages As Collection ' of cls_ResultPage
    Call ResultPagesReader.LoadResultPages(strWebsiteAddress, strSearchTerm, Category, strLocation, Radius)
    Set ResultPages = ResultPagesReader.ResultPages

    Dim ResultPage As cls_ResultPage
    For Each ResultPage In ResultPages

        Dim Ad As cls_Ad
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

    Call mdl_TableWriter.WriteAds(AllAds, "Data")


End Function



Public Function Obsolete_SearchEbay()
    
    Dim strSearchTerm As String
    Dim strLocation As String
    Dim Category As eCategory
    Dim Radius As eRadius
    strSearchTerm = "kinderfahrrad gelb"
    strLocation = "65428 Rüsselsheim"
    Category = eCategory.all
    Radius = eRadius.KM_150
    
    Dim strWebsiteAddress As String
    strWebsiteAddress = "https://www.ebay-kleinanzeigen.de/"
    
    Dim Html As HTMLDocument
    Set Html = mdl_EbayRequester.GetResultpage(strWebsiteAddress, strSearchTerm, Category, strLocation, Radius)

    Dim Ads As Collection
    Set Ads = mdl_HtmlParser.ReadResults(strWebsiteAddress, Html)

    Dim Ad As cls_Ad
    For Each Ad In Ads
        Debug.Print "Ad.AdDate=" & Ad.AdDate
        Debug.Print "Ad.AdName=" & Ad.AdName
        Debug.Print "Ad.LinkAddress=" & Ad.LinkAddress
        Debug.Print "Ad.Location=" & Ad.Location
        Debug.Print "Ad.Price=" & Ad.Price
    Next
    Debug.Print "(" & Ads.Count & ") Ads found."


    Call mdl_TableWriter.WriteAds(Ads, "Data")


End Function
