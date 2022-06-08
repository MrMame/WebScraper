Attribute VB_Name = "mdl_main"
Option Explicit


Public Function SearchEbay()
    
    Dim strSearchTerm As String
    Dim strLocation As String
    Dim CategoryValue As Integer
    Dim RadiusValue As Integer
    
    Dim theCategories As New cls_ApiCategories
    Dim theRadius As New cls_ApiRadius
        
    strSearchTerm = "roland mc707"
    strLocation = "65428 Rüsselsheim"
    CategoryValue = theCategories.GetCategoryAPIValue("all")
    RadiusValue = theRadius.GetRadiusAPIValue("KM_5")
    
    
    
    
    Dim strWebsiteAddress As String
    strWebsiteAddress = "https://www.ebay-kleinanzeigen.de/"
    


    Dim AllAds As New Collection ' of cls_Ad
    
    Dim ResultPagesReader As New cls_ResultPagesReader
    Dim ResultPages As Collection ' of cls_ResultPage
    Call ResultPagesReader.LoadResultPages(strWebsiteAddress, strSearchTerm, CategoryValue, strLocation, RadiusValue)
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



