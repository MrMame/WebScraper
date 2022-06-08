Attribute VB_Name = "mdl_main"
Option Explicit


Public Function SearchEbay()
    
    Dim strSearchTerm As String
    Dim strLocation As String
    Dim CategoryValue As Integer
    Dim Radius As eRadius
    
    Dim theCategories As New cls_eValCategories
    
    strSearchTerm = "roland mc707"
    strLocation = "65428 Rüsselsheim"
    CategoryValue = theCategories.APIValue_All
    Radius = eRadius.KM_200
    
    
    
    
    Dim strWebsiteAddress As String
    strWebsiteAddress = "https://www.ebay-kleinanzeigen.de/"
    


    Dim AllAds As New Collection ' of cls_Ad
    
    Dim ResultPagesReader As New cls_ResultPagesReader
    Dim ResultPages As Collection ' of cls_ResultPage
    Call ResultPagesReader.LoadResultPages(strWebsiteAddress, strSearchTerm, CategoryValue, strLocation, Radius)
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



