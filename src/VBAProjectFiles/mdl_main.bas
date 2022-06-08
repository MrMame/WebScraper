Attribute VB_Name = "mdl_main"
Option Explicit


Public Function SearchEbay(strSearchTerm As String, strLocation As String, CategoryValue As Integer, RadiusValue As Integer)

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



