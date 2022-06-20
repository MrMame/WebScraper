# EBay-Kleinanzeigen Site Scraper

VBA-Excel Macro to collect Data from EBay-Kleinanzeigen.
At the moment, there is no free/public API available.
Collected Ads data will be presented inside an Excel table.

**Be adviced !
Ebay Kleinanzeigen AGB §5 forbid the usage of crawlers like this one.
You can find the german AGB's [here](https://themen.ebay-kleinanzeigen.de/nutzungsbedingungen/).**

# Installation

## Step1. Check if all necessary references are selected inside the workbook.
Open the reference manager inside the VBA IDE. Ensure that following references are checked :

- Microsoft XML, v3.0
- Microsoft HTML Object Library
- Microsoft Visual Basic for Applications Extensibility 5.3

![grafik](https://user-images.githubusercontent.com/51000524/174578399-2dea6a13-c7ff-4227-bdc6-aa87eedcb080.png)
![grafik](https://user-images.githubusercontent.com/51000524/174578719-2ad788e5-63a0-43f7-9aab-126cad3c2d12.png)

## Step2. Install the accUnit Test Framework
Got to [AccUnit](https://accunit.access-codelib.net/) Webpage and install version 0.9.10 of the accUnit Test Framework. After the installation, you will have new options inside the VBA IDE for testing. For further explanations, please read the documentation on the website itself.

![grafik](https://user-images.githubusercontent.com/51000524/174579550-ac45ca14-1ece-4279-9c1c-d06ca71d3b22.png)


# Dependencies
 - **Microsoft Visual Basic for Applications Extensibility 5.3** for having access to the VBA Object Model
 - Unit Test Framework is **[AccUnit](https://accunit.access-codelib.net/)** v0.9.10 
 - **Microsoft HTML Object Library** to have easy access to the requested HTML Documents
 - **Microsoft XML, v3.0** to do XML HTTP Requests
 - Used Microsoft Office Excel version was v14.0.7268.5000(32bit) (Microsoft Office Professional Plus 2010)

# Deploy
If you want to release the file for production usage, you have to remove the accUnit testcode and its references first.
The accunit add-in has a nice auto-remove feature. By clicking on **"Extras/AccUnit-de/Testumgebung entfernen"** inside 
the VBA-IDE, all test releated stuff gets removed. Be sure
to do so. Otherwise there will be reference errors on the customers PC because of missing accunit libraries.

![grafik](https://user-images.githubusercontent.com/51000524/174429022-08de955d-0cde-48e5-adf7-c591f6f3a6e5.png)



# Features
## v.0.1.0

You can enter the search query by an inputform.
![grafik](https://user-images.githubusercontent.com/51000524/173420075-a62c3883-e84e-47a0-960b-bf9062cd7bd9.png)


Your results will be written into an excel table inside a worksheet called "data".
- "Datum" is the ad's placing date.
- "Ort" is the ad's location.
- "Preis" is the price.
- "Verhandelbar" indicates if the ad's price is negotiable or not.
- "Name" is containing a link to the ad.

![grafik](https://user-images.githubusercontent.com/51000524/173420460-8cb2e0a3-a16d-4971-872e-4f589de10cad.png)

