# EBay-Kleinanzeigen Site Scraper

VBA-Excel Macro to collect Data from EBay-Kleinanzeigen.
At the moment, there is no free/public API available.
Collected Ads data will be presented inside an Excel table.

**Be adviced !
Ebay Kleinanzeigen AGB §5 forbid the usage of crawlers like this one.
You can find the german AGB's [here](https://themen.ebay-kleinanzeigen.de/nutzungsbedingungen/).**



# Dependencies
 - **Microsoft Visual Basic for Applications Extensibility 5.3** for having access to the VBA Object Model
 - Testing Framework is [AccUnit](https://accunit.access-codelib.net/) for unit testsing. Simply VBUnitFramework v3.0 is used by AccUnit internally. 
 - **Microsoft HTML Object Library** to have easy access to the requested HTML Documents.
 - **Microsoft XML, v3.0** to do XML HTTP Requests.



v.0.1.0

You can enter the search query by an inputform.
![grafik](https://user-images.githubusercontent.com/51000524/173420075-a62c3883-e84e-47a0-960b-bf9062cd7bd9.png)


Your results will be written into an excel table inside a worksheet called "data".
- "Datum" is the ad's placing date.
- "Ort" is the ad's location.
- "Preis" is the price.
- "Verhandelbar" indicates if the ad's price is negotiable or not.
- "Name" is containing a link to the ad.

![grafik](https://user-images.githubusercontent.com/51000524/173420460-8cb2e0a3-a16d-4971-872e-4f589de10cad.png)

