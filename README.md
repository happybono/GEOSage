# GEOSage <br> <img src="https://github.com/happybono/GEOSage/blob/master/Resources/powered_by_msexcel_on_white.png" alt="Powered by MSExcel logo" width="217"/>

This VBA module provides functions to interact with the Google Geocoding API. It can perform both addresses to latitude / longitude conversion (geocoding) and latitude / longitude to address conversion (reverse geocoding). The code supports various types of API plans, including free, Enterprise, and Premium.

**Please note : The Google Maps Platform Premium Plan is no longer available for sign-up or new customers since November 1, 2018.**

<div align="center">
<img alt="GitHub Last Commit" src="https://img.shields.io/github/last-commit/happybono/GEOSage"> 
<img alt="GitHub Repo Size" src="https://img.shields.io/github/repo-size/happybono/GEOSage">
<img alt="GitHub Repo Languages" src="https://img.shields.io/github/languages/count/happybono/GEOSage">
<img alt="GitHub Top Languages" src="https://img.shields.io/github/languages/top/HappyBono/GEOSage">
</div>

## What's New
### February 05, 2020
>[Initial release.](https://dev.azure.com/happybono/FinedustMonitorWithGPS/_versionControl?path=%24%2FFinedustMonitorWithGPS%2FMaps%2FSpreadSheet%2FReverseGeocoding.vb)

### February 11, 2020
> [Released as a standalone from the [FineDustMonitorWithGPS] project.](https://dev.azure.com/happybono/GEOSage)

### February 26, 2020
> [Performance improvements (up to 2× as faster than before) in the ADDRGEOCODE function.](https://dev.azure.com/happybono/GEOSage/_versionControl?path=%24%2FGEOSage%2FGEOSage.vb&line=85&lineStyle=plain&lineEnd=112&lineStartColumn=1&lineEndColumn=15)<br> <br>
> [Now supports Unicode using Microsoft ActiveX Data Objects Library in the ADDRGEOCODE function.](https://dev.azure.com/happybono/GEOSage/_versionControl?path=%24%2FGEOSage%2FGEOSage.vb&line=115&lineStyle=plain&lineEnd=152&lineStartColumn=1&lineEndColumn=1)

### February 27, 2020
> [Added GEOSage sample files.](https://dev.azure.com/happybono/GEOSage/_versionControl) <br>
> GEOSage sample includes Excel files that use demonstation data using Google Maps Geocoding API Key. The 
API key used in this project for geocoding and reverse geocoding feature is not provided for your use. 
The mock data demonstrates all functions with static result values as Google Maps geocoding API Key and 
VBA Add-in code are not included in the GEOSage sample.

### March 03, 2020
> [Added GEOSage.bas file to support directly import from Microsoft Excel.](https://dev.azure.com/happybono/GEOSage/_versionControl?itemPath=%24%2FGEOSage%2FGEOSage.bas)

### March 24, 2020
> [Officially released to the public as a standalone project.](https://github.com/happybono/GEOSage)

## Setup
1. Obtain a Google Maps API key from [Google Developers Console](https://developers.google.com/maps/documentation/geocoding/get-api-key).
2. Replace the placeholder `[Your Google Maps API Key]` in the code with your actual API key.
3. For Enterprise Geocoder, also provide your client ID by replacing the placeholder `[Your Google Maps ClientID]`.

## API Constants
- **gstrGeocodingDomain** : The domain for the Google Geocoding API.
- **gstrGeocodingURL** : The endpoint for geocoding requests.
- **gintType** : The type of API plan to use. Set to 0 for free, 1 for Enterprise, and 2 for Premium.
- **gstrKey** : Your Google Maps API key.
- **gstrClientID** : Your Google Maps client ID for Enterprise Geocoder.

## Prerequisites
* **Enable [Developer] tab** in **Microsoft Excel**. <br><br><img src="https://github.com/happybono/GEOSage/blob/master/Resources/GEOSage-SBS-001.png" alt="Step By Step 001" width="1658"><br><br>
* Within the **[Visual Basic]** IDE, add **"Microsoft XML, v6.0"**, **"Microsoft ActiveX Data Objects 6.1 Library"** both as  **References**. Could be found within **[Tools]** - **[References]**. <br><br><img src="https://github.com/happybono/GEOSage/blob/master/Resources/GEOSage-SBS-002.png" alt="Step By Step 002" width="600">

<img src="https://github.com/happybono/GEOSage/blob/master/HAPPYBONO-DMS-I.png" alt="HAPPYBONO-DMS-I-MonitoredRecords"/>

## Functions
### 1. `ADDRGEOCODE(address As String) As String`
Converts an address to latitude and longitude.

**Parameters:**
- `address`: The address to be geocoded.

**Returns:**
- A string containing the latitude and longitude, separated by a comma, or an error message if not found.

### 2. `URLEncode(ByVal StringVal As String, Optional SpaceAsPlus As Boolean = False) As String`
Encodes a string for use in a URL.

**Parameters:**
- `StringVal`: The string to be encoded.
- `SpaceAsPlus`: Optional boolean to encode spaces as plus signs (`+`) instead of `%20`.

**Returns:**
- The URL-encoded string.

### 3. `REVSGEOCODE(lat As String, lng As String) As String`
Converts latitude and longitude to an address.

**Parameters:**
- `lat`: The latitude.
- `lng`: The longitude.

**Returns:**
- The address corresponding to the given latitude and longitude, or an error message if not found.

### 4. `Base64_HMACSHA1(ByVal strTextToHash As String, ByVal strSharedSecretKey As String)`
Generates a Base64-encoded HMAC-SHA1 hash.

**Parameters:**
- `strTextToHash`: The text to be hashed.
- `strSharedSecretKey`: The shared secret key for hashing.

**Returns:**
- The Base64-encoded HMAC-SHA1 hash.

### 5. `Base64Decode(ByVal strData As String) As Byte()`
Decodes a Base64-encoded string to a byte array.

**Parameters:**
- `strData`: The Base64-encoded string.

**Returns:**
- The decoded byte array.

### 6. `Base64Encode(ByRef arrData() As Byte) As String`
Encodes a byte array to a Base64-encoded string.

**Parameters:**
- `arrData`: The byte array to be encoded.

**Returns:**
- The Base64-encoded string.

## Dependencies
- Microsoft XML, v6.0 (MSXML2.DOMDocument60, MSXML2.XMLHTTP60)
- ADODB.Stream for UTF-8 encoding

## Usage
### =ADDRGEOCODE([Address]) <br>
>Takes in the address of the location we want to geocode and returns the first latitude, longitude pair from GEOSage.

### =REVSGEOCODE([Latitude], [Longitude]) <br>
>Takes in a latitude, longitude pair and returns the first formatted address from GEOSage.

```vba
Sub ExampleUsage()
    Dim address As String
    Dim latlng As String
    Dim lat As String
    Dim lng As String
    Dim addressFromLatLng As String

    address = "B92 GARAGE, REDMOND, WA"
    latlng = ADDRGEOCODE(address)
    Debug.Print "Geocoded Address: " & latlng

    lat = "37.4219999"
    lng = "-122.0840575"
    addressFromLatLng = REVSGEOCODE(lat, lng)
    Debug.Print "Reverse Geocoded Address: " & addressFromLatLng
End Sub
```

## TO-DO List
* Functionality for [ArcGIS](https://www.arcgis.com/), [Bing Maps](https://www.bing.com/maps/), [Data Science Toolkit](http://www.datasciencetoolkit.org/) etc.

## Copyright 
Copyright ⓒ HappyBono 2020 - 2025. All rights Reserved.

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.
