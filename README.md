# GEOSage <br> <img src="https://github.com/happybono/GEOSage/blob/master/Resources/powered_by_msexcel_on_white.png" alt="Powered by MSExcel logo" width="217"/>

A VBA application for geocoding and reverse geocoding in Excel. Supports both Google's free and enterprise for business geocoder (Google Maps APIs for Business, Google Maps for Work or Google Maps).

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

## Prerequisites
* **Enable [Developer] tab** in **Microsoft Excel**.
* Within the **[Visual Basic IDE]**, add **"Microsoft XML, v6.0"**, **"Microsoft ActiveX Data Objects 6.1 Library"** both as  **References**. Could be found within **[Tools]** - **[References]**.

## Usage
#### =ADDRGEOCODE([Address]) <br>
>Takes in the address of the location we want to geocode and returns the first latitude, longitude pair from GEOSage.

#### =REVSGEOCODE([Latitude], [Longitude]) <br>
>Takes in a latitude, longitude pair and returns the first formatted address from GEOSage.

## TO-DO List
* Functionality for Bing Maps, Data Science Toolkit, ArcGIS etc.

## Copyright / End User License
### Copyright
Copyright ⓒ HappyBono 2020. All rights Reserved.

### MIT License
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Contact Information
[Jaewoong Mun](mailto:happybono@outlook.com)
