''
''    FILE: GEOSage.bas
''  AUTHOR: Jaewoong Mun (happybono@outlook.com)
'' CREATED: February 05, 2020
''
'' Released to the public domain
''

Attribute VB_Name = "GEOSage"

Option Explicit

' domain and URL for Google Geocoding API
Public Const gstrGeocodingDomain = "https://maps.googleapis.com"
Public Const gstrGeocodingURL = "/maps/api/geocode/xml?"

' set gintType = 1 to use the Enterprise Geocoder (requires clientID and Google Maps Geocoding API Key)
' set gintType = 2 to use the API Premium Plan (requires Google Maps Geocoding API Key)
' leave gintType = 0 to use the free-ish Google geocoder (requires Google Maps Geocoding API Key!
' see https://developers.google.com/maps/documentation/geocoding/get-api-key)
Public Const gintType = 0

' key for Enterprise Geocoder or API Premium Plan or free-ish geocoder
Public Const gstrKey = "AIzaSyCnjOmTA2_RdfxloaWBwsHCvw6D0_aBqjo"

' clientID for Enterprise Geocoder (if applicable)
Public Const gstrClientID = "[Your Google Maps ClientID]"

' kludge to not overdo the API calls and add a delay
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Public Function ADDRGEOCODE(address As String) As String
    Dim strAddress As String
    Dim strQuery As String
    Dim strLatitude As String
    Dim strLongitude As String
    Dim strQueryBland As String

    strAddress = URLEncode(address)

    'assemble the query string
    strQuery = gstrGeocodingURL
    strQuery = strQuery & "address=" & strAddress

    If gintType = 0 Then ' free-ish Google Geocoder - required an API key!
        strQuery = strQuery & "&key=" & gstrKey
    ElseIf gintType = 1 Then ' Enterprise Geocoder
        strQuery = strQuery & "&client=" & gstrClientID
        strQuery = strQuery & "&signature=" & Base64_HMACSHA1(strQuery, gstrKey)
    ElseIf gintType = 2 Then ' API Premium Plan
        strQuery = strQuery & "&key=" & gstrKey
    End If

    'define XML and HTTP components
    Dim googleResult As New MSXML2.DOMDocument60
    Dim googleService As New MSXML2.XMLHTTP60
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode

    Sleep (5)

    'make sure to have create HTTP request to query URL
    googleService.Open "GET", gstrGeocodingDomain & strQuery, False
    googleService.send
    googleResult.LoadXML (googleService.responseText)

    Set oNodes = googleResult.getElementsByTagName("geometry")

    If oNodes.Length = 1 Then
        For Each oNode In oNodes
            Debug.Print oNode.Text
            strLatitude = oNode.ChildNodes(0).ChildNodes(0).Text
            strLongitude = oNode.ChildNodes(0).ChildNodes(1).Text
            ADDRGEOCODE = strLatitude & "," & " " & strLongitude
        Next oNode
    Else
        ADDRGEOCODE = "Not Found (You may have reached your daily limit. Please check your daily quota and try again.)"
    End If
End Function


' Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
'   Dim StringLen As Long: StringLen = Len(StringVal)
'
'   If StringLen > 0 Then
'     ReDim result(StringLen) As String
'     Dim i As Long, CharCode As Integer
'     Dim Char As String, Space As String
'
'     If SpaceAsPlus Then Space = "+" Else Space = "%20"
'
'     For i = 1 To StringLen
'       Char = Mid$(StringVal, i, 1)
'       CharCode = asc(Char)
'       Select Case CharCode
'         Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
'           result(i) = Char
'         Case 32
'           result(i) = Space
'         Case 0 To 15
'           result(i) = "%0" & Hex(CharCode)
'         Case Else
'           result(i) = "%" & Hex(CharCode)
'       End Select
'     Next i
'     URLEncode = Join(result, "")
'   End If
' End Function


Public Function URLEncode(ByVal StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
  Dim bytes() As Byte, b As Byte, i As Integer, space As String

  If SpaceAsPlus Then space = "+" Else space = "%20"

  If Len(StringVal) > 0 Then
    With New ADODB.Stream
      .Mode = adModeReadWrite
      .Type = adTypeText
      .Charset = "UTF-8"
      .Open
      .WriteText StringVal
      .Position = 0
      .Type = adTypeBinary
      .Position = 3 ' skip BOM
      bytes = .Read
    End With

    ReDim result(UBound(bytes)) As String

    For i = UBound(bytes) To 0 Step -1
      b = bytes(i)
      Select Case b
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Chr(b)
        Case 32
          result(i) = space
        Case 0 To 15
          result(i) = "%0" & Hex(b)
        Case Else
          result(i) = "%" & Hex(b)
      End Select
    Next i

    URLEncode = Join(result, "")
  End If
End Function


Public Function REVSGEOCODE(lat As String, lng As String) As String
    Dim strAddress As String
    Dim strLat As String
    Dim strLng As String
    Dim strQuery As String
    Dim strLatitude As String
    Dim strLongitude As String

    strLat = URLEncode(lat)
    strLng = URLEncode(lng)

    'assembles the query string
    strQuery = gstrGeocodingURL
    strQuery = strQuery & "latlng=" & strLat & "," & strLng
    
    If gintType = 0 Then ' free-ish Google Geocoder - required an API key!
        strQuery = strQuery & "&key=" & gstrKey
    ElseIf gintType = 1 Then ' Enterprise Geocoder
        strQuery = strQuery & "&client=" & gstrClientID
        strQuery = strQuery & "&signature=" & Base64_HMACSHA1(strQuery, gstrKey)
    ElseIf gintType = 2 Then ' API Premium Plan
        strQuery = strQuery & "&key=" & gstrKey
    End If

    'define XML and HTTP components
    Dim googleResult As New MSXML2.DOMDocument60
    Dim googleService As New MSXML2.XMLHTTP60
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode

    Sleep (5)

    'create HTTP request to query URL - make sure to have
    googleService.Open "GET", gstrGeocodingDomain & strQuery, False
    googleService.send
    googleResult.LoadXML (googleService.responseText)

    Set oNodes = googleResult.getElementsByTagName("formatted_address")
  
    If oNodes.Length > 0 Then
        REVSGEOCODE = oNodes.Item(0).Text
    Else
        REVSGEOCODE = "Not Found (You may have reached your daily limit. Please check your daily quota and try again.)"
    End If
End Function


Public Function Base64_HMACSHA1(ByVal strTextToHash As String, ByVal strSharedSecretKey As String)
    Dim asc As Object
    Dim enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")
    
    strSharedSecretKey = Replace(Replace(strSharedSecretKey, "-", "+"), "_", "/")
    SharedSecretKey = Base64Decode(strSharedSecretKey)
    enc.Key = SharedSecretKey
    
    TextToHash = asc.Getbytes_4(strTextToHash)
    bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = Replace(Replace(Base64Encode(bytes), "+", "-"), "/", "_")
End Function


Public Function Base64Decode(ByVal strData As String) As Byte()
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement

    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    Base64Decode = objNode.nodeTypedValue

    Set objNode = Nothing
    Set objXML = Nothing
End Function


Public Function Base64Encode(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement

    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    Base64Encode = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing
End Function
