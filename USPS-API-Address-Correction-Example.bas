Attribute VB_Name = "AddressCorrection"
Option Compare Database
Option Explicit


' Replace these with your USPS credentials
Const CLIENT_ID As String = "[your consumerKey]"  ' ~48-character string
Const CLIENT_SECRET As String = "[your ConsumerSecret]"    ' ~64 character string

' Get these values from USPS developer site
' https://developers.usps.com/
' Create an account. Add an app. Client credentials will be generated for you.



Function GetUSPSAccessToken( _
        Optional clientID As String = CLIENT_ID, _
        Optional clientSecret As String = CLIENT_SECRET) As String

' Generates an oAuth2 token. Valid for up to 8 hours. (as of 2025)
' Best practice is to reuse tokens as much as possible during a session.

' Note
' CLIENT_ID     is the USPS ConsumerKey (~48 characters)
' CLIENT_SECRET is the USPS ConsumerSecret: (~64 characters)


    Dim http As Object
    Dim url As String
    Dim payload As String
    Dim responseText As String
    Dim json As Object

    url = "https://apis.usps.com/oauth2/v3/token"      'Live system
'    url = "https://apis-tem.usps.com/oauth2/v3/token"    'For Testing - same ConsumerKey and ConsumerSecret as live system
    payload = "grant_type=client_credentials&client_id=" & clientID & _
                    "&client_secret=" & clientSecret & _
                   "&scope=addresses"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send payload
    responseText = http.responseText
'    Debug.Print ("Token Reply:" & responseText)
    Set json = JsonConverter.ParseJson(responseText)
    GetUSPSAccessToken = json("access_token")
End Function



Function VerifyUSPSAddress(sAccessToken As String, _
                        sAddress1 As String, _
                        sState As String, _
                        Optional sAddress2 As String = "", _
                        Optional sCity As String = "", _
                        Optional sZip5 As String = "", _
                        Optional sZipPlus4 As String = "") As String
    Dim objHTTP As Object
    Dim sURL As String
    Dim sResponse As String
    
    If sAccessToken = "" Then
        VerifyUSPSAddress = "{""Error"": ""Error - No access token generated.""}"
        Exit Function ' Exit if no Access Token is retrieved
    End If
    
'       sURL = "https://apis-tem.usps.com/addresses/v3/address?" & _  ' for testing
        sURL = "https://apis.usps.com/addresses/v3/address?" & _
              "streetAddress=" & URLEncode(sAddress1) & _
              IIf(Nz(sAddress2) <> "", "&secondaryAddress=" & URLEncode(sAddress2), "") & _
              IIf(Nz(sCity) <> "", "&city=" & URLEncode(sCity), "") & _
              "&state=" & URLEncode(sState) & _
              IIf(Nz(sZip5) <> "", "&ZIPCode=" & URLEncode(sZip5), "") & _
              IIf(Nz(sZipPlus4) <> "", "&ZIPPlus4=" & URLEncode(sZipPlus4), "")


    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
                          ' or("WinHttp.WinHttpRequest.5.1")
    
    With objHTTP
        .Open "GET", sURL, False
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & sAccessToken
'        .SetRequestHeader "Content-Type", "application/json"
        .Send
        sResponse = .responseText
    End With
    Set objHTTP = Nothing
    
'    Debug.Print sResponse ' Display the JSON response
    
    VerifyUSPSAddress = sResponse

End Function


Function URLEncode(str As String) As String
' Convert spaces and other "illegal" URL characters to %hex codes
    Dim i As Long
    Dim ch As String
    Dim encoded As String
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        Select Case Asc(ch)
            Case 48 To 57, 65 To 90, 97 To 122
                encoded = encoded & ch
            Case Else
                encoded = encoded & "%" & Hex(Asc(ch))
        End Select
    Next i

    URLEncode = encoded
End Function

Sub TestUSPSAddressVerification()

    Dim sReplyJSON As String
    Dim sAccessToken As String
    Dim json As Object
    Dim sCity
    Dim sState
    Dim sZip
    
    sAccessToken = GetUSPSAccessToken()
    If sAccessToken = "" Then
        Debug.Print "{""Error"": ""Error - No access token generated.""}"
        Exit Sub ' Exit if no Access Token is retrieved
    End If

    sReplyJSON = VerifyUSPSAddress(sAccessToken, "111 E Monroe St", "IL", "", "Springfield", "", "")
    Debug.Print sReplyJSON
    
'   Parse sReplyJSON
    Set json = JsonConverter.ParseJson(sReplyJSON)

'   JsonConverter requires VBA-JSON from:
'   https://github.com/VBA-tools/VBA-JSON
'   VBA-JSON requires including a reference to "Microsoft Scripting Runtime"

    sCity = json("address")("city")
    sState = json("address")("state")
    sZip = json("address")("ZIPCode")
    
    Debug.Print ("The ZIP code is: " & sZip)
    

' SAMPLE OF SUCCESSFUL REPLY JSON AND THE EXTRACTED ZIP CODE
'
'        {
'          "firm": "",
'          "address": {
'            "streetAddress": "111 E MONROE ST",
'            "streetAddressAbbreviation": "111 E MONROE ST",
'            "secondaryAddress": "",
'            "cityAbbreviation": "SPRINGFIELD",
'            "city": "SPRINGFIELD",
'            "state": "IL",
'            "ZIPCode": "62701",
'            "ZIPPlus4": "1103",
'            "urbanization": ""
'          },
'          "additionalInfo": {
'            "deliveryPoint": "11",
'            "carrierRoute": "C003",
'            "DPVConfirmation": "Y",
'            "DPVCMRA": "N",
'            "business": "Y",
'            "centralDeliveryPoint": "N",
'            "vacant": "N"
'          },
'          "corrections": [
'            {
'              "code": "",
'              "text": ""
'            }
'          ],
'          "matches": [
'            {
'              "code": "31",
'              "text": "Single Response - exact match"
'            }
'          ]
'        }
'
'The ZIP code is: 62701

End Sub

