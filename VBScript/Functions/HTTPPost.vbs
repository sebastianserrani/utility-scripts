'===============================================================================
' Name: HTTPPost
' Input:
'    ByVal strUrl String - URL to POST
'    ByVal strType String - The content type
'    ByVal strRequest String - Payload
'    ByVal booReturnResponse Bool - Return or echo the response
' Output:
'    HTTPPost String - HTTP response
' Purpose: Allows to POST to an URL endpoint
' Remarks: setOption options:
'          SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2 : Ignore certificate errors
'          SXH_SERVER_CERT_IGNORE_UNKNOWN_CA = 256 : Unknown certificate authority
'          SXH_SERVER_CERT_IGNORE_WRONG_USAGE = 512 : Malformed certificate
'          SXH_SERVER_CERT_IGNORE_CERT_CN_INVALID = 4096 : Mismatch between hostname and certificate
'          SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID = 8192 : The date in the certificate is invalid
'          SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056 : All certificate errors
' Author: Sebastian Serrani
' Usage:
'     WScript.Echo HTTPPost("https://reqres.in/api/users/", "json", "{""name"": ""morpheus"", ""job"": ""leader""}", True)
'     HTTPPost("http://web.url/endpoint/", "x-www-form-urlencoded", "message=" & strText, False)
'===============================================================================

Function HTTPPost(strUrl, strType, strRequest, booReturnResponse)
	Err.Clear
	Dim objHTTP : Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	Dim strHTTPStatus: strHTTPStatus = "---"
	objHTTP.Open "POST", strUrl, False
	objHTTP.setRequestHeader "User-Agent", "HTTPPost"
	objHTTP.SetRequestHeader "Content-Type", "application/" & strType
	objHTTP.SetRequestHeader "Content-Length", Len(strRequest)
	objHTTP.SetOption 2, 13056
	On Error Resume Next
	objHTTP.Send strRequest
	If booReturnResponse = True Then
		HTTPPost = objHTTP.responseText
	Else
		WScript.Echo "[HTTP:" & objHTTP.Status & "][" & objHTTP.StatusText & "] " & objHTTP.responseText
	End If
	Err.Clear
End Function
