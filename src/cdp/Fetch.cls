VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpFetch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpFetch"
' Fetch
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique request identifier.
' RequestId AS string

' Stages of the request to handle. Request will intercept before the request is
'   sent. Response will intercept after the response is received (but before response
'   body is received).
' RequestStage AS string
Public Enum RequestStage
    RS_Request
    RS_Response
End Enum

' 
' RequestPattern AS object

' Response HTTP header entry
' HeaderEntry AS object

' Authorization challenge for HTTP status code 401 or 407.
' AuthChallenge AS object

' Response to an AuthChallenge.
' AuthChallengeResponse AS object


Public Function RequestStageToString(ByVal rs As RequestStage) As String
    Dim retVal As String
    Select Case rs
        Case RS_Request
            retVal = "Request"
        Case RS_Response
            retVal = "Response"
        Case Else
            Debug.Print "Warning, unknown value " & rs
    End Select
    RequestStageToString = retVal
End Function

Public Function StringToRequestStage(ByVal s As String) As RequestStage
    Dim retVal As RequestStage
    Select Case s
        Case "Request"
            retVal = RS_Request
        Case "Response"
            retVal = RS_Response
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToRequestStage = retVal
End Function


Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Disables the fetch domain.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Fetch.disable", params
End Sub

' Enables issuing of requestPaused events. A request will be paused until client
' calls one of failRequest, fulfillRequest or continueRequest/continueWithAuth.
Public Sub enable( _
    Optional ByVal patterns AS Variant, _
    Optional ByVal handleAuthRequests AS Variant _
)
    ' patterns: array(optional) If specified, only requests matching any of these patterns will produce
    '   fetchRequested event and will be paused until clients response. If not set,
    '   all requests will be affected.
    ' handleAuthRequests: boolean(optional) If true, authRequired events will be issued and requests will be paused
    '   expecting a call to continueWithAuth.

    Dim params As New Dictionary
    If Not IsMissing(patterns) Then Set params("patterns") = patterns
    If Not IsMissing(handleAuthRequests) Then params("handleAuthRequests") = CBool(handleAuthRequests)

    cdp.InvokeMethod "Fetch.enable", params
End Sub

' Causes the request to fail with specified reason.
Public Sub failRequest( _
    ByVal requestId AS string, _
    ByVal errorReason AS string _
)
    ' requestId: RequestId An id the client received in requestPaused event.
    ' errorReason: Network.ErrorReason Causes the request to fail with the given reason.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)
    params("errorReason") = CStr(errorReason)

    cdp.InvokeMethod "Fetch.failRequest", params
End Sub

' Provides response to the request.
Public Sub fulfillRequest( _
    ByVal requestId AS string, _
    ByVal responseCode AS Long, _
    Optional ByVal responseHeaders AS Variant, _
    Optional ByVal binaryResponseHeaders AS Variant, _
    Optional ByVal body AS Variant, _
    Optional ByVal responsePhrase AS Variant _
)
    ' requestId: RequestId An id the client received in requestPaused event.
    ' responseCode: integer An HTTP response code.
    ' responseHeaders: array(optional) Response headers.
    ' binaryResponseHeaders: binary(optional) Alternative way of specifying response headers as a \0-separated
    '   series of name: value pairs. Prefer the above method unless you
    '   need to represent some non-UTF8 values that can't be transmitted
    '   over the protocol as text.
    ' body: binary(optional) A response body. If absent, original response body will be used if
    '   the request is intercepted at the response stage and empty body
    '   will be used if the request is intercepted at the request stage.
    ' responsePhrase: string(optional) A textual representation of responseCode.
    '   If absent, a standard phrase matching responseCode is used.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)
    params("responseCode") = CLng(responseCode)
    If Not IsMissing(responseHeaders) Then Set params("responseHeaders") = responseHeaders
    If Not IsMissing(binaryResponseHeaders) Then params("binaryResponseHeaders") = CStr(binaryResponseHeaders)
    If Not IsMissing(body) Then params("body") = CStr(body)
    If Not IsMissing(responsePhrase) Then params("responsePhrase") = CStr(responsePhrase)

    cdp.InvokeMethod "Fetch.fulfillRequest", params
End Sub

' Continues the request, optionally modifying some of its parameters.
Public Sub continueRequest( _
    ByVal requestId AS string, _
    Optional ByVal url AS Variant, _
    Optional ByVal method AS Variant, _
    Optional ByVal postData AS Variant, _
    Optional ByVal headers AS Variant, _
    Optional ByVal interceptResponse AS Variant _
)
    ' requestId: RequestId An id the client received in requestPaused event.
    ' url: string(optional) If set, the request url will be modified in a way that's not observable by page.
    ' method: string(optional) If set, the request method is overridden.
    ' postData: binary(optional) If set, overrides the post data in the request.
    ' headers: array(optional) If set, overrides the request headers.
    ' interceptResponse: boolean(optional) If set, overrides response interception behavior for this request.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)
    If Not IsMissing(url) Then params("url") = CStr(url)
    If Not IsMissing(method) Then params("method") = CStr(method)
    If Not IsMissing(postData) Then params("postData") = CStr(postData)
    If Not IsMissing(headers) Then Set params("headers") = headers
    If Not IsMissing(interceptResponse) Then params("interceptResponse") = CBool(interceptResponse)

    cdp.InvokeMethod "Fetch.continueRequest", params
End Sub

' Continues a request supplying authChallengeResponse following authRequired event.
Public Sub continueWithAuth( _
    ByVal requestId AS string, _
    ByVal authChallengeResponse AS Dictionary _
)
    ' requestId: RequestId An id the client received in authRequired event.
    ' authChallengeResponse: AuthChallengeResponse Response to  with an authChallenge.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)
    Set params("authChallengeResponse") = authChallengeResponse

    cdp.InvokeMethod "Fetch.continueWithAuth", params
End Sub

' Continues loading of the paused response, optionally modifying the
' response headers. If either responseCode or headers are modified, all of them
' must be present.
' Experimental
Public Sub continueResponse( _
    ByVal requestId AS string, _
    Optional ByVal responseCode AS Variant, _
    Optional ByVal responsePhrase AS Variant, _
    Optional ByVal responseHeaders AS Variant, _
    Optional ByVal binaryResponseHeaders AS Variant _
)
    ' requestId: RequestId An id the client received in requestPaused event.
    ' responseCode: integer(optional) An HTTP response code. If absent, original response code will be used.
    ' responsePhrase: string(optional) A textual representation of responseCode.
    '   If absent, a standard phrase matching responseCode is used.
    ' responseHeaders: array(optional) Response headers. If absent, original response headers will be used.
    ' binaryResponseHeaders: binary(optional) Alternative way of specifying response headers as a \0-separated
    '   series of name: value pairs. Prefer the above method unless you
    '   need to represent some non-UTF8 values that can't be transmitted
    '   over the protocol as text.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)
    If Not IsMissing(responseCode) Then params("responseCode") = CLng(responseCode)
    If Not IsMissing(responsePhrase) Then params("responsePhrase") = CStr(responsePhrase)
    If Not IsMissing(responseHeaders) Then Set params("responseHeaders") = responseHeaders
    If Not IsMissing(binaryResponseHeaders) Then params("binaryResponseHeaders") = CStr(binaryResponseHeaders)

    cdp.InvokeMethod "Fetch.continueResponse", params
End Sub

' Causes the body of the response to be received from the server and
' returned as a single string. May only be issued for a request that
' is paused in the Response stage and is mutually exclusive with
' takeResponseBodyForInterceptionAsStream. Calling other methods that
' affect the request or disabling fetch domain before body is received
' results in an undefined behavior.
' Returns:  - 
Public Function getResponseBody( _
    ByVal requestId AS string _
) AS Dictionary
    ' requestId: RequestId Identifier for the intercepted request to get body for.

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Fetch.getResponseBody", params)
    If cdp.ErrorCode = 0 Then
        Set getResponseBody = results
    End If
End Function

' Returns a handle to the stream representing the response body.
' The request must be paused in the HeadersReceived stage.
' Note that after this command the request can't be continued
' as is -- client either needs to cancel it or to provide the
' response body.
' The stream only supports sequential read, IO.read will fail if the position
' is specified.
' This method is mutually exclusive with getResponseBody.
' Calling other methods that affect the request or disabling fetch
' domain before body is received results in an undefined behavior.
' Returns: stream - 
Public Function takeResponseBodyAsStream( _
    ByVal requestId AS string _
) AS string
    ' requestId: RequestId 

    Dim params As New Dictionary
    params("requestId") = CStr(requestId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Fetch.takeResponseBodyAsStream", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("stream") Then takeResponseBodyAsStream = results("stream")
    End If
End Function

