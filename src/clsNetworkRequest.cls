VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNetworkRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' various information about network requests, see Network.requestWillBeSent and Network.responseReceived events
Option Explicit


Public Enum NetworkRequestState
    NRS_Pending
    NRS_Complete
End Enum

' state of request, see NetworkRequestState enum
Public state As NetworkRequestState

' Request identifier.
Public requestId As String ' requestId

' Loader identifier. Empty string if the request is fetched from worker.
Public loaderId As String ' loaderId

' URL of the document this request is loaded for.
Public documentURL As String

' Request data.
Public request As String ' request

' Timestamp.
Public timestampStart As Double ' MonotonicTime
Public timestampStop As Double ' MonotonicTime

' Timestamp. - ' see https://groups.google.com/g/chrome-debugging-protocol/c/X-Ts2-KygYA
Public wallTime As Double ' TimeSinceEpoch of (seconds since January 1, 1970)

' calculated as wallTime - timestamp, should be tracked per request
Public monotonicTimeOffset As Double

' Timestamp. VBA time event received
Public vbaTimestampStart As Date
Public vbaTimestampStop As Date

' Request initiator.
Public initiator As String ' initiator

' In the case that redirectResponse is populated, this flag indicates whether requestWillBeSentExtraInfo and responseReceivedExtraInfo events will be or were emitted for the request which was just redirected. EXPERIMENTAL
Public hasExtraInfo As Boolean

' [Redirect] response data.
Public response As String ' response

' Type of this resource.
Public resourceType As String ' ResourceType

' Frame identifier.
Public frameId As String ' Page.frameId

' Whether the request is initiated by a user gesture. Defaults to false.
Public hasUserGesture As Boolean

