VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSessionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' maintains all state associated with an attached target, i.e. session specific information for automating via Chrome Developer Protocol
Option Explicit

' how many seconds since last network request completed received to consider for not pending requests
Private Const SecondsSinceLastNetworkRequest As Integer = -2

' session identifier - which session is this, as returned from attach call
Public sessionId As String

' target information session applies to
Public targetInfo As cdpTargetTargetInfo

' flags set as various events received
Public eventFrameStoppedLoading As Boolean
Public eventDomContentEventFired As Boolean

' cached hWnd of browser window
Public browserHWnd As Long

' cached root node as returned by DOM.GetDocument
Public rootNode As cdpDOMNode

' active Frames in current page, includes IFrames
' keyed on frame id of cdpPageFrame objects
' Note: new frames added only on frameNavigated event; only most recent frame with matching id kept
Public frames As Dictionary

' list of active named frames, maps name to id
' Note: if duplicate frames with same name then only latest is kept
Public frameNames As Dictionary  ' of String

' list of active execution context id to frame id mappings
' keyed on both frame id and execution context id
' Note: updated as executionContextCreatead and executionContextDestroyed occur
Public frameExecCtxMapping As Dictionary  ' of String

' frames currently loading, keyed on frame id; empty if no frames still loading
Public framesLoading As Dictionary

' network requests pending completion, keyed on request id; empty if no pending requests
' value is a clsNetworkRequest object
Public pendingNetworkRequest As Dictionary

' has there been a window open event since last wait, click, navigate, or session initially created
Public windowOpenEventOccurred As Boolean

' if a new window is being opened, the url associated with it (so can match to
Public childWindowUrl As String

' returns number of pending network requests after removing stale requests
' a stale request is any request completed before last 500ms
Public Function pendingNetworkRequests() As Integer
    Dim requestInfo As clsNetworkRequest, v As Variant
    Dim epoch As Date: epoch = DateAdd("s", SecondsSinceLastNetworkRequest, Now)
    For Each v In pendingNetworkRequest.Keys
        Set requestInfo = pendingNetworkRequest.Item(v)
        If requestInfo.state = NetworkRequestState.NRS_Complete Then
            If requestInfo.vbaTimestampStop < epoch Then
                pendingNetworkRequest.Remove v
            End If
        End If
    Next v
    pendingNetworkRequests = pendingNetworkRequest.count
End Function


Private Sub Class_Initialize()
    ' state variables
    Set targetInfo = New cdpTargetTargetInfo
    Set frames = New Dictionary
    Set frameNames = New Dictionary
    Set frameExecCtxMapping = New Dictionary
    Set framesLoading = New Dictionary
    Set pendingNetworkRequest = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set targetInfo = Nothing
    Set frames = Nothing
    Set frameNames = Nothing
    Set frameExecCtxMapping = Nothing
    Set framesLoading = Nothing
    Set pendingNetworkRequest = Nothing
End Sub
