VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWebSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' see https://github.com/EagleAglow/vba-websocket-class
' MIT licensed
Option Explicit

Private pState As Long         ' connection state
Private pSCC As Boolean        ' flag for "server closed connection"
Private pServer As String      ' host name
Private pPort As Long          ' internet port, usually 80 or 443
Private pPath As String        ' at least "/"
Private pProtocol As String    ' e.g. WASM or echo-protocol
Private pErrorText As String   ' set to text explaining error
Private pAgentHeader As String ' CWebSocket

Private pOutgoingMessageUTF8 As String
Private pOutgoingMessageBinary() As Byte
Private pLastReceivedMessageUTF8 As String
Private pLastReceivedMessageBinary() As Byte

Private hSessionHandle As LongPtr
Private hConnectionHandle As LongPtr
Private hRequestHandle As LongPtr
Private hWebSocketHandle As LongPtr
Private fStatus As LongPtr

' State Property
Public Property Get state() As String
Select Case pState
  Case 1
    state = "Not Connected"
  Case 2
    state = "Connecting"
  Case 3
    state = "Connected"
  Case 4
    state = "Disconnecting"
  Case Else
    state = "Error"
End Select
End Property

' Server Property
Public Property Get Server() As String
Server = pServer
End Property
Public Property Let Server(value As String)
If pState = 1 Then ' only set value while not connected
  pServer = value
End If
End Property

' Port Property
Public Property Get port() As Long
port = pPort
End Property
Public Property Let port(value As Long)
If pState = 1 Then ' only set value while not connected
  pPort = value
End If
End Property

' Path Property
Public Property Get path() As String
path = pPath
End Property
Public Property Let path(value As String)
If pState = 1 Then ' only set value while not connected
  pPath = value
End If
End Property

' Protocol Property
Public Property Get Protocol() As String
Protocol = pProtocol
End Property
Public Property Let Protocol(value As String)
If pState = 1 Then ' only set value while not connected
  pProtocol = value
End If
End Property

' AgentHeader Property
Public Property Get AgentHeader() As String
Protocol = pAgentHeader
End Property
Public Property Let AgentHeader(value As String)
If pState = 1 Then ' only set value while not connected
  pAgentHeader = value
End If
End Property

' ErrorText Property
Public Property Get ErrorText() As String
ErrorText = pErrorText
End Property

' WebSocketHandle Property
Public Property Get WebSocketHandle() As String
WebSocketHandle = hWebSocketHandle
End Property

' RequestHandle Property
Public Property Get RequestHandle() As String
RequestHandle = hRequestHandle
End Property

' SessionHandle Property
Public Property Get SessionHandle() As String
SessionHandle = hSessionHandle
End Property

' ConnectionHandle Property
Public Property Get ConnectionHandle() As String
ConnectionHandle = hConnectionHandle
End Property

' pSCC flag Property
Public Property Get SCC() As Boolean
SCC = pSCC
End Property

' connects to HTTP server an issues GET request, returns results
Public Function HttpGetMessage(ByVal pServer As String, ByVal port As Long, ByVal path As String) As String
    On Error GoTo ErrHandler
    Dim hSessionHandle As LongPtr
    Dim hConnectionHandle As LongPtr
    Dim hRequestHandle As LongPtr
    pErrorText = "None"

    ' Create session handle
    hSessionHandle = WinHttpOpen(StrPtr(AgentHeader), WINHTTPxACCESSxTYPExDEFAULTxPROXY, 0, 0, WINHTTPxFLAGxSYNC)
    If hSessionHandle = 0 Then
        pErrorText = "Could not create WinHttp session handle"
        Exit Function
    End If

    ' Create connection handle
    hConnectionHandle = WinHttpConnect(hSessionHandle, StrPtr(pServer), port, 0)
    If hConnectionHandle = 0 Then
        pErrorText = "Failed to reach server:port at: " & pServer & ":" & port
        GoTo cleanupSession
    End If

    ' Create request handle - use 0 for null pointer to empty strings: Version, Referrer, AcceptTypes
    Dim method As String: method = "GET" ' always
    hRequestHandle = WinHttpOpenRequest(hConnectionHandle, StrPtr(method), StrPtr(path), 0, 0, 0, 0)
    If hRequestHandle = 0 Then
        pErrorText = "Request failed for path: " & path
        GoTo cleanupConnection
    End If

    ' actually send it
    If WinHttpSendRequest(hRequestHandle, 0, 0, 0, 0, 0, 0) = 0 Then
        pErrorText = "Sending request failed for path: " & path
        GoTo cleanupConnection
    End If

    ' wait for a response
    If WinHttpReceiveResponse(hRequestHandle, 0) = 0 Then
        pErrorText = "Error waiting for response"
        GoTo cleanupConnection
    End If
    
    ' and read the response
    Dim receiveBuffer() As Byte
    ReDim receiveBuffer(1023) As Byte
    Dim dwBufferLength As LongPtr
    dwBufferLength = (UBound(receiveBuffer) - LBound(receiveBuffer) + 1)
    ReDim receiveBuffer(4096) As Byte
    dwBufferLength = (UBound(receiveBuffer) - LBound(receiveBuffer) + 1)
    Dim dwBytesTransferred As LongPtr
    dwBytesTransferred = 0
    If WinHttpReadData(hRequestHandle, receiveBuffer(0), CLng(dwBufferLength), dwBytesTransferred) = 0 Then
        pErrorText = "Error reading data"
        GoTo cleanupConnection
    End If
    
    ' and return it
    ' resize buffer to dwBytesTransferred, then convert buffer into VBA string
    ReDim Preserve receiveBuffer(CInt(dwBytesTransferred))
    HttpGetMessage = Utf8BytesToString(receiveBuffer)

cleanupConnection:
    If (hConnectionHandle <> 0) Then
        WinHttpCloseHandle (hConnectionHandle)
        hConnectionHandle = 0
    End If

cleanupSession:
    If (hSessionHandle <> 0) Then
        WinHttpCloseHandle (hSessionHandle)
        hSessionHandle = 0
    End If
    
    Exit Function
ErrHandler:
    Debug.Print "HttpGetMessage() - Error: " & Err.description
    Stop
    Resume
End Function

Public Function Connect()
If pState <> 1 Then ' only run this while not connected
  pErrorText = "Must be disconnected before attempting to connect"
  Exit Function
End If
Connect = False ' unless all steps complete
pState = 2 ' connecting
' check to see if server is set
If Len(pServer) = 0 Then
  pErrorText = "Missing ServerName"
  pState = 1
  Exit Function
End If
' check to see if path begins with "/"
If left(pPath, 1) <> "/" Then
  pErrorText = "Inalid Path: " & pPath
  pState = 1
  Exit Function
End If

' Create session handle
hSessionHandle = WinHttpOpen(StrPtr(AgentHeader), _
      WINHTTPxACCESSxTYPExDEFAULTxPROXY, 0, 0, WINHTTPxFLAGxSYNC)
If hSessionHandle = 0 Then
  pErrorText = "Could not create WinHttp session handle"
  pState = 1
  Exit Function
End If

' Create connection handle
hConnectionHandle = WinHttpConnect(hSessionHandle, StrPtr(pServer), port, 0)
If hConnectionHandle = 0 Then
  pErrorText = "Failed to reach server:port at: " & pServer & ":" & port
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  pState = 1
  Exit Function
End If

' Create request handle - use 0 for null pointer to empty strings: Version, Referrer, AcceptTypes
Dim method As String
method = "GET" ' always
hRequestHandle = WinHttpOpenRequest(hConnectionHandle, StrPtr(method), StrPtr(pPath), 0, 0, 0, 0)
If hRequestHandle = 0 Then
  pErrorText = "Request failed for path: " & pPath
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  pState = 1
  Exit Function
End If

' Prepare to request client protocol upgrade from http to websocket, returns true if success
fStatus = WinHttpSetOption(hRequestHandle, WINHTTPxOPTIONxUPGRADExTOxWEBxSOCKET, 0, 0)
If (fStatus = 0) Then ' failed
  pErrorText = "Upgrade from http to websocket failed (Step 1/6)"
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  pState = 1
  Exit Function
End If

' Perform websocket handshake by sending the upgrade request to server
' --------------------------------------------------------------------
' Application may specify additional headers if needed.
' --------------------------------------------------------------------
' Not in original code...
' Each header except the last must be terminated by a carriage return/line feed (vbCrLf).
' Uses an odd API feature: passes string length as -1, API figures out length
' Note: This is where websocket (internal, RFC "subprotocol") protocol is set
' --------------------------------------------------------------------
Dim HeaderText As String
Dim HeaderTextLength As Long
HeaderText = ""
HeaderText = HeaderText & "Host: " & pServer & vbCrLf   ' may be redundant or unnecessary
HeaderText = HeaderText & "Sec-WebSocket-Version: 13" & vbCrLf  ' 8 or 13, may be redundant or unnecessary
HeaderText = HeaderText & "Sec-Websocket-Protocol: echo-protocol" & vbCrLf  ' subprotocol
' setup for API call, trim any trailing vbCrLf
If (Right(HeaderText, 2) = vbCrLf) Then
  HeaderText = left(HeaderText, Len(HeaderText) - 2)
End If

HeaderText = vbNullString
If Len(HeaderText) > 0 Then ' let the API figure it out
  HeaderTextLength = -1
  fStatus = WinHttpSendRequest(hRequestHandle, StrPtr(HeaderText), _
               HeaderTextLength, WINHTTPxNOxREQUESTxDATA, 0, 0, 0)
Else  ' call without adding headers
  fStatus = WinHttpSendRequest(hRequestHandle, WINHTTPxNOxADDITIONALxHEADERS, _
               0, WINHTTPxNOxREQUESTxDATA, 0, 0, 0)
End If
If (fStatus = 0) Then ' failed
  pErrorText = "Upgrade from http to websocket failed (Step 2/6)"
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  Exit Function
  pState = 1
End If

' Receive server reply
fStatus = WinHttpReceiveResponse(hRequestHandle, 0)
If (fStatus = 0) Then ' failed
  pErrorText = "Upgrade from http to websocket failed (Step 3/6)"
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  Exit Function
  pState = 1
End If

' See if the HTTP Response confirms the upgrade, with HTTP status code 101.
Dim dwStatusCode As Long
Dim sizeStatusCode As Long  ' for HTTP result request
sizeStatusCode = 4 ' four bytes for long ' for HTTP result request
fStatus = WinHttpQueryHeaders(hRequestHandle, _
    (WINHTTPxQUERYxSTATUSxCODE Or WINHTTPxQUERYxFLAGxNUMBER), _
    WINHTTPxHEADERxNAMExBYxINDEX, _
    dwStatusCode, sizeStatusCode, WINHTTPxNOxHEADERxINDEX)
If (fStatus = 0) Then ' failed
  pErrorText = "Upgrade from http to websocket failed (Step 4/6)"
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  Exit Function
  pState = 1
End If

If dwStatusCode <> 101 Then
  Debug.Print "Code needs to be 101, ending..."
  pErrorText = "Upgrade from http to websocket failed (Step 5/6)"
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  Exit Function
  pState = 1
End If

' finally, get handle to websocket
hWebSocketHandle = WinHttpWebSocketCompleteUpgrade(hRequestHandle, 0)
If hWebSocketHandle = 0 Then
  pErrorText = "Upgrade from http to websocket failed (Step 6/6)"
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  Exit Function
  pState = 1
End If


' The request handle is not needed anymore. From now on we will use the websocket handle.
WinHttpCloseHandle (hRequestHandle)
hRequestHandle = 0

' connected!
pState = 3
Connect = True ' all steps complete
Debug.Print "Succesfully upgraded to websocket protocol at: " & pServer & ":" & pPort & pPath

End Function






Public Function Disconnect()
If pState <> 3 Then ' only run this while connected
  pErrorText = "Must be connected before attempting to disconnect"
  Exit Function
End If
Disconnect = False ' unless all steps complete
pState = 4 ' disconnecting

' wrap up
Dim dwError As Long
dwError = WinHttpWebSocketClose(hWebSocketHandle, WINHTTPxWEBxSOCKETxSUCCESSxCLOSExSTATUS, 0, 0)
If (dwError <> ERRORxSUCCESS) Then
  pErrorText = "Websocket close failed"
  If (hWebSocketHandle <> 0) Then
    WinHttpCloseHandle (hWebSocketHandle)
    hWebSocketHandle = 0
  End If
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  Exit Function
  pState = 1
End If

' Check close status returned by the server.
Dim rgbCloseReasonBuffer(123) As Byte
Dim dwCloseReasonBufferLength As Long
Dim dwCloseReasonLength As Long
dwCloseReasonLength = 0
Dim usStatus As Integer
usStatus = 0
' get correct value in case of change to buffer size
dwCloseReasonBufferLength = (UBound(rgbCloseReasonBuffer) - LBound(rgbCloseReasonBuffer) + 1)
dwError = WinHttpWebSocketQueryCloseStatus(hWebSocketHandle, usStatus, _
             rgbCloseReasonBuffer(0), dwCloseReasonBufferLength, CLngPtr(dwCloseReasonLength))
If (dwError <> ERRORxSUCCESS) Then
  pErrorText = "Websocket close reason request failed"
  If (hWebSocketHandle <> 0) Then
    WinHttpCloseHandle (hWebSocketHandle)
    hWebSocketHandle = 0
  End If
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
  Exit Function
  pState = 1
End If

' report result
If usStatus = WINHTTPxWEBxSOCKETxSUCCESSxCLOSExSTATUS Then
  Debug.Print "Program ended correctly"
Else
  Debug.Print "The server closed the connection with status code: " & usStatus
  Dim strCloseReason As String
  Dim i As Long
  strCloseReason = ""
  For i = LBound(rgbCloseReasonBuffer) To UBound(rgbCloseReasonBuffer)
    If rgbCloseReasonBuffer(i) <> 0 Then
      strCloseReason = strCloseReason & Chr(rgbCloseReasonBuffer(i))
    End If
  Next
  If Len(strCloseReason) > 0 Then
    Debug.Print " and reason: " & strCloseReason
   End If
End If

' finally, just to be sure the handles are closed
If (hWebSocketHandle <> 0) Then
  WinHttpCloseHandle (hWebSocketHandle)
  hWebSocketHandle = 0
End If
If (hRequestHandle <> 0) Then
  WinHttpCloseHandle (hRequestHandle)
  hRequestHandle = 0
End If
If (hConnectionHandle <> 0) Then
  WinHttpCloseHandle (hConnectionHandle)
  hConnectionHandle = 0
End If
If (hSessionHandle <> 0) Then
  WinHttpCloseHandle (hSessionHandle)
  hSessionHandle = 0
End If

' disconnected!
pState = 1
Disconnect = True ' all steps complete
'Debug.Print "Succesfully disconnected"

End Function


Public Function SendMessageUTF8(SendThis As String) As Boolean
pErrorText = "None"
SendMessageUTF8 = False ' unless completed
If pState <> 3 Then
  pErrorText = "Must be connected to send"
  Exit Function
End If

Dim dwError As Long
Dim Utf8Message() As Byte
Utf8Message = Utf8BytesFromString(SendThis)
Dim cdwMessageLength As Long
cdwMessageLength = BytesLength(Utf8Message)

dwError = WinHttpWebSocketSend(hWebSocketHandle, _
             WINHTTPxWEBxSOCKETxUTF8xMESSAGExBUFFERxTYPE, _
             VarPtr(Utf8Message(0)), cdwMessageLength)
If (dwError <> ERRORxSUCCESS) Then
  pErrorText = "SendMessageUTF8 Error: " & dwError & "/" & GetLastError
End If
SendMessageUTF8 = True
End Function


Public Function GetMessageUTF8() As String
On Error GoTo ErrHandler
GetMessageUTF8 = vbNullString  ' default return
pErrorText = "None"
If pState <> 3 Then
  pErrorText = "Must be connected to receive"
  Exit Function
End If

Dim rgbBuffer() As Byte  ' do it this way, so we can redim it later
ReDim rgbBuffer(1023) As Byte
Dim dwBufferLength As LongPtr
' get correct value in case of change to buffer size
dwBufferLength = (UBound(rgbBuffer) - LBound(rgbBuffer) + 1)

Dim pbCurrentBufferPointer As LongPtr
pbCurrentBufferPointer = VarPtr(rgbBuffer(0))
Dim dwBytesTransferred As LongPtr
dwBytesTransferred = 0
Dim eBufferType As LongPtr
eBufferType = 0
Dim dwCloseReasonLength As Long
dwCloseReasonLength = 0
Dim usStatus As Integer
usStatus = 0

' if we get a message fragment packet, loop until we get a complete packet
Dim dwError As Long

Dim MessageComplete As Boolean
MessageComplete = False
Do Until (MessageComplete)
  If (dwBufferLength = 0) Then
    dwError = ERRORxNOTxENOUGHxMEMORY  ' note - not using underscores, can confuse class compilation
    pErrorText = "Need a bigger buffer!"
    Exit Function
  End If
  
  ' receive - will hang if no server response
  dwError = WinHttpWebSocketReceive(hWebSocketHandle, rgbBuffer(0), CLng(dwBufferLength), dwBytesTransferred, eBufferType)
  If (dwError <> ERRORxSUCCESS) Then
      pErrorText = "GetMessageUTF8 Error: " & dwError & "/" & GetLastError
      Exit Function
  End If
  
  ' check for close message
  If (eBufferType = WINHTTPxWEBxSOCKETxCLOSExBUFFERxTYPE) Then ' server wants to quit
    pSCC = True
    pErrorText = "Server Closed Connection"
    Exit Function
  End If
  
' If we receive just part of the message repeat the receive operation.
  pbCurrentBufferPointer = pbCurrentBufferPointer + dwBytesTransferred
  dwBufferLength = dwBufferLength - dwBytesTransferred
  If Not (eBufferType = WINHTTPxWEBxSOCKETxUTF8xFRAGMENTxBUFFERxTYPE) Then
    MessageComplete = True
  End If
Loop

' Expected server to just echo UTF8 message - complain if different
If (eBufferType <> WINHTTPxWEBxSOCKETxUTF8xMESSAGExBUFFERxTYPE) Then
  dwError = 87 ' ERRORxINVALIDxPARAMETER
  pErrorText = "Unexpected buffer type: " & eBufferType
  Exit Function
End If


' resize buffer to dwBytesTransferred, then convert buffer into VBA string
ReDim Preserve rgbBuffer(CInt(dwBytesTransferred))
GetMessageUTF8 = Utf8BytesToString(rgbBuffer)

Exit Function
ErrHandler:
    Debug.Print "GetMessageUTF8() - Error: " & Err.description
    Stop
    'Resume
End Function


Public Function SendMessageBinary(SendThis() As Byte) As Boolean
Dim dwError As Long
SendMessageBinary = False
If pState <> 3 Then
  pErrorText = "Must be connected to send"
  Exit Function
End If
Dim cdwMessageLength As Long
cdwMessageLength = (UBound(SendThis) - LBound(SendThis) + 1)
dwError = WinHttpWebSocketSend(hWebSocketHandle, _
             WINHTTPxWEBxSOCKETxBINARYxMESSAGExBUFFERxTYPE, VarPtr(SendThis(0)), cdwMessageLength)
If (dwError <> ERRORxSUCCESS) Then
  pErrorText = "SendMessageBinary Error: " & dwError & "/" & GetLastError
End If
SendMessageBinary = True
End Function

Public Function GetMessageBinary() As Byte()
pErrorText = "None"
If pState <> 3 Then
  pErrorText = "Must be connected to receive"
  Exit Function
End If

Dim otherBuffer() As Byte  ' do it this way, so we can redim it later
ReDim otherBuffer(1023) As Byte
Dim dwBufferLength As LongPtr
' get correct value in case of change to buffer size
dwBufferLength = (UBound(otherBuffer) - LBound(otherBuffer) + 1)

Dim pbCurrentBufferPointer As LongPtr
pbCurrentBufferPointer = VarPtr(otherBuffer(0))
Dim dwBytesTransferred As LongPtr
dwBytesTransferred = 0
Dim eBufferType As LongPtr
eBufferType = 0
Dim dwCloseReasonLength As Long
dwCloseReasonLength = 0
Dim usStatus As Integer
usStatus = 0

' if we get a message fragment packet, loop until we get a complete packet
Dim dwError As Long
Dim MessageComplete As Boolean
MessageComplete = False
Do Until (MessageComplete)
  If (dwBufferLength = 0) Then
    dwError = ERRORxNOTxENOUGHxMEMORY  ' note - not using underscores, can confuse class compilation
    pErrorText = "Need a bigger buffer!"
    Exit Function
  End If
  ' receive - will hang if no server response
  dwError = WinHttpWebSocketReceive(hWebSocketHandle, otherBuffer(0), CLng(dwBufferLength), dwBytesTransferred, eBufferType)
  If (dwError <> ERRORxSUCCESS) Then
    pErrorText = "GetMessageUTF8 Error: " & dwError & "/" & GetLastError
    Exit Function
  End If
  
  ' check for close message
  If (eBufferType = WINHTTPxWEBxSOCKETxCLOSExBUFFERxTYPE) Then ' server wants to quit
    pSCC = True
    pErrorText = "Server Closed Connection"
    Exit Function
  End If
  
' If we receive just part of the message repeat the receive operation.
  pbCurrentBufferPointer = pbCurrentBufferPointer + dwBytesTransferred
  dwBufferLength = dwBufferLength - dwBytesTransferred
  If Not (eBufferType = WINHTTPxWEBxSOCKETxBINARYxFRAGMENTxBUFFERxTYPE) Then
    MessageComplete = True
  End If
Loop

' Expected server to just echo single binary message - complain if different
If (eBufferType <> WINHTTPxWEBxSOCKETxBINARYxMESSAGExBUFFERxTYPE) Then
  dwError = 87 ' ERRORxINVALIDxPARAMETER
  pErrorText = "Unexpected buffer type: " & eBufferType
  Exit Function
End If

' resize buffer to dwBytesTransferred, then return
ReDim Preserve otherBuffer(CInt(dwBytesTransferred) - 1)
GetMessageBinary = otherBuffer
End Function







Private Sub Class_Initialize()
hSessionHandle = 0
hConnectionHandle = 0
hRequestHandle = 0
hWebSocketHandle = 0
pState = 1 ' not connected
pSCC = False
pAgentHeader = "CWebSocket"  ' default
pServer = ""  ' default, but will cause Connect method to fail
pPort = 80    ' default
pPath = "/"   ' default
  
End Sub

Private Sub Class_Terminate()
Call CleanUpHandles
End Sub

Private Sub CleanUpHandles()
  If (hWebSocketHandle <> 0) Then
    WinHttpCloseHandle (hWebSocketHandle)
    hWebSocketHandle = 0
  End If
  If (hRequestHandle <> 0) Then
    WinHttpCloseHandle (hRequestHandle)
    hRequestHandle = 0
  End If
  If (hConnectionHandle <> 0) Then
    WinHttpCloseHandle (hConnectionHandle)
    hConnectionHandle = 0
  End If
  If (hSessionHandle <> 0) Then
    WinHttpCloseHandle (hSessionHandle)
    hSessionHandle = 0
  End If
End Sub
