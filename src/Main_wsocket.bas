Attribute VB_Name = "Main_wsocket"
Option Explicit

' Note: synchronous communication - fails if server does not reply
' ===================================================
' =  Echo server thanks to: https://websocket.org/  =
' ===================================================


Sub NewMain()
Dim ws As clsWebSocket
Set ws = New clsWebSocket

'string
Dim stringMessage As String
stringMessage = "Hello World!"

'binary
Dim k As Long
Dim binaryMessage(7) As Byte
For k = 0 To 7
  binaryMessage(k) = k + 1
Next

With ws
  .Server = "echo.websocket.events"  ' Note: echo.websocket.org shutdown in 2021
  .Connect
  ' send string
  .SendMessageUTF8 (stringMessage)
  Debug.Print "Server string reply: " & .GetMessageUTF8
  
  ' check after receiving each message to see if server wants to quit
  If .SCC Then ' close connection
    .Disconnect
  Else
    ' send binary
    .SendMessageBinary binaryMessage()
    ' convert reply
    Dim stringReply As String
    Dim bytesReply() As Byte
    bytesReply = .GetMessageBinary
    'For k = LBound(bytesReply) To UBound(bytesReply)
    '  stringReply = stringReply & bytesReply(k)
    'Next
    Debug.Print "Server binary reply: " & stringReply
    ' for this example, at this point, don't care if server wants to quit
  End If
  .Disconnect
End With
Set ws = Nothing
End Sub

