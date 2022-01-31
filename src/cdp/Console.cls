VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpConsole"
' Console
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Console message.
' ConsoleMessage AS object



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

' Does nothing.
Public Sub clearMessages()
    Dim params As New Dictionary
    cdp.InvokeMethod "Console.clearMessages", params
End Sub

' Disables console domain, prevents further console messages from being reported to the client.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Console.disable", params
End Sub

' Enables console domain, sends the messages collected so far to the client by means of the
' `messageAdded` notification.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Console.enable", params
End Sub

