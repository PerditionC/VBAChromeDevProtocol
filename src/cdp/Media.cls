VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpMedia"
' Media [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Players will get an ID that is unique within the agent context.
' PlayerId AS string

' 
' Timestamp AS number

' Have one type per entry in MediaLogRecord::Type
'   Corresponds to kMessage
' PlayerMessage AS object

' Corresponds to kMediaPropertyChange
' PlayerProperty AS object

' Corresponds to kMediaEventTriggered
' PlayerEvent AS object

' Corresponds to kMediaError
' PlayerError AS object



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

' Enables the Media domain
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Media.enable", params
End Sub

' Disables the Media domain.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Media.disable", params
End Sub

