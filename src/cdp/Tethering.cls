VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpTethering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpTethering"
' Tethering [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' None


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

' Request browser port binding.
Public Sub bind( _
    ByVal port AS Long _
)
    ' port: integer Port number to bind.

    Dim params As New Dictionary
    params("port") = CLng(port)

    cdp.InvokeMethod "Tethering.bind", params
End Sub

' Request browser port unbinding.
Public Sub unbind( _
    ByVal port AS Long _
)
    ' port: integer Port number to unbind.

    Dim params As New Dictionary
    params("port") = CLng(port)

    cdp.InvokeMethod "Tethering.unbind", params
End Sub

