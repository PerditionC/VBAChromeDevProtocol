VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDeviceOrientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpDeviceOrientation"
' DeviceOrientation [Experimental]
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

' Clears the overridden Device Orientation.
Public Sub clearDeviceOrientationOverride()
    Dim params As New Dictionary
    cdp.InvokeMethod "DeviceOrientation.clearDeviceOrientationOverride", params
End Sub

' Overrides the Device Orientation.
Public Sub setDeviceOrientationOverride( _
    ByVal alpha AS Double, _
    ByVal beta AS Double, _
    ByVal gamma AS Double _
)
    ' alpha: number Mock alpha
    ' beta: number Mock beta
    ' gamma: number Mock gamma

    Dim params As New Dictionary
    params("alpha") = CDbl(alpha)
    params("beta") = CDbl(beta)
    params("gamma") = CDbl(gamma)

    cdp.InvokeMethod "DeviceOrientation.setDeviceOrientationOverride", params
End Sub

