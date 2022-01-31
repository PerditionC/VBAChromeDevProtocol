VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpCast"
' Cast [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' 
' Sink AS object



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

' Starts observing for sinks that can be used for tab mirroring, and if set,
' sinks compatible with |presentationUrl| as well. When sinks are found, a
' |sinksUpdated| event is fired.
' Also starts observing for issue messages. When an issue is added or removed,
' an |issueUpdated| event is fired.
Public Sub enable( _
    Optional ByVal presentationUrl AS Variant _
)
    ' presentationUrl: string(optional) 

    Dim params As New Dictionary
    If Not IsMissing(presentationUrl) Then params("presentationUrl") = CStr(presentationUrl)

    cdp.InvokeMethod "Cast.enable", params
End Sub

' Stops observing for sinks and issues.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Cast.disable", params
End Sub

' Sets a sink to be used when the web page requests the browser to choose a
' sink via Presentation API, Remote Playback API, or Cast SDK.
Public Sub setSinkToUse( _
    ByVal sinkName AS string _
)
    ' sinkName: string 

    Dim params As New Dictionary
    params("sinkName") = CStr(sinkName)

    cdp.InvokeMethod "Cast.setSinkToUse", params
End Sub

' Starts mirroring the tab to the sink.
Public Sub startTabMirroring( _
    ByVal sinkName AS string _
)
    ' sinkName: string 

    Dim params As New Dictionary
    params("sinkName") = CStr(sinkName)

    cdp.InvokeMethod "Cast.startTabMirroring", params
End Sub

' Stops the active Cast session on the sink.
Public Sub stopCasting( _
    ByVal sinkName AS string _
)
    ' sinkName: string 

    Dim params As New Dictionary
    params("sinkName") = CStr(sinkName)

    cdp.InvokeMethod "Cast.stopCasting", params
End Sub

