VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpTarget"
' Target
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' 
' TargetID AS string

' Unique identifier of attached debugging session.
' SessionID AS string

' 
' TargetInfo AS object

' 
' RemoteLocation AS object



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

' Activates (focuses) the target.
Public Sub activateTarget( _
    ByVal targetId AS string _
)
    ' targetId: TargetID 

    Dim params As New Dictionary
    params("targetId") = CStr(targetId)

    cdp.InvokeMethod "Target.activateTarget", params
End Sub

' Attaches to the target with given id.
' Returns: sessionId - Id assigned to the session.
Public Function attachToTarget( _
    ByVal targetId AS string, _
    Optional ByVal flatten AS Variant _
) AS string
    ' targetId: TargetID 
    ' flatten: boolean(optional) Enables "flat" access to the session via specifying sessionId attribute in the commands.
    '   We plan to make this the default, deprecate non-flattened mode,
    '   and eventually retire it. See crbug.com/991325.

    Dim params As New Dictionary
    params("targetId") = CStr(targetId)
    If Not IsMissing(flatten) Then params("flatten") = CBool(flatten)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.attachToTarget", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("sessionId") Then attachToTarget = results("sessionId")
    End If
End Function

' Attaches to the browser target, only uses flat sessionId mode.
' Returns: sessionId - Id assigned to the session.
' Experimental
Public Function attachToBrowserTarget() AS string
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.attachToBrowserTarget", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("sessionId") Then attachToBrowserTarget = results("sessionId")
    End If
End Function

' Closes the target. If the target is a page that gets closed too.
' Returns: success - Always set to true. If an error occurs, the response indicates protocol error.
Public Function closeTarget( _
    ByVal targetId AS string _
) AS boolean
    ' targetId: TargetID 

    Dim params As New Dictionary
    params("targetId") = CStr(targetId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.closeTarget", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("success") Then closeTarget = results("success")
    End If
End Function

' Inject object to the target's main frame that provides a communication
' channel with browser target.
' 
' Injected object will be available as `window[bindingName]`.
' 
' The object has the follwing API:
' - `binding.send(json)` - a method to send messages over the remote debugging protocol
' - `binding.onmessage = json => handleMessage(json)` - a callback that will be called for the protocol notifications and command responses.
' Experimental
Public Sub exposeDevToolsProtocol( _
    ByVal targetId AS string, _
    Optional ByVal bindingName AS Variant _
)
    ' targetId: TargetID 
    ' bindingName: string(optional) Binding name, 'cdp' if not specified.

    Dim params As New Dictionary
    params("targetId") = CStr(targetId)
    If Not IsMissing(bindingName) Then params("bindingName") = CStr(bindingName)

    cdp.InvokeMethod "Target.exposeDevToolsProtocol", params
End Sub

' Creates a new empty BrowserContext. Similar to an incognito profile but you can have more than
' one.
' Returns: browserContextId - The id of the context created.
' Experimental
Public Function createBrowserContext( _
    Optional ByVal disposeOnDetach AS Variant, _
    Optional ByVal proxyServer AS Variant, _
    Optional ByVal proxyBypassList AS Variant _
) AS string
    ' disposeOnDetach: boolean(optional) If specified, disposes this context when debugging session disconnects.
    ' proxyServer: string(optional) Proxy server, similar to the one passed to --proxy-server
    ' proxyBypassList: string(optional) Proxy bypass list, similar to the one passed to --proxy-bypass-list

    Dim params As New Dictionary
    If Not IsMissing(disposeOnDetach) Then params("disposeOnDetach") = CBool(disposeOnDetach)
    If Not IsMissing(proxyServer) Then params("proxyServer") = CStr(proxyServer)
    If Not IsMissing(proxyBypassList) Then params("proxyBypassList") = CStr(proxyBypassList)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.createBrowserContext", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("browserContextId") Then createBrowserContext = results("browserContextId")
    End If
End Function

' Returns all browser contexts created with `Target.createBrowserContext` method.
' Returns: browserContextIds - An array of browser context ids.
' Experimental
Public Function getBrowserContexts() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.getBrowserContexts", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("browserContextIds") Then Set getBrowserContexts = results("browserContextIds")
    End If
End Function

' Creates a new page.
' Returns: targetId - The id of the page opened.
Public Function createTarget( _
    ByVal url AS string, _
    Optional ByVal width AS Variant, _
    Optional ByVal height AS Variant, _
    Optional ByVal browserContextId AS Variant, _
    Optional ByVal enableBeginFrameControl AS Variant, _
    Optional ByVal newWindow AS Variant, _
    Optional ByVal background AS Variant _
) AS string
    ' url: string The initial URL the page will be navigated to. An empty string indicates about:blank.
    ' width: integer(optional) Frame width in DIP (headless chrome only).
    ' height: integer(optional) Frame height in DIP (headless chrome only).
    ' browserContextId: Browser.BrowserContextID(optional) The browser context to create the page in.
    ' enableBeginFrameControl: boolean(optional) Whether BeginFrames for this target will be controlled via DevTools (headless chrome only,
    '   not supported on MacOS yet, false by default).
    ' newWindow: boolean(optional) Whether to create a new Window or Tab (chrome-only, false by default).
    ' background: boolean(optional) Whether to create the target in background or foreground (chrome-only,
    '   false by default).

    Dim params As New Dictionary
    params("url") = CStr(url)
    If Not IsMissing(width) Then params("width") = CLng(width)
    If Not IsMissing(height) Then params("height") = CLng(height)
    If Not IsMissing(browserContextId) Then params("browserContextId") = CStr(browserContextId)
    If Not IsMissing(enableBeginFrameControl) Then params("enableBeginFrameControl") = CBool(enableBeginFrameControl)
    If Not IsMissing(newWindow) Then params("newWindow") = CBool(newWindow)
    If Not IsMissing(background) Then params("background") = CBool(background)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.createTarget", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("targetId") Then createTarget = results("targetId")
    End If
End Function

' Detaches session with given id.
Public Sub detachFromTarget( _
    Optional ByVal sessionId AS Variant, _
    Optional ByVal targetId AS Variant _
)
    ' sessionId: SessionID(optional) Session to detach.
    ' targetId: TargetID(optional) Deprecated.

    Dim params As New Dictionary
    If Not IsMissing(sessionId) Then params("sessionId") = CStr(sessionId)
    If Not IsMissing(targetId) Then params("targetId") = CStr(targetId)

    cdp.InvokeMethod "Target.detachFromTarget", params
End Sub

' Deletes a BrowserContext. All the belonging pages will be closed without calling their
' beforeunload hooks.
' Experimental
Public Sub disposeBrowserContext( _
    ByVal browserContextId AS string _
)
    ' browserContextId: Browser.BrowserContextID 

    Dim params As New Dictionary
    params("browserContextId") = CStr(browserContextId)

    cdp.InvokeMethod "Target.disposeBrowserContext", params
End Sub

' Returns information about a target.
' Returns: targetInfo - 
' Experimental
Public Function getTargetInfo( _
    Optional ByVal targetId AS Variant _
) AS Dictionary
    ' targetId: TargetID(optional) 

    Dim params As New Dictionary
    If Not IsMissing(targetId) Then params("targetId") = CStr(targetId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.getTargetInfo", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("targetInfo") Then Set getTargetInfo = results("targetInfo")
    End If
End Function

' Retrieves a list of available targets.
' Returns: targetInfos - The list of targets.
Public Function getTargets() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Target.getTargets", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("targetInfos") Then Set getTargets = results("targetInfos")
    End If
End Function

' Sends protocol message over session with given id.
' Consider using flat mode instead; see commands attachToTarget, setAutoAttach,
' and crbug.com/991325.
Public Sub sendMessageToTarget( _
    ByVal message AS string, _
    Optional ByVal sessionId AS Variant, _
    Optional ByVal targetId AS Variant _
)
    ' message: string 
    ' sessionId: SessionID(optional) Identifier of the session.
    ' targetId: TargetID(optional) Deprecated.

    Dim params As New Dictionary
    params("message") = CStr(message)
    If Not IsMissing(sessionId) Then params("sessionId") = CStr(sessionId)
    If Not IsMissing(targetId) Then params("targetId") = CStr(targetId)

    cdp.InvokeMethod "Target.sendMessageToTarget", params
End Sub

' Controls whether to automatically attach to new targets which are considered to be related to
' this one. When turned on, attaches to all existing related targets as well. When turned off,
' automatically detaches from all currently attached targets.
' This also clears all targets added by `autoAttachRelated` from the list of targets to watch
' for creation of related targets.
' Experimental
Public Sub setAutoAttach( _
    ByVal autoAttach AS boolean, _
    ByVal waitForDebuggerOnStart AS boolean, _
    Optional ByVal flatten AS Variant _
)
    ' autoAttach: boolean Whether to auto-attach to related targets.
    ' waitForDebuggerOnStart: boolean Whether to pause new targets when attaching to them. Use `Runtime.runIfWaitingForDebugger`
    '   to run paused targets.
    ' flatten: boolean(optional) Enables "flat" access to the session via specifying sessionId attribute in the commands.
    '   We plan to make this the default, deprecate non-flattened mode,
    '   and eventually retire it. See crbug.com/991325.

    Dim params As New Dictionary
    params("autoAttach") = CBool(autoAttach)
    params("waitForDebuggerOnStart") = CBool(waitForDebuggerOnStart)
    If Not IsMissing(flatten) Then params("flatten") = CBool(flatten)

    cdp.InvokeMethod "Target.setAutoAttach", params
End Sub

' Adds the specified target to the list of targets that will be monitored for any related target
' creation (such as child frames, child workers and new versions of service worker) and reported
' through `attachedToTarget`. The specified target is also auto-attached.
' This cancels the effect of any previous `setAutoAttach` and is also cancelled by subsequent
' `setAutoAttach`. Only available at the Browser target.
' Experimental
Public Sub autoAttachRelated( _
    ByVal targetId AS string, _
    ByVal waitForDebuggerOnStart AS boolean _
)
    ' targetId: TargetID 
    ' waitForDebuggerOnStart: boolean Whether to pause new targets when attaching to them. Use `Runtime.runIfWaitingForDebugger`
    '   to run paused targets.

    Dim params As New Dictionary
    params("targetId") = CStr(targetId)
    params("waitForDebuggerOnStart") = CBool(waitForDebuggerOnStart)

    cdp.InvokeMethod "Target.autoAttachRelated", params
End Sub

' Controls whether to discover available targets and notify via
' `targetCreated/targetInfoChanged/targetDestroyed` events.
Public Sub setDiscoverTargets( _
    ByVal discover AS boolean _
)
    ' discover: boolean Whether to discover available targets.

    Dim params As New Dictionary
    params("discover") = CBool(discover)

    cdp.InvokeMethod "Target.setDiscoverTargets", params
End Sub

' Enables target discovery for the specified locations, when `setDiscoverTargets` was set to
' `true`.
' Experimental
Public Sub setRemoteLocations( _
    ByVal locations AS Collection _
)
    ' locations: array List of remote locations.

    Dim params As New Dictionary
    Set params("locations") = locations

    cdp.InvokeMethod "Target.setRemoteLocations", params
End Sub

