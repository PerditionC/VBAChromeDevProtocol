VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutomateBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' adds a higher level API to Chrome Developer Protocol
Option Explicit

Public cdp As clsCDP


' how waiting after navigation should be handled
Public Enum NavigationStrategy
    Normal      ' default, waits for page load event
    eager       ' waits only for dom access to be ready, but other resources may still be loading
    NetworkIdle ' waits for network to be idle, currently no more than N network connections in last 500ms
    WindowOpen  ' waits for window open event
    None        ' does not wait at all
End Enum

' maximum wait (timeout) when waiting for page loads
Private Const maxWaitSeconds As Integer = 60

' information on all currently attached targets
Public sessions As Dictionary ' of clsSessionInfo

' primary session
Public mainSession As clsSessionInfo

' currently active session
Private curSession As clsSessionInfo

' list of active targets
' keyed on targetId
' Note: current target info is retrievable via attachedTargets(curSession.targetInfo.targetId)
'       and parent windows is retrievable via attachedTargets(curSession.targetInfo.openerId)
Private attachedTargets As Dictionary ' of cdpTargetTargetInfo


Private Sub Class_Initialize()
    Set cdp = New clsCDP
    ' register event handlers that we use
    cdp.registerEventHandler "Target.attachedToTarget", Me
    cdp.registerEventHandler "Target.detachedFromTarget", Me
    cdp.registerEventHandler "Target.targetCreated", Me
    cdp.registerEventHandler "Target.targetInfoChanged", Me
    cdp.registerEventHandler "Target.targetDestroyed", Me
    cdp.registerEventHandler "Runtime.executionContextCreated", Me
    cdp.registerEventHandler "Runtime.executionContextDestroyed", Me
    cdp.registerEventHandler "Runtime.executionContextsCleared", Me
    cdp.registerEventHandler "Page.frameStartedLoading", Me
    cdp.registerEventHandler "Page.frameStoppedLoading", Me
    cdp.registerEventHandler "Page.domContentEventFired", Me
    cdp.registerEventHandler "Page.frameRequestedNavigation", Me
    cdp.registerEventHandler "Page.frameNavigated", Me
    cdp.registerEventHandler "Page.windowOpen", Me
    cdp.registerEventHandler "Network.requestWillBeSent", Me
    cdp.registerEventHandler "Network.responseReceived", Me
    cdp.registerEventHandler "DOM.documentUpdated", Me
    cdp.registerEventHandler "DOM.setChildNodes", Me
    cdp.registerEventHandler "DOM.attributeModified", Me
    
    ' other state variables
    Set mainSession = New clsSessionInfo ' initialized by call to launch
    Set sessions = New Dictionary
    Set attachedTargets = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set cdp = Nothing
    Set mainSession = Nothing
    Set sessions = Nothing
    Set attachedTargets = Nothing
End Sub


' internal event handler
' Note: if registering any event handled here, caller must be sure to call to our event handler!
Public Function processEvent(ByVal eventName As String, ByVal eventData As Dictionary) As Boolean
    On Error GoTo errHandler
    Dim frameId As String, frame As cdpPageFrame
    Dim requestId As String, requestInfo As clsNetworkRequest
    Dim contextId As Long, context As cdpRuntimeExecutionContextDescr
    Dim targetId As String, targetInfo As cdpTargetTargetInfo
    Dim sessionId As String, session As clsSessionInfo
    If eventData.Exists("sessionId") Then sessionId = eventData("sessionId")
    If sessions.Exists(sessionId) Then
        Set session = sessions(sessionId)
    Else
        Set session = curSession
    End If
    
    Dim params As Dictionary
    If eventData.Exists("params") Then
        Set params = eventData("params")
    Else
        Set params = New Dictionary ' so we don't have to keep testing if Nothing, just if expected parameter exists
    End If
    
    Select Case eventName
        Case "Page.frameStartedLoading"
            session.eventFrameStoppedLoading = False
            'session.eventDomContentEventFired = False ' can arrive out of order, ie DomContentEventFired before frameStartLoading
            If params.Exists("frameId") Then
                frameId = params("frameId")
                session.framesLoading.Item(frameId) = "loading" ' add or replace, value is arbitrary and currently unused
            End If
        Case "Page.frameStoppedLoading"
            If params.Exists("frameId") Then
                frameId = params("frameId")
                If session.framesLoading.Exists(frameId) Then
                    session.framesLoading.Remove frameId
                End If
            End If
            session.eventFrameStoppedLoading = session.framesLoading.count < 1
        Case "Page.domContentEventFired"
            session.eventDomContentEventFired = True
        Case "Page.frameNavigated"
            If params.Exists("type") Then
                Debug.Print "Navigation type is " & params("type")
            End If
            If params.Exists("frame") Then
                Set frame = New cdpPageFrame
                frame.init params("frame")
                Set session.frames(frame.id) = frame ' add or replace
                If frame.name <> vbNullString Then session.frameNames.Item(frame.name) = frame.id ' add or replace
            End If
        Case "Page.frameDetached"
            If params.Exists("frameId") Then
                frameId = params("frameId")
                If session.frames.Exists(frameId) Then
                    If frame.name <> vbNullString Then
                        If session.frameNames.Exists(frame.name) Then session.frameNames.Remove frame.name
                    End If
                    session.frames.Remove frameId
                End If
                If session.framesLoading.Exists(frameId) Then session.framesLoading.Remove frameId
            End If
        Case "Page.frameRequestedNavigation"
        Case "Page.frameNavigated"
        Case "Page.windowOpen"
            ' can be used to get url about to navigate to,e.g. if opened a window the url of new window
            session.windowOpenEventOccurred = True
        Case "Runtime.executionContextCreated"
            ' update add frame/context mapping
            If params.Exists("context") Then
                Set context = New cdpRuntimeExecutionContextDescr
                context.init params("context")
                If context.auxData.Exists("frameId") Then
                    frameId = context.auxData.Item("frameId")
                    ' map FrameId (string) to ExecutionContextId (integer)
                    session.frameExecCtxMapping.Item(frameId) = CLng(context.id)
                    ' map FrameId (string) to Unique(Execution)Id (string)
                    session.frameExecCtxMapping.Item("$" & frameId) = CStr(context.uniqueId)
                    ' map ExecutionId (integer) to FrameId (string)
                    session.frameExecCtxMapping.Item(context.id) = CStr(frameId)
                End If
            End If
        Case "Network.requestWillBeSent"
            If params.Exists("requestId") Then
                Set requestInfo = New clsNetworkRequest
                requestInfo.state = NetworkRequestState.NRS_Pending
                ' parameters that should be set
                requestInfo.requestId = params("requestId")
                requestInfo.loaderId = params("loaderId")
                requestInfo.documentURL = params("documentURL")
                'requestInfo.request = params("request")
                requestInfo.timestampStart = params("timestamp")
                requestInfo.wallTime = params("wallTime")
                requestInfo.monotonicTimeOffset = requestInfo.wallTime - requestInfo.timestampStart
                requestInfo.vbaTimestampStart = Now
                If params.Item("initiator").Exists("type") Then requestInfo.initiator = params.Item("initiator").Item("type")
                requestInfo.hasExtraInfo = CBool(params("redirectHasExtraInfo"))
                ' optional parameters
                'If params.Exists("redirectResponse") Then requestInfo.response = params("redirectResponse")
                If params.Exists("type") Then requestInfo.resourceType = params("type")
                If params.Exists("frameId") Then requestInfo.frameId = params("frameId")
                If params.Exists("hasUserGesture") Then requestInfo.hasUserGesture = CBool(params("hasUserGesture"))
                
                Set session.pendingNetworkRequest.Item(requestInfo.requestId) = requestInfo
            End If
        Case "Network.responseReceived"
            requestId = params("requestId")
            If session.pendingNetworkRequest.Exists(requestId) Then
                Set requestInfo = session.pendingNetworkRequest.Item(requestId)
                
                requestInfo.timestampStop = params("timestamp")
                requestInfo.vbaTimestampStop = Now
            Else
                Set requestInfo = New clsNetworkRequest
            
                requestInfo.requestId = params("requestId")
                requestInfo.loaderId = params("loaderId")
                requestInfo.timestampStop = params("timestamp")
                requestInfo.vbaTimestampStop = Now
                requestInfo.resourceType = params("type")
                requestInfo.hasExtraInfo = CBool(params("hasExtraInfo"))
                'requestInfo.response = params("response")
                If params.Exists("frameId") Then requestInfo.frameId = params("frameId")
            End If
            requestInfo.state = NetworkRequestState.NRS_Complete
            ' update/add
            Set session.pendingNetworkRequest.Item(requestId) = requestInfo
        Case "Runtime.executionContextDestroyed"
            ' remove stale mappings
            If params.Exists("executionContextId") Then
                contextId = params("executionContextId")
                If session.frameExecCtxMapping.Exists(contextId) Then
                    frameId = session.frameExecCtxMapping.Item(contextId)
                    session.frameExecCtxMapping.Remove contextId
                    If session.frameExecCtxMapping.Exists(frameId) Then
                        session.frameExecCtxMapping.Remove frameId
                        session.frameExecCtxMapping.Remove "$" & frameId
                    End If
                End If
            End If
        Case "Runtime.executionContextsCleared"
            ' reset frame/context mapping
            session.frameExecCtxMapping.RemoveAll
        Case "Target.targetCreated"
            ' add/update target list - ie new windows that are opened
            If params.Exists("targetInfo") Then
                Set targetInfo = New cdpTargetTargetInfo
                targetInfo.init params("targetInfo")
                Set attachedTargets(targetInfo.targetId) = targetInfo
            End If
        Case "Target.targetInfoChanged"
            ' TODO update attachedTargets
        Case "Target.targetDestroyed"
            If params.Exists("targetId") Then
                targetId = params("targetId")
                If attachedTargets.Exists(targetId) Then
                    attachedTargets.Remove targetId
                End If
            End If
        Case "Target.attachedToTarget"
            ' maintain cached values associated with attached window
            If params.Exists("sessionId") Then
                Set session = New clsSessionInfo
                session.sessionId = params("sessionId")
                Set session.targetInfo = New cdpTargetTargetInfo
                If params.Exists("targetInfo") Then session.targetInfo.init params("targetInfo")
                Set sessions(session.sessionId) = session
            End If
        Case "Target.detachedFromTarget"
            If params.Exists("sessionId") Then
                sessionId = params("sessionId")
                ' if detached from current session then use primary session until explicit switch is made
                If sessionId = curSession.sessionId Then
                    Set curSession = mainSession
                    cdp.sessionId = mainSession.sessionId
                End If
                sessions.Remove sessionId
            End If
        Case "DOM.documentUpdated"
            ' mark cache invalid
            Set session.rootNode = Nothing
        Case "DOM.setChildNodes"
            ' update cached nodes
            If Not session.rootNode Is Nothing Then
                Debug.Print "Child Nodes added"
            End If
        Case "DOM.attributeModified"
            ' update attribute
            Debug.Print "attribut modified: " & JsonConverter.ConvertToJson(params)
        Case Else
            ' do nothing, we will receive more than we care about
    End Select
    Exit Function
errHandler:
    Debug.Print Err.description
    Stop
    Resume
End Function



' returns current position and state of window
' does not include if window is hidden or visible, assumed visible
' windowId is optional, browser window id to return bounds for, defaults to current window attached to (current target)
' Returns WindowsRect with x,y & width,height of window
'         if state variable is provided, then will set to current WindowState
Public Function getWindowBounds(Optional ByVal windowId As Variant, Optional ByRef state) As cdpBrowserBounds
    Dim bounds As Dictionary
    If IsMissing(windowId) Then
        Dim dict As Dictionary
        Set dict = cdp.Browser.getWindowForTarget ' returns both windowId and current window information
        windowId = dict("windowId")
        Set bounds = dict("bounds")
    Else
        Set bounds = cdp.Browser.getWindowBounds(windowId)
    End If
    
    Dim browserBounds As cdpBrowserBounds: Set browserBounds = New cdpBrowserBounds
    browserBounds.init bounds
    
    If Not IsMissing(state) Then
        state = StringToWindowState(browserBounds.windowState)
    End If
    
    Set getWindowBounds = browserBounds
End Function

' sets just the window state
' windowId is the id of the window to change
' state is the desired state to set window to, one of: normal, minimized, maximized, ...
Public Sub setWindowState(ByVal windowId As Long, ByVal state As windowState)
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict("windowId") = windowId
    dict("windowState") = WindowStateToString(state)
    cdp.Browser.setWindowBounds windowId, dict
End Sub


' gets browser window hWnd
Public Function getBrowserHWnd() As Long
    ' if not cached, then find hWnd
    If curSession.browserHWnd = 0 Then
        Dim title As String
        title = getTitle()
        'title = Me.Document.getElementsByTagName("title").index(1).InnerText
        curSession.browserHWnd = findBrowserHWnd(title, limitTo:=cdp.browserProgram)
    End If
    
    getBrowserHWnd = curSession.browserHWnd
End Function

' bring browser to foreground
Public Sub focusBrowser()
    BringWindowToTop getBrowserHWnd
End Sub

' if hidden, shows browser window, if minimized or maximized will restore to normal
Public Sub showNormal()
    Dim hWnd As Long: hWnd = getBrowserHWnd
    If IsIconic(hWnd) Then
        ShowWindowAsync hWnd, ShowWindowCmd.SW_RESTORE
    Else
        ShowWindowAsync hWnd, ShowWindowCmd.SW_NORMAL
    End If
End Sub

Public Sub ShowMaximized()
    ShowWindowAsync getBrowserHWnd, ShowWindowCmd.SW_MAXIMIZE
    'resizeWindow state:=WindowState.Maximized
End Sub

Public Sub ShowMinimized()
    ShowWindowAsync getBrowserHWnd, ShowWindowCmd.SW_MINIMIZE
    'resizeWindow state:=WindowState.Minimized
End Sub

' hides browser window
Public Sub hide()
    ShowWindowAsync getBrowserHWnd, ShowWindowCmd.SW_HIDE
End Sub

' adjust positions, size, and/or state of browser window
' if position or size specified and window currently maximized, window will be returned to normal - moved/sized then maximized again
' if window is hidden, setting to any other state will unhide it
' Long: windowId is the window to resize, defaults to current window attached to (active target)
' Long: left, top are the X,Y position of the upper left corner of the window; note negative values are for secondary screen to the left of/above the primary screen
' Long: width, height are size of the window
Public Sub resizeWindow(Optional ByVal windowId As Variant, _
                  Optional ByVal left As Variant, Optional ByVal top As Variant, _
                  Optional ByVal width As Variant, Optional ByVal height As Variant, _
                  Optional ByVal state As Variant)
    
    Dim rect As cdpBrowserBounds
    If IsMissing(windowId) Then
        ' we need to get windowId and bounds
        Dim dict As Dictionary
        Set dict = cdp.Browser.getWindowForTarget ' returns both windowId and current window information
        If dict Is Nothing Then Exit Sub ' exit early and leave error message intact if get window information failed
        windowId = dict("windowId")
        Set rect = New cdpBrowserBounds
        rect.init dict("bounds")
    Else
        Set rect = getWindowBounds(windowId)
    End If
    
    ' determine if maximized or not, if so then temp restore to normal for resize (but only if resizing)
    If rect.windowState = "maximized" Then ' windowState.Maximized
        If Not (IsMissing(left) And IsMissing(top) And IsMissing(width) And IsMissing(height)) Then
            setWindowState windowId, windowState.WS_normal
        End If
    End If
    
    ' resize window
    If Not (IsMissing(left) And IsMissing(top) And IsMissing(width) And IsMissing(height)) Then
        Dim bounds As Dictionary: Set bounds = New Dictionary ' we don't use .asDictionary as we are using different values
        If Not IsMissing(left) Then bounds("left") = CLng(left)
        If Not IsMissing(top) Then bounds("top") = CLng(top)
        If Not IsMissing(width) Then bounds("width") = CLng(width)
        If Not IsMissing(height) Then bounds("height") = CLng(height)
        cdp.Browser.setWindowBounds windowId, bounds
    End If
    
    ' set window state (restores to maximized if needed)
    If (Not IsMissing(state)) Or (rect.windowState = "maximized") Then 'windowState.Maximized
        setWindowState windowId, state
    End If
End Sub

' This method starts (and attaches) up the browser, see cdp.launch
' Returns True if successfully initialized browser, False otherwise
' Warning! the url must exactly match and may fail if browser changes expected url unless partialMatch is True
Public Function launch( _
    Optional url As String = "about:blank", _
    Optional partialMatch As Boolean = True, _
    Optional useWebSocket As Boolean = False, _
    Optional useExistingBrowser As Boolean = False, _
    Optional reconnectUrl As String = vbNullString, _
    Optional byVal whichBrowser as browserType = anyBrowser, _
    Optional byVal additionnalInlineCommands as string = vbNullString, _
    Optional byVal killWithoutAsking as boolean = False _
    ) As Boolean
    If cdp.launch(url, autoAttach:=True, partialMatch:=partialMatch, useWebSocket:=useWebSocket, useExistingBrowser:=useExistingBrowser, reconnectUrl:=reconnectUrl, whichBrowser:=whichBrowser, additionnalInlineCommands :=additionnalInlineCommands, killWithoutAsking:=killWithoutAsking) Then
        ' initialize our original session
        If cdp.sessionId = vbNullString Then Stop
        Set mainSession = sessions(cdp.sessionId)
        Set curSession = mainSession
    
        launch = True
    End If
End Function

' switch focus to a different window (target)
' returns current (prior to switch) target
Public Function switchTo(ByVal sessionId As String) As String
    If sessions.Exists(sessionId) Then
        Set curSession = sessions(sessionId)
        cdp.sessionId = curSession.sessionId
    End If
End Function

' convenience function to return to initial window (page target)
Public Sub switchToMainWindow()
    switchTo mainSession.sessionId
End Sub

' returns a Collection of child targets, i.e. new targets/windows opened from current session
Public Function getChildTargets(Optional targetId As String = vbNullString) As Collection ' of cdpTargetTargetInfo
    Dim children As Collection, v As Variant
    Set children = New Collection: Set getChildTargets = children
    
    If targetId = vbNullString Then targetId = curSession.targetInfo.targetId
    
    For Each v In cdp.Target.getTargets
        Dim Target As cdpTargetTargetInfo
        Set Target = New cdpTargetTargetInfo
        Target.init v
        If Target.fieldType = "page" Then
            If Target.openerId = targetId Then children.Add Target
        End If
    Next v
End Function

' attaches to all children targets/windows not already attached to
Public Sub attachToChildTargets(Optional waitStrategy As NavigationStrategy = NavigationStrategy.NetworkIdle)
    Dim children As Collection, v As Variant
    Set children = getChildTargets()
    For Each v In children
        Dim v2 As Variant, Target As cdpTargetTargetInfo, foundSession As Boolean
        Set Target = v
        ' first check if we already have a session for this target
        foundSession = False
        For Each v2 In sessions.items
            Dim session As clsSessionInfo
            Set session = v2
            If session.targetInfo.targetId = Target.targetId Then
                foundSession = True
                Exit For
            End If
        Next v2
        ' if not, then attach and optionally wait for page to finish loading
        If Not foundSession Then
            Dim sessionId As String
            sessionId = cdp.attach(targetId:=Target.targetId)
            If waitStrategy <> NavigationStrategy.None Then
                If sessions.Exists(sessionId) Then
                    Set session = sessions(sessionId)
                    waitForPageToLoad waitStrategy, session
                End If
            End If
        End If
    Next v
End Sub

' waits for page loading complete events
' returns True if timeout expired instead of event received
Private Function waitForPageToLoad(strategy As NavigationStrategy, Optional session As clsSessionInfo)
    If session Is Nothing Then Set session = curSession
    
    ' ensure events are enabled on specified session so we can transition
    Dim oldSessionId As String: oldSessionId = curSession.sessionId: curSession.sessionId = session.sessionId
    If strategy = NavigationStrategy.NetworkIdle Then cdp.Network.enable
    If strategy = NavigationStrategy.Normal Then cdp.Page.enable
    If strategy = NavigationStrategy.eager Or strategy = NavigationStrategy.Normal Then cdp.DOM.enable
    
    If strategy <> NavigationStrategy.None Then
        ' wait
        Dim maxWait As Date: maxWait = DateAdd("s", maxWaitSeconds, Now)
        Do
            cdp.peakMessage
            If strategy = NavigationStrategy.NetworkIdle And (session.pendingNetworkRequests() < 1) Then Exit Do ' no recent network activity
            If strategy = NavigationStrategy.Normal And session.eventFrameStoppedLoading And session.eventDomContentEventFired Then Exit Do ' all frames stopped loading and domcontent done
            If strategy = NavigationStrategy.eager And session.eventDomContentEventFired Then Exit Do ' some frames may still be loading or additional resources, but domcontent done
            If strategy = NavigationStrategy.WindowOpen And session.windowOpenEventOccurred Then Exit Do
            
            If Now > maxWait Then  ' timeout
                waitForPageToLoad = True
                Exit Do
            End If
        Loop While True
        'Debug.Print mainSession.pendingNetworkRequests()
        
        Sleep 0.5 ' give page a little bit longer to stabilize
        cdp.peakMessage ' process any last minute messages
        
        ' turn off events
        ' we need to ensure Page events are enabled so we can get notified of load complete or dom complete events
        'If strategy = NavigationStrategy.Normal Or strategy = NavigationStrategy.eager Then cdp.Page.disable
        ' we need to ensure Network events are enabled so we can see when network activity has subsided
        If strategy = NavigationStrategy.NetworkIdle Then cdp.Network.disable
        
        ' reset for next wait
        session.eventFrameStoppedLoading = False
        session.eventDomContentEventFired = False
        session.windowOpenEventOccurred = False
    End If
    
    ' restore sessionId
    curSession.sessionId = oldSessionId
End Function

' goto specified url
Public Function navigate( _
    ByVal url As String, _
    Optional ByVal referrer As Variant, _
    Optional ByVal transitionType As Variant, _
    Optional ByVal frameId As Variant, _
    Optional ByVal referrerPolicy As Variant, _
    Optional strategy As NavigationStrategy = NavigationStrategy.Normal _
) As Dictionary
    ' reset flag for window open event
    If Not curSession Is Nothing Then
        'curSession.eventFrameStoppedLoading = False
        curSession.eventDomContentEventFired = False
        curSession.windowOpenEventOccurred = False
    End If
    ' we need to ensure Page events are enabled so we can get notified of load complete or dom complete events
    If strategy = NavigationStrategy.Normal Or strategy = NavigationStrategy.eager Then cdp.Page.enable
    ' we need to ensure Network events are enabled so we can see when network activity has subsided
    If strategy = NavigationStrategy.NetworkIdle Then cdp.Network.enable
    
    Set navigate = cdp.Page.navigate(url, referrer, transitionType, frameId, referrerPolicy)
    
    waitForPageToLoad strategy
End Function

' This method tells the browser to close
Public Sub quit()
    ' this seems to work, but should we cycle through all active windows and close them?
    cdp.Browser.closeBrowser
    Debug.Print "Terminating browser session..."
End Sub

' display an alert box on the web page
Public Sub alert(message As String)
    message = Replace(message, """", "'") ' convert double quotes to single quotes
    Debug.Print message
    jsEval "alert(""" & message & """);"
End Sub

' returns current Title of browser window
Public Function getTitle() As String
    getTitle = mainSession.targetInfo.title
End Function

' This function allows to evaulate a javascript expression in the context of the page
' see Runtime.evaluate for details - minimally processes javascript, evaluates it, and if returns a value, unwraps it
Public Function jsEval( _
    expression As String, _
    Optional objectGroup As Variant, _
    Optional includeCommandLineAPI As Variant, _
    Optional silent As Variant, _
    Optional contextId As Variant, _
    Optional returnByValue As Variant, _
    Optional generatePreview As Variant, _
    Optional userGesture As Variant, _
    Optional awaitPromise As Variant, _
    Optional throwOnSideEffect As Variant, _
    Optional timeout As Variant, _
    Optional disableBreaks As Variant, _
    Optional replMode As Variant, _
    Optional allowUnsafeEvalBlockedByCSP As Variant, _
    Optional uniqueContextId As Variant _
) As Variant
    Dim objRes As Dictionary
    
    ' automatically terminate expression with semicolon if not already terminated
    Dim lastChar As String: lastChar = Right(Trim(expression), 1)
    If (lastChar <> ";") And (lastChar <> "}") Then expression = expression & ";"
    
    ' evaluate the javascript expression
    Set objRes = cdp.Runtime.evaluate(expression, objectGroup, includeCommandLineAPI, silent, contextId, returnByValue, generatePreview, userGesture, awaitPromise, throwOnSideEffect, timeout, disableBreaks, replMode, allowUnsafeEvalBlockedByCSP, uniqueContextId)
       
    If cdp.ErrorCode <> 0 Then
        ' Oops, there was an error in our javascript expression
        Stop
        Exit Function
    Else
        ' If the return type has a specific type
        ' we can return the result
        Set jsEval = objRes
        If objRes.Exists("result") Then
            Set objRes = objRes.Item("result")
            If objRes.Exists("type") Then
                Dim valueType As String: valueType = objRes.Item("type")
                If (valueType = "string") Or (valueType = "boolean") Or (valueType = "number") Or (valueType = "integer") Then
                    jsEval = objRes.Item("value")
                ElseIf (valueType = "object") Or (valueType = "array") Then
                    Set jsEval = objRes ' return the Dictionary that represents the object or Collecton that represents the array
                ElseIf (valueType = "undefined") Then
                    jsEval = vbEmpty
                End If
            End If
        End If
    End If
End Function


' given a quad Collection returns a rect Dictionary
' Returns Dictionary with x,y,width,height of corresponding quad [x1,y1,x2,y2,x3,y3,x4,y4]
Public Function quadToRect(quad As Collection) As Dictionary
    Dim topLeftX As Double, topLeftY As Double, topRightX As Double, topRightY As Double, bottomRightX As Double, bottomRightY As Double, bottomLeftX As Double, bottomLeftY As Double
    topLeftX = quad(1): topLeftY = quad(2)
    topRightX = quad(3): topRightY = quad(4)
    bottomRightX = quad(5): bottomRightY = quad(6)
    bottomLeftX = quad(7): bottomLeftY = quad(8)
    
    Dim x As Double, y As Double, width As Double, height As Double
    x = topLeftX: y = topLeftY
    width = topRightX - topLeftX: height = bottomLeftY - topLeftY
    
    Set quadToRect = New Dictionary
    quadToRect.Item("x") = x: quadToRect.Item("y") = y: quadToRect.Item("width") = width: quadToRect.Item("height") = height
End Function

' given a nodeId (or backendNodeId), returns a rect Dictionary of the node's position and size (relative CSS pixels)
Public Function getNodeRect(Optional nodeId As Long, Optional backendNodeId) As Dictionary
    Dim dict As Dictionary, quad As Collection
    If backendNodeId = 0 Then
        Set dict = cdp.DOM.getBoxModel(nodeId:=nodeId)
    Else
        Set dict = cdp.DOM.getBoxModel(backendNodeId:=backendNodeId)
    End If
    Set quad = dict.Item("content")
    Dim x As Double, y As Double, width As Double, height As Double
    x = quad(1): y = quad(2): width = CDbl(dict("width")): height = CDbl(dict("height"))
    Set dict = New Dictionary
    dict("x") = x: dict("y") = y: dict("width") = width: dict("height") = height
    Set getNodeRect = dict
End Function

' given a nodeId (or backendNodeId), calculates the center of the node
Public Sub getNodeCenter(ByRef x As Double, ByRef y As Double, Optional nodeId As Long = 0, Optional backendNodeId As Long = 0)
    ' get rect of node
    Dim dict As Dictionary, quad As Collection
    If nodeId <> 0 Then
        Set dict = cdp.DOM.getBoxModel(nodeId:=nodeId)
    Else
        Set dict = cdp.DOM.getBoxModel(backendNodeId:=nodeId)
    End If
    If cdp.ErrorCode = 0 Then
        Set quad = dict.Item("content")
        Dim x1 As Double, y1 As Double, width As Double, height As Double
        x1 = quad(1): y1 = quad(2): width = CDbl(dict("width")): height = CDbl(dict("height"))
        ' and use roughly its center
        x = x1 + (width / 2#): y = y1 + (height / 2#)
    End If
End Sub

' attempt to make node the currently focused (active) element
Public Sub focus(Optional nodeId As Long = 0, Optional backendNodeId As Long = 0, Optional objectId As String = vbNullString)
    If nodeId <> 0 Then
        cdp.DOM.Focus nodeId:=nodeId
        cdp.DOM.scrollIntoViewIfNeeded nodeId:=nodeId
    ElseIf objectId <> vbNullString Then
        cdp.DOM.Focus objectId:=objectId
        cdp.DOM.scrollIntoViewIfNeeded objectId:=objectId
    Else
        cdp.DOM.Focus backendNodeId:=backendNodeId
        cdp.DOM.scrollIntoViewIfNeeded backendNodeId:=backendNodeId
    End If
End Sub

' Simulates a user clicking a node
' Note: this follows same algorithm and Selenium and other similar tools as opposed to javascript .click() method that does not always work as expected
' Warning: if click opens a new window/tab as opposed to navigating within same page then may return before new page has finished loading - must switch to new session and then wait
Public Sub Click(Optional nodeId As Long = 0, Optional backendNodeId As Long = 0, Optional strategy As NavigationStrategy = NavigationStrategy.None, Optional button As String = "left", Optional clickCount As Integer = 1)
    ' reset flag for window open event
    If Not curSession Is Nothing Then curSession.windowOpenEventOccurred = False

    ' first make sure node is visible to click
    Focus nodeId, backendNodeId
    
    ' then get its center x,y to click
    Dim x As Double, y As Double
    getNodeCenter x, y, nodeId:=nodeId, backendNodeId:=backendNodeId
    If cdp.ErrorCode = 0 Then
        ' we need to ensure Page events are enabled so we can get notified of load complete or dom complete events
        If strategy = NavigationStrategy.Normal Or strategy = NavigationStrategy.eager Then cdp.Page.enable
        ' we need to ensure Network events are enabled so we can see when network activity has subsided
        If strategy = NavigationStrategy.NetworkIdle Then cdp.Network.enable
    
        ' now simulate click via mouse button press down and up
        cdp.SimulateInput.dispatchMouseEvent "mousePressed", x, y, button:=button, clickCount:=clickCount
        cdp.SimulateInput.dispatchMouseEvent "mouseReleased", x, y, button:=button, clickCount:=clickCount

        ' if click causes navigation, optionally wait for page loading event
        waitForPageToLoad strategy
    End If
End Sub

' given a DOM node as a cdpDomNode object, returns true if a node has corresponding attribute
' does not include its value, see getNodeAttribute to retrieve value
Public Function hasNodeAttribute(ByVal node As cdpDOMNode, ByVal attrib As String) As Boolean
    Dim i As Integer
    attrib = Trim$(attrib)
    With node
        For i = 1 To .attributes.count Step 2
            Dim key As String
            key = .attributes(i)
            If attrib = key Then
                hasNodeAttribute = True
                Exit For
            End If
        Next i
    End With
End Function

' given a DOM node as a cdpDomNode object, returns value of corresponding attribute's value or empty string
' will return an empty string if attribute not found or its value is "", see hasNodeAttribute to determine if node has attribute but set to ""
Public Function getNodeAttribute(ByVal node As cdpDOMNode, ByVal attrib As String) As String
    Dim i As Integer
    attrib = Trim$(attrib)
    With node
        For i = 1 To .attributes.count Step 2
            Dim key As String
            key = .attributes(i)
            If attrib = key Then
                getNodeAttribute = .attributes(i + 1)
                Exit For
            End If
        Next i
    End With
End Function

' given a DOM node, returns its id if has one, otherwise its name if has one, otherwise empty string
Public Function getNodeIdOrName(ByVal node As cdpDOMNode) As String
    Dim idOrName As String
    idOrName = getNodeAttribute(node, "id")
    If idOrName = vbNullString Then idOrName = getNodeAttribute(node, "name")
    getNodeIdOrName = idOrName
End Function

' takes a nodeId or backendNodeId (or RemoteObject.objectId) and returns cdpDOMNode object
Public Function getDomNode(Optional ByVal nodeId As Long = 0, Optional ByVal backendNodeId As Long = 0, Optional ByVal objectId As String = vbNullString, _
                           Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False) As cdpDOMNode
    Dim dict As Dictionary
    If objectId <> vbNullString Then
        Set dict = cdp.DOM.describeNode(objectId:=objectId, depth:=depth, pierce:=pierce)
    ElseIf backendNodeId <> 0 Then
        Set dict = cdp.DOM.describeNode(backendNodeId:=backendNodeId, depth:=depth, pierce:=pierce)
    Else ' assume nodeId is valid
        Set dict = cdp.DOM.describeNode(nodeId:=nodeId, depth:=depth, pierce:=pierce)
    End If
    If cdp.ErrorCode = 0 Then
        Dim node As cdpDOMNode
        Set node = New cdpDOMNode
        node.init dict, Me, nodeId
        Set getDomNode = node
    End If
End Function

' converts a Runtime.RemoteObject into a DOM.Node
Public Function RemoteObjectToDomNode(ByVal remoteObj As cdpRuntimeRemoteObject, Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False) As cdpDOMNode
    Set RemoteObjectToDomNode = getDomNode(objectId:=remoteObj.objectId, depth:=depth, pierce:=pierce)
End Function

' takes a Collection of nodeIds and returns a Dictionary of cdpDomNode objetcs
' Note: resulting dictionary is keyed on id if node has one, else name if has one, else a unique numerical value
Public Function getDomNodes(Optional ByVal nodeIds As Collection, Optional ByVal backendNodeIds As Collection, Optional objectIds As Collection, Optional ByVal remoteNodeList As cdpRuntimeRemoteObject, _
                            Optional depth As Long = -1, Optional ByVal pierce As Boolean = False) As Dictionary
    Dim nodes As Dictionary: Set nodes = New Dictionary
    Dim node As cdpDOMNode
    Dim v As Variant, key As Variant
    
    If Not objectIds Is Nothing Then
        For Each v In objectIds
            Set node = getDomNode(objectId:=v, depth:=depth, pierce:=pierce)
            key = getNodeIdOrName(node)
            If key = vbNullString Or nodes.Exists(key) Then key = nodes.count
            Set nodes(key) = node
            Set node = Nothing
        Next v
    ElseIf Not backendNodeIds Is Nothing Then
        For Each v In backendNodeIds
            Set node = getDomNode(backendNodeId:=v, depth:=depth, pierce:=pierce)
            key = getNodeIdOrName(node)
            If key = vbNullString Or nodes.Exists(key) Then key = nodes.count
            Set nodes(key) = node
            Set node = Nothing
        Next v
    ElseIf Not nodeIds Is Nothing Then
        For Each v In nodeIds
            Set node = getDomNode(nodeId:=v, depth:=depth, pierce:=pierce)
            key = getNodeIdOrName(node)
            If key = vbNullString Or nodes.Exists(key) Then key = nodes.count
            Set nodes(key) = node
            Set node = Nothing
        Next v
    ElseIf Not remoteNodeList Is Nothing Then
            ' using GetProperties we can actually retrieve the values in the NodeList
            Dim nodeList As Dictionary, result As Collection
            Set nodeList = cdp.Runtime.GetProperties(remoteNodeList.objectId, ownProperties:=True) ' we just want the nodes so skip all properties that are inherited
            Set result = nodeList("result")
            For Each v In result
                Dim propDesc As cdpRuntimePropertyDescriptor
                Set propDesc = New cdpRuntimePropertyDescriptor
                propDesc.init v
                If Not propDesc.Value Is Nothing Then
                    If propDesc.Value.subtype = "node" Then
                        Set node = getDomNode(objectId:=propDesc.Value.objectId, depth:=depth, pierce:=pierce)
                        key = getNodeIdOrName(node)
                        If key = vbNullString Or nodes.Exists(key) Then key = nodes.count
                        Set nodes(key) = node
                        Set node = Nothing
                    End If
                End If
                Set propDesc = Nothing
            Next v
            Set result = Nothing
            Set nodeList = Nothing
    Else
        Debug.Print "getDomNodes - No source of DOM nodes provided!"
        Stop
    End If
    
    Set getDomNodes = nodes
    Exit Function
ErrHandler:
    Debug.Print "getDomNodes() - " & Err.description
    Stop
End Function

' internally to support getting elements in iframes, we have to run
' javascript relative to the frame, for nonframe elements we can just call cdp.DOM.querySelector
' This compiles a javascript function that does that.
' Returns a Dictionary of ScriptId (string)
Private Function compileSupportRoutines() As Dictionary
    Dim compiledScripts As Dictionary: Set compiledScripts = New Dictionary
    Dim results As Dictionary, scriptId As String, ex As cdpRuntimeExceptionDetails
    
    Const querySelectorFn = "function __querySelectorFn(queryString, element) { if (element) return element.querySelector(queryString); else return document.querySelector(queryString); }"
    Const querySelectorAllFn = "return false;"
    
    Set results = cdp.Runtime.compileScript(querySelectorFn, vbNullString, persistScript:=True) ', executionContextId:=contextId)
    If results.Exists("scriptId") Then
        scriptId = results("scriptId")
        compiledScripts("querySelectorFn") = scriptId
    ElseIf results.Exists("exceptionDetails") Then
        Set ex = New cdpRuntimeExceptionDetails
        ex.init results("exceptionDetails")
        Debug.Print ex.exception
    End If
    
    Set results = cdp.Runtime.compileScript(querySelectorAllFn, vbNullString, persistScript:=True) ', executionContextId:=contextId)
    If results.Exists("scriptId") Then
        scriptId = results("scriptId")
        compiledScripts("querySelectorAllFn") = scriptId
    ElseIf results.Exists("exceptionDetails") Then
        Set ex = New cdpRuntimeExceptionDetails
        ex.init results("exceptionDetails")
        Debug.Print ex.exception
    End If
    
    Set compileSupportRoutines = compiledScripts
End Function

' returns the frameId matching a corresponding url [fragment] that can be passed to QuerySelector and QuerySelectorAll
Public Function GetFrameIdByUrl(ByVal url As String) As String
    ' attempt to find frame with matching partial url
    Dim v As Variant, frameId As String, frame As cdpPageFrame
    
    ' 1st cycle through looking for an exact match
    For Each v In curSession.frames.items
        Set frame = v
        If frame.url = url Then
            frameId = frame.id
            Exit For
        End If
    Next v
    
    ' if no match found, then look for partial match
    If frameId = vbNullString Then
        For Each v In curSession.frames.items
            Set frame = v
            If InStr(1, frame.url, url, vbTextCompare) > 0 Then
                frameId = frame.id
                Exit For
            End If
        Next v
    End If

    GetFrameIdByUrl = frameId
End Function

' convenience function that maps frame name to id and then calls QuerySelector with frameId
' Note: if frameName is not provided then frameid of root node is used, if rootId is not provided then a frame url fragment must be provided, otherwise nothing is returned
Public Function QuerySelectorOnFrame(ByVal selector As String, Optional ByVal frameName As String, Optional ByVal rootId As Long = 0, Optional ByVal url As String = vbNullString, _
                                     Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False, _
                                     Optional waitUntilFound As Boolean = True, Optional waitUntilVisible As Boolean = True) As cdpDOMNode
    If curSession.frameNames.Exists(frameName) Then
        Set QuerySelectorOnFrame = querySelector(selector, rootId:=rootId, frameId:=curSession.frameNames(frameName), depth:=depth, pierce:=pierce, waitUntilFound:=waitUntilFound, waitUntilVisible:=waitUntilVisible)
    ElseIf rootId <> 0 Then
        ' use frame associated with rootId object
        Dim node As cdpDOMNode
        Set node = getDomNode(nodeId:=rootId)
        Set QuerySelectorOnFrame = querySelector(selector, rootId:=rootId, frameId:=node.frameId, depth:=depth, pierce:=pierce, waitUntilFound:=waitUntilFound, waitUntilVisible:=waitUntilVisible)
    ElseIf url <> vbNullString Then
        Dim frameId As String
        frameId = GetFrameIdByUrl(url)
        ' if we found a matching frame, then run query
        If frameId <> vbNullString Then
            Set QuerySelectorOnFrame = querySelector(selector, rootId:=rootId, frameId:=frameId, depth:=depth, pierce:=pierce, waitUntilFound:=waitUntilFound, waitUntilVisible:=waitUntilVisible)
        End If
    End If
End Function

' if frameId = "" then assumes executing in main document context
' rootId is unused if frameId is specified
Public Function querySelector(ByVal selector As String, Optional ByVal rootId As Long = 0, Optional ByVal frameId As String = vbNullString, _
                              Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False, _
                              Optional waitUntilFound As Boolean = True, Optional waitUntilVisible As Boolean = True) As cdpDOMNode
    On Error GoTo ErrHandler
    Dim maxRetries As Long: maxRetries = 60 * 5 ' we sleep for second after each retry, don't try longer than 5 minutes
    
    Do
    Dim nodeId As Long, docRootId As Long, dict As Dictionary
    If frameId = vbNullString Then
        ' if no base element provided, use the document root of current target
        ' Note: we store in docRootId instead of updating rootId as if we need to retry, document may not be fully loaded so we may need to reget Document()
        If rootId < 1 Then
            Dim doc As cdpDOMNode
            Set doc = Document()
            docRootId = doc.nodeId
        Else
            docRootId = rootId
        End If
        ' using either provided nodeId or document nodeId, run querySelector; but return Nothing if no base element
        If docRootId > 0 Then
            nodeId = cdp.DOM.querySelector(docRootId, selector)
            If nodeId > 0 Then Set querySelector = getDomNode(nodeId, depth:=depth, pierce:=pierce)
        End If
    Else
        ' get unique(Excution)Id associated with frame
        If curSession.frameExecCtxMapping.Exists("$" & frameId) Then
            Dim contextId As String
            contextId = curSession.frameExecCtxMapping.Item("$" & frameId)
            Set dict = jsEval("document.querySelector('" & selector & "');", silent:=True, uniqueContextId:=contextId)
            If cdp.ErrorCode = 0 Then
                Dim remoteObj As cdpRuntimeRemoteObject
                Set remoteObj = New cdpRuntimeRemoteObject
                remoteObj.init dict
                If remoteObj.objectId <> vbNullString Then
                    Set querySelector = getDomNode(objectId:=remoteObj.objectId, depth:=depth, pierce:=pierce)
                    'Set QuerySelector = getDomNode(objectId:=dict("objectId"), depth:=depth, pierce:=pierce)
                    If querySelector.nodeId = 0 Then
                        querySelector.nodeId = cdp.DOM.requestNode(remoteObj.objectId)
                    End If
                Else
                    If IsNull(remoteObj.value) And (Not waitUntilFound) Then Debug.Print "Selector '" & selector & "' di not match any elements!"
                End If
            End If
        End If
    End If
    If waitUntilFound And (querySelector Is Nothing) Then
        maxRetries = maxRetries - 1
        Debug.Print "Try " & (60 * 5) - maxRetries
        Sleep 1
    End If
    Loop While (querySelector Is Nothing) And waitUntilFound And (maxRetries > 0)
    
    ' selector has to be found to check for visibility, so waitUntilVisible ignored if waitUntilFound is false
    If waitUntilFound And waitUntilVisible Then
        Do While Not querySelector.isVisible
            Sleep 1
        Loop
    End If
    
    Exit Function
errHandler:
    Debug.Print "Browser.querySelector - " & Err.description
    Stop
    'Resume
End Function


' convenience function to convert "abc def" class list into CSS format compatible with querySelectorAll
' classList is a space separated list of class names e.g. "classname1 classname2"
' or optionally an element type with class list in in square brackets, e.g. "input[classname1 classname2]"
' output format: element[class~='classname1'][class~='classname2'][...] (which is equivalent to element.classname1.classname2....)
Public Function getNodesByClassName(ByVal classList As String, Optional ByVal rootId As Long = 0, Optional ByVal frameId As String = vbNullString, _
                                 Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False) As Dictionary ' of cdpDomNode
    Dim selector As String, ndx As Integer
    If InStr(1, classList, "[", vbBinaryCompare) > 0 Then
        ' extract element names and class list
        ndx = InStr(1, classList, "[", vbBinaryCompare)
        If ndx > 1 Then selector = Trim$(left(classList, ndx - 1))
        classList = Mid$(classList, ndx + 1)
        ndx = InStr(1, classList, "]", vbBinaryCompare)
        If ndx > 0 Then classList = left(classList, ndx - 1)
    End If
    
    classList = Trim$(classList)
    If classList = vbNullString Then
        selector = "*"
    Else
        Dim classes() As String
        classes = Split(classList, " ", Compare:=vbBinaryCompare)
        For ndx = LBound(classes) To UBound(classes)
            Dim className As String
            className = Trim$(classes(ndx))
            If className <> vbNullString Then selector = selector & "[class~='" & className & "']"
        Next ndx
    End If
    
    Set getNodesByClassName = querySelectorAll(selector, rootId, frameId, depth, pierce)
End Function

' convenience function that maps frame name to id and then calls QuerySelectorAll with frameId
' Note: if frameName is not provided then frameid of root node is used, if rootId is not provided then a frame url fragment must be provided, otherwise nothing is returned
Public Function QuerySelectorAllOnFrame(ByVal selector As String, Optional ByVal frameName As String, Optional ByVal rootId As Long = 0, Optional ByVal url As String = vbNullString, _
                                     Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False) As cdpDOMNode
    If curSession.frameNames.Exists(frameName) Then
        Set QuerySelectorAllOnFrame = querySelectorAll(selector, rootId:=rootId, frameId:=curSession.frameNames(frameName), Depth:=Depth, pierce:=pierce)
    ElseIf rootId <> 0 Then
        ' use frame associated with rootId object
        Dim node As cdpDOMNode
        Set node = getDomNode(nodeId:=rootId)
        Set QuerySelectorAllOnFrame = querySelectorAll(selector, rootId:=rootId, frameId:=node.frameId, Depth:=Depth, pierce:=pierce)
    ElseIf url <> vbNullString Then
        Dim frameId As String
        frameId = GetFrameIdByUrl(url)
        ' if we found a matching frame, then run query
        If frameId <> vbNullString Then
            Set QuerySelectorAllOnFrame = querySelectorAll(selector, rootId:=rootId, frameId:=frameId, Depth:=Depth, pierce:=pierce)
        End If
    End If
End Function

' if frameId = "" then assumes executing in main document context
' returns a Dictionary of cdpDomNodes matching selector, see getDomNodes
' Note: resulting Dictionary will be empty (0 items) if an error or no matching elements found
' see: https://www.w3.org/TR/CSS21/selector.html%23id-selectors for a quick overview of selector formatting
Public Function querySelectorAll(ByVal selector As String, Optional ByVal rootId As Long = 0, Optional ByVal frameId As String = vbNullString, _
                                 Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False) As Dictionary ' of cdpDomNode
    On Error GoTo ErrHandler
    Set querySelectorAll = New Dictionary
    Dim nodeIds As Collection, dict As Dictionary
    Dim remoteObj As cdpRuntimeRemoteObject
    If frameId = vbNullString Then
        ' if no base element provided, use the document root of current target
        If rootId < 1 Then
            Dim doc As cdpDOMNode
            Set doc = Document()
            rootId = doc.nodeId
        End If
        ' using either provided nodeId or document nodeId, run querySelectorAll; but return default empty Dictionary if no base element
        If rootId > 0 Then
            Set nodeIds = cdp.DOM.querySelectorAll(rootId, selector)
            If Not nodeIds Is Nothing Then Set querySelectorAll = getDomNodes(nodeIds, depth:=depth, pierce:=pierce)
        End If
    Else
        ' get unique(Excution)Id associated with frame
        If curSession.frameExecCtxMapping.Exists("$" & frameId) Then
            Dim contextId As String
            contextId = curSession.frameExecCtxMapping.Item("$" & frameId)
            ' we need returnByValue to be unset or False to ensure we get the RemoteObject id not a bunch of empty objects
            Set dict = jsEval("document.querySelectorAll('" & selector & "');", silent:=True, uniqueContextId:=contextId, returnByValue:=False)
            If cdp.ErrorCode = 0 Then
                Set remoteObj = New cdpRuntimeRemoteObject
                remoteObj.init dict
                Set querySelectorAll = getDomNodes(remoteNodeList:=remoteObj, depth:=depth, pierce:=pierce)
            End If
        End If
    End If
    Exit Function
errHandler:
    Debug.Print "AutomateBrowser.querySelectorAll - " & Err.description
    Stop
    Resume
End Function

' given a set of nodes (Dictionary of cdpDOMNode objects), returns first node found with matching elementValue
' returns Nothing if no matching node found
Public Function GetNodeWithValue(ByRef nodes As Dictionary, ByVal value As String) As cdpDOMNode
    On Error GoTo ErrHandler
    Dim v As Variant, node As cdpDOMNode
    For Each v In nodes.items
        Set node = v
        If Trim(node.elementValue) = value Then
            ' we found our matching node, so return it
            Set GetNodeWithValue = node
            Exit For
        End If
    Next v
    Exit Function
ErrHandler:
    Debug.Print "GetNodeWithValue() - Error: " & Err.description
    Stop
End Function

' performs a querySelectorAll using specified selector and returns first node whose value matches label
Public Function QueryNodeWithValue(ByVal selector As String, ByVal value As String) As cdpDOMNode
    Dim nodes As Dictionary
    Set nodes = querySelectorAll(selector)
    Set QueryNodeWithValue = GetNodeWithValue(nodes, value)
End Function


' returns document root and all child elements of main frame, optinally piercing frames and shadow DOMs
Public Function Document(Optional ByVal depth As Long = 0, Optional ByVal pierce As Boolean = False) As cdpDOMNode
    If curSession.rootNode Is Nothing Then
        Dim dict As Dictionary
        Set dict = cdp.DOM.getDocument(Depth, pierce)
        If Not dict Is Nothing Then
            Set curSession.rootNode = New cdpDOMNode
            curSession.rootNode.init dict, Me
        End If
    End If
    Set Document = curSession.rootNode
End Function

' returns Dictionary of cdpRuntimePropertyDescriptor for given node
Public Function GetProperties(ByVal node As cdpDOMNode) As Dictionary
    Dim props As Dictionary
    Set props = New Dictionary
    If Not node.resolveNode Is Nothing Then
        Dim dict As Dictionary
        Set dict = cdp.Runtime.GetProperties(node.resolveNode.objectId)
        If cdp.ErrorCode = 0 Then
            Dim v As Variant
            If dict.Exists("result") Then
                Dim propArray As Collection
                Set propArray = dict("result")
                For Each v In propArray
                    Dim prop As cdpRuntimePropertyDescriptor
                    Set prop = New cdpRuntimePropertyDescriptor
                    prop.init v
                    props.Add prop.name, prop
                Next v
            End If
        End If
    End If
    Set GetProperties = props
End Function
