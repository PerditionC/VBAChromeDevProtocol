VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Basic implementation of Chrome Devtools Protocol (CDP)
Option Explicit

' public interface to access methods in supported Chrome Developer Protocol domains
' see https://chromedevtools.github.io/devtools-protocol/ for further documentation
' Note: the protocol supported by the VBA objects includes experimental components and may be a superset of what is actually supported by the browser
' !!! Edge returns tip of tree protocol JSON but claims only support for version 1.3, refer to documentation an limit to 1.3 compatible commands for best compatibility
Public Accessibility As cdpAccessibility
Public Browser As cdpBrowser
Public CSS As cdpCSS
Public DOM As cdpDOM
Public SimulateInput As cdpInput
Public Network As cdpNetwork
Public Overlay As cdpOverlay
Public Page As cdpPage
Public Runtime As cdpRuntime
Public Target As cdpTarget


' error information is set after results of each call, code is set to 0
' and message is empty string on success; data may be blank even on error
Public ErrorCode As Integer
Public ErrorMessage As String
Public ErrorData As Variant

' current session attached to direct message to
' Note: sendMessage automatically adds to each message sent to browser
'       set to vbNullString to send avoid adding sessionId to message (sends to browser sessionless target)
Public sessionId As String

' which browser did we connect to
Public browserProgram As browserType

' link to browser for sending and receiving protocol messages, only 1 of these (objBrowser, wsBrowser) should be valid
Private objBrowser As clsProcess
Private wsBrowser As clsWebSocket

' every message sent over Chrome Developer Protocol has an id, this is id of last message we sent
Private lngLastID As Long

'this buffer holds messages that are not yet processed
Private strBuffer As String

' maintains all messages received between last message sent and response to current message sent, usually these are events
Private objAllMessages As Dictionary

' registered event handlers
' key is event name, only 1 event handler allowed per event name
' where the item is an Object that implments
' Public Sub processEvent(ByVal eventName As String, ByVal eventData As Dictionary)
' Note: the same object can process multiple events, so event name is provided
Private eventHandlers As Dictionary


' CDP messages received from chrome are null-terminated
' It seemed to me you cant search for vbnull in a string
' in vba. Thats why i re-implemented the search function
Private Function searchNull() As Long
    Dim i As Long
    
    Dim lngBufferLength As Long
    lngBufferLength = Len(strBuffer)
    searchNull = 0
    
    If lngBufferLength = 0 Then
        Exit Function
    End If
    
    For i = 1 To lngBufferLength
        If Mid(strBuffer, i, 1) = vbNullChar Then
            searchNull = i
            Exit Function
        End If
    Next i
End Function


' read in any pending data from connected browser pipe/socket and append to buffer
' nonblocking, will return immediately if nothing new to add
Private Sub readRawMessageData()
    Dim errorText As String, strResult As String
    
    If Not objBrowser Is Nothing Then
        strResult = objBrowser.GetMessageData(errorText)
    ElseIf Not wsBrowser Is Nothing Then
        strResult = wsBrowser.GetMessageUTF8()
        errorText = wsBrowser.errorText
    Else
        Debug.Print "readRawMessageData() - Error: no implementation to communicate with browser available!"
        Stop
    End If
    
    If errorText = "None" Then
        strBuffer = strBuffer & strResult
    End If
End Sub

' store and retrieve messages from a queue event processing or delayed response handling
Private Sub pushMessageToQueue(ByRef objRes As Dictionary)
    objAllMessages.Add CStr(objAllMessages.count), objRes
End Sub

' returns Nothing if no more messages, otherwise returns the most recent message received
Public Function popMessageFromQueue() As Dictionary
    On Error GoTo ErrHandler
    Dim key As String
    key = CStr(objAllMessages.count - 1) ' last item added key is one less than current count, 0 based
    If objAllMessages.Exists(key) Then
        Set popMessageFromQueue = objAllMessages(key)
        objAllMessages.Remove key
    End If
    
    Exit Function
ErrHandler:
    Debug.Print "ERROR: popMessageFromQueue() - " & Err.description
    Err.Raise Err.Number, Err.source, "VBAChromeDevProtocol.clsCDP.popMessageFromQueue: " & Err.description, Err.HelpFile, Err.HelpContext
End Function

' clears any existing queued messages (events and responses)
Public Sub clearMessageQueue()
    Set objAllMessages = New Dictionary
    'strBuffer = vbNullString
End Sub

' returns all currently queued messages
' Note: peakMessage will continue to add new messages
'       sendMessage clears the internal queue, so getMessageQueue must be called again
'       after a sendMessage to get any new messages (i.e. the new queue Dictionary)
Public Function getMessageQueue() As Dictionary
    Set getMessageQueue = objAllMessages
End Function

' looks for any new incoming messages
' event handlers are called and if returns true event not queued
' queues and returns any message found
Public Function peakMessage() As Dictionary
    On Error GoTo ErrHandler
    Dim lngNullCharPos As Long
    
    ' get any new data if available since last call
    readRawMessageData
        
    ' process message if found
    lngNullCharPos = searchNull()
    If lngNullCharPos > 0 Then
        Dim strRes As String
        ' get next message from buffer
        strRes = left(strBuffer, lngNullCharPos - 1)
        ' remove message from the buffer
        strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
                        
        Dim objRes As Dictionary
        'Debug.Print (strRes)
            
        ' is there a message to process?
        If strRes <> "" Then
            ' convert to easier to use form (i.e. convert raw string to json object)
            On Error Resume Next
            Set objRes = JsonConverter.ParseJson(strRes)
            If Err.number <> 0 Then
                strRes = vbNullString
                GoTo messageHandled
            End If
            On Error GoTo ErrHandler
                                    
            ' see if this is an event (no "id") or response to a message (has "id" of message)
            ' process event or queue message
            If objRes.Exists("id") Then
                If objRes("id") = lngLastID Then
                    ' if expected reply recieved then return it as well as queue it
                    Set peakMessage = objRes
                End If
            ElseIf objRes.Exists("error") Then
                Set peakMessage = objRes
            Else 'If Not objRes.Exists("id") Then
                ' for events, if handler registered then run it and queue only if returns false, otherwise just queue it
                Dim eventName As String
                If objRes.Exists("method") Then
                    eventName = objRes("method")
                    Debug.Print "Event: " & eventName
                    If eventHandlers.Exists(eventName) Then
                        Dim eventHandler As Object
                        Set eventHandler = eventHandlers(eventName)
                        If eventHandler.processEvent(eventName, objRes) Then ' if returns True then don't queue event message, assume handled
                            GoTo messageHandled
                        End If
                    End If
                End If
            End If
queueMessage:
            'Debug.Print JsonConverter.ConvertToJson(objRes)  ' this will slow down processing if lots of messages or big messages, e.g. get properties
            pushMessageToQueue objRes
messageHandled:
        End If
    End If

    DoEvents
    Exit Function
ErrHandler:
    Debug.Print "peakMessage() - Error: " & Err.description
    Stop
    Resume
End Function

' sends a CDP message to browser
' by default waits for reply to message (returned message with same id as sent message)
' but if called with nowait=True then returns without waiting for reply - you must call peakMessage in a loop to receive reply
' peakMessage is called at least once to retrieve any immediately available data
Private Function sendMessage(ByVal strMessage As String, Optional ByVal nowait As Boolean = False) As Dictionary
    On Error GoTo ErrHandler
    ' lngLastId is id of current most recent message sent (it is incremented when the next message is sent)
    'We increase the global ID counter, this is the id of the message we are sending now
    lngLastID = lngLastID + 1
    
    ' Before sending a message the messagebuffer is emptied
    ' All messages that we have received sofar cannot be an answer
    ' to the message that we will send
    ' So they can be safely discarded
    ' Note: strBuffer is not cleared so we can still run event handler for any events not yet processed
    clearMessageQueue ' discard any messages not already processed

    ' sometimes edge writes to stdout
    ' we clear stdout here, too.
    Dim ignored As String
    'objBrowser.readProcSTD ignored

    ' We add the currentID and sessionID to the message (assume flat messages and sessionId required)
    strMessage = left(strMessage, Len(strMessage) - 1)
    
    If sessionId <> vbNullString Then
        strMessage = strMessage & ", ""sessionId"":""" & sessionId & """"
    End If
    
    strMessage = strMessage & ", ""id"":" & lngLastID & "}" & vbNullChar
    
    ' write message to browser
    Dim errorSending As Boolean
    If Not objBrowser Is Nothing Then
        errorSending = (objBrowser.writeProc(strMessage) = 0)
    ElseIf Not wsBrowser Is Nothing Then
        errorSending = Not wsBrowser.SendMessageUTF8(strMessage)
    Else
        errorSending = True
        Debug.Print "Warning objBrowser and wsBrowser are both nothing, unable to communicate with browser!"
    End If
    If errorSending Then
        Debug.Print "-----"
        Debug.Print "Failed to write CDP message!"
        Debug.Print strMessage
        Debug.Print "-----"
        Err.Raise Err.LastDllError
        Exit Function
    End If
        
    
    ' We have some failsafe timer in order not to loop forever
    Dim maxWait As Date: maxWait = DateAdd("s", 5 * 60, Now) ' 5 minutes
    
    ' We read from edge and then process messages until we receive a
    ' message with our ID or until timeout or because user requested not to wait
    Do While (Now <= maxWait) And (Not nowait) ' not timeout or no wait
        DoEvents
        
        Dim objRes As Dictionary
        Set objRes = peakMessage()
        
        ' did we get a response matching our request (has matching id)
        If Not objRes Is Nothing Then ' then objRes.Item("id") = lngCurrentId
            popMessageFromQueue
            Set sendMessage = objRes
            Exit Function
        End If
    Loop
    
    If Not nowait Then
        Debug.Print "-----"
        Debug.Print "Timeout waiting for response to message " & lngLastID
        Debug.Print strMessage
        Debug.Print "-----"
    End If
    
    Exit Function
ErrHandler:
    Debug.Print "ERROR: sendMessage(" & strMessage & ", nowait=" & nowait & ") - " & Err.description
    Err.Raise Err.Number, Err.source, "VBAChromeDevProtocol.clsCDP.sendMessage: " & Err.description, Err.HelpFile, Err.HelpContext
End Function


' check for error, storing for later use and displays information to console (immediate window)
Private Function ProcessErrorMessage(result As Dictionary) As Boolean
    If result Is Nothing Then
        ErrorCode = -1 'arbitrary nonzero value, may change in future
        ErrorMessage = "Timeout or No results"
        ErrorData = vbEmpty
        Debug.Print "Error: [" & ErrorCode & "] " & ErrorMessage
        ProcessErrorMessage = True
    ElseIf result.Exists("error") Then
        Dim errorDetails As Dictionary
        Set errorDetails = result("error")
        ErrorCode = errorDetails("code")
        ErrorMessage = errorDetails("message")
        Debug.Print "Error: [" & ErrorCode & "] " & ErrorMessage
        If errorDetails.Exists("data") Then
            ErrorData = errorDetails("data")
            Debug.Print "       " & ErrorData
        Else
            ErrorData = vbEmpty ' vbNullString? is errorDetails("data") a string or object?
        End If
        ProcessErrorMessage = True
    Else
        ErrorCode = 0
        errorMessage = vbNullString
        ErrorData = vbEmpty
    End If
End Function


' convenience function to invoke a given method with JSON encoded parameters
Function InvokeMethod(methodName As String, params As Dictionary) As Dictionary
    Dim paramStr As String: paramStr = ConvertToJson(params)
    
    Set InvokeMethod = sendMessage("{""method"":""" & methodName & """,""params"":" & paramStr & "}")
    
    ' automatically display error if one returned, sets ErrorCode and related values
    ' exit early on error, either explicit InvokeMethod.Exists("error") or InvokeMethod Is Nothing
    If ProcessErrorMessage(InvokeMethod) Then Exit Function
    
    ' otherwise if successful, just return result object
    If InvokeMethod.Exists("result") Then
        Set InvokeMethod = InvokeMethod.Item("result")
    'Else return default value (Nothing)
    End If
End Function

Private Sub Class_Initialize()
    Set eventHandlers = New Dictionary

    Set Accessibility = New cdpAccessibility
    Accessibility.init Me
    Set Browser = New cdpBrowser
    Browser.init Me
    Set CSS = New cdpCSS
    CSS.init Me
    Set DOM = New cdpDOM
    DOM.init Me
    Set SimulateInput = New cdpInput
    SimulateInput.init Me
    Set Network = New cdpNetwork
    Network.init Me
    Set Overlay = New cdpOverlay
    Overlay.init Me
    Set Page = New cdpPage
    Page.init Me
    Set Runtime = New cdpRuntime
    Runtime.init Me
    Set Target = New cdpTarget
    Target.init Me
    
    ' avoid issue where uniqueContextId and other ids mistaken for large numbers instead of strings (may cause truncation of some double values)
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
End Sub

Private Sub Class_Terminate()
    Set eventHandlers = Nothing
    
    ' ensure we don't have cyclic dependencies
    Set Accessibility = Nothing
    Set Browser = Nothing
    Set CSS = Nothing
    Set DOM = Nothing
    Set SimulateInput = Nothing
    Set Network = Nothing
    Set Overlay = Nothing
    Set Page = Nothing
    Set Runtime = Nothing
    Set Target = Nothing
End Sub


' internal event handler
' Note: if registering any event handled here, caller must be sure to call to our event handler!
Public Function processEvent(ByVal eventName As String, ByVal eventData As Dictionary) As Boolean
    ' none currently registered
    Select Case eventName
        Case "Domain.eventName"
        Case Else
    End Select
End Function


' ********* Public interface of clsCDP **********************
' launch(...) must be called before any other methods
' all other public methods on clsCDP are convenience functions


' registers an object to handle events as received as opposed to processing message queue after response
' e.g. for a download progress bar while download is running
' returns current handler if one exists, set to Nothing to remove handler
Public Function registerEventHandler(ByVal eventName As String, eventHandler As Object) As Object
    ' get current handler to return
    If eventHandlers.Exists(eventName) Then Set registerEventHandler = eventHandlers(eventName)
    ' see if removing handler or installing new one
    If eventHandler Is Nothing Then
        If eventHandlers.Exists(eventName) Then eventHandlers.Remove eventName
    Else
        Set eventHandlers(eventName) = eventHandler
    End If
End Function


' This method starts up the browser
' It will attempt to terminate the browser if found to already be running (user is prompted prior to abort)
' If autoAttach is True then after connection to browser established will initiate attach call automatically
' if useWebSocket is True then will connect to browser websocket on localhost port 9222, otherwise connects directly via a pipe
' if useWebSocket is True and useExistingBrowser then does not (kill and) spawn browser before connecting, ignored if useWebSocket is false
' Warning! if autoAttach is True then the url must exactly match and may fail if browser changes expected url unless partialMatch is True
' Returns True if successfully initialized browser, False otherwise
Public Function launch( _
    Optional url As String = vbNullString, _
    Optional autoAttach As Boolean = True, _
    Optional partialMatch As Boolean = True, _
    Optional useWebSocket As Boolean = False, _
    Optional useExistingBrowser As Boolean = False, _
    Optional whichBrowser As browserType = browserType.AnyBrowser, _
    Optional byVal additionnalInlineCommands as string = vbNullString, _
    Optional byVal killWithoutAsking as boolean = False, _
    Optional reconnectUrl As String = vbNullString _
    ) As Boolean
    On Error GoTo ErrHandler
    
    browserProgram = whichBrowser ' specify requested browser, LaunchBrowser sets to browser we end up using
    Dim o As Object
    Set o = LaunchBrowser(browserProgram, url, useWebSocket, useExistingBrowser, additionnalInlineCommands, killWithoutAsking)
    If o Is Nothing Then Exit Function ' error starting browser
    If TypeName(o) = "clsProcess" Then
        Set objBrowser = o
    ElseIf TypeName(o) = "clsWebSocket" Then
        Set wsBrowser = o
    End If
    
    ' initialize message id used to track message responses
    lngLastID = 1
    
    ' if connected via pipe we need to clear pipe from any data written during startup
    Dim strRes As String
    Dim intRes As Integer: intRes = 0
    Dim intCounter As Integer: intCounter = 0
    
    Do Until intRes > 0 Or intCounter > 1000 Or useWebSocket
        intRes = objBrowser.readProcSTD(strRes)
        DoEvents
        Call Sleep(0.1)
        intCounter = intCounter + 1
    Loop
    
    ' automatically attach to requested page or any page if no specific one requested
    ' WARNING if we didn't spawn browser then url may not match any page Target!
    If autoAttach Then
        If Me.attach(url, partialMatch:=partialMatch, altUrl:=reconnectUrl) = vbNullString Then
            If useWebSocket Then wsBrowser.Disconnect
            Exit Function ' failed to attach on launch
        End If
    End If
    
    ' assume success
    launch = True
    Exit Function
ErrHandler:
    Debug.Print "cldCDP.launch() - Error: " & Err.description
    Stop
    Resume
End Function


' returns True if urls match
Private Function urlMatch(ByRef targetUrl As String, ByRef strUrl As String, ByVal partialMatch As Boolean)
    If targetUrl = strUrl Then
foundMatch:
        urlMatch = True
    Else
        ' check if http: url was automatically upgraded to https:
        If (left(strUrl, Len("http:")) = "http:") And (left(targetUrl, Len("https:")) = "https:") Then
            ' if comparison ignoring http: and https: protocols match, then assume our target
            If Mid(targetUrl, Len("https:") + 1) = Mid(strUrl, Len("http:") + 1) Then GoTo foundMatch
        End If
        
        ' if no protocol was provided, e.g. just "www.google.com" then do match ignoring protocol
        If left(strUrl, 4) <> "http" Then
            If Mid(targetUrl, Len("http://") + 1) = strUrl Then GoTo foundMatch
            If Mid(targetUrl, Len("https://") + 1) = strUrl Then GoTo foundMatch
        End If
        
        ' finally if user requested partial matches to work, check for that
        If partialMatch Then
            If InStr(1, targetUrl, strUrl, vbTextCompare) > 0 Then GoTo foundMatch
        End If
    End If
End Function

' This function must be called after launch and before all other methods
' This function attaches to a session of the browser
' If targetId is not "" then attempts to attach to specified target, otherwise first determines targetId for url and then attaches
' strUrl is the page to match, by default looks for newtab page if left as ""
' if strUrl is provided and partial match is false, then may fail to attach if browser changes url (e.g. Google site may append "?gws_rd=ssl" to url)
' if partialMatch is true then attaches to first target whose url contains strUrl
' a second altUrl may also be specified, such as if reconnecting where the initial url may differ from an initial connection
' Returns the sessionId if successfully attached
Public Function attach(Optional ByVal strUrl As String = vbNullString, Optional ByVal targetId As String = vbNullString, Optional partialMatch As Boolean = False, Optional altUrl As String = vbNullString) As String
    If targetId <> vbNullString Then GoTo TargetId_Provided
    
    If strUrl = vbNullString Then
        If (browserProgram And browserType.Edge) = browserType.Edge Then
            strUrl = "edge://newtab/"
        ElseIf (browserProgram And browserType.Chrome) = browserType.Chrome Then
            strUrl = "chrome://newtab/"
        Else ' firefox or other browser
            strUrl = "about:blank"
        End If
    End If
    
    ' if url passed missing protocol and path (no / anywhere) and not a special page (e.g. about:blank) then assume path is "/"
    If (InStr(1, strUrl, "/", vbBinaryCompare) < 1) And (InStr(1, strUrl, ":", vbBinaryCompare) < 1) Then strUrl = strUrl & "/"
    
    Target.setDiscoverTargets discover:=True ' Note: nothing is returned, but causes objAllMessages to queue Target.targetCreated events
    
    Dim i As Integer
    Dim boolFound As Boolean
    
    Dim strKey As Variant
    
    Dim objRes As Dictionary
    Dim objDic2 As Dictionary
    Dim objDic3 As Dictionary
    
checkAgain:
    For Each strKey In objAllMessages.Keys
        Set objRes = objAllMessages.Item(strKey)
        
        If Not objRes.Exists("params") Then GoTo nextloop1
        Set objDic2 = objRes.Item("params")
        
        If Not objDic2.Exists("targetInfo") Then GoTo nextloop1
        Set objDic3 = objDic2.Item("targetInfo")
        
        If objDic3.Item("type") <> "page" And objDic3.Item("type") <> "iframe" Then GoTo nextloop1
        
        Dim targetUrl As String
        targetUrl = objDic3.Item("url")
        If urlMatch(targetUrl, strUrl, partialMatch) Then GoTo foundTarget
        If altUrl <> vbNullString Then If urlMatch(targetUrl, altUrl, partialMatch) Then GoTo foundTarget
        GoTo nextloop1
        
foundTarget:
        boolFound = True
        Exit For

nextloop1:
    Next strKey
    
    If Not boolFound Then
        ErrorCode = -1 ' arbitrary error
        ErrorMessage = "Failed to attach - unable to find target [url=" & strUrl & "]"
        Debug.Print "Error: [" & ErrorCode & "] " & ErrorMessage
        Exit Function
    End If

TargetId_Provided:

    If targetId = vbNullString Then targetId = objDic3.Item("targetId")
    attach = Target.attachToTarget(targetId, flatten:=True)
    If ErrorCode <> 0 Then Exit Function
    If sessionId = vbNullString Then sessionId = attach ' only update current session if none previously established

    ' leave enabled so we can monitor available targets
    'Target.setDiscoverTargets discover:=False
    
    ' enable events we normally want to listen to by default
    enableSessionEvents attach
End Function


' enables all the event messages we normally want to see, these are needed for AutomateBrowser to function correctly
Private Sub enableSessionEvents(ByVal sessionId As String)
        ' we need Runtime enabled so we can maintain a mapping of execution contexts and sessions
        ' disabling will likely cause events to be missed and this class to not function as expected
        Runtime.enable
        
        ' automatically turned on on wait to enable frame start/stop message, and ensure all frameid to name mappings
        ' if disabled then will not be able to access other than main frame nor wait for pages to finish loading
        Page.enable
        
        ' used to notify browser to track DOM nodes, i.e. ensure nodeId returned is nonzero
        ' must be enabled for Document cache to remain consistent and not stale
        DOM.enable
        
        ' used to track network connections
        ' must be enabled if wait for networkidle
        ' this causes too many events to skip on websites that have background connections, so we enable/disable on navigate/click and waitForPageToLoad explicitly
        'Network.enable
        
        ' automatically attach to related targets, avoids need to manually attach to target if new window opened
        ' can be enabled or disabled as desired, but if disabled then must manually attach to any new targets (new windows)
        ' so as to have a session to control it
        'Target.setAutoAttach autoAttach:=True, waitForDebuggerOnStart:=False
End Sub
