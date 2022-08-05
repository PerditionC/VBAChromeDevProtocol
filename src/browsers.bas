Attribute VB_Name = "browsers"
' provides information about supported browsers
Option Explicit

' supported browsers process name and installation sub path
Private Const processNameEdge As String = "msedge.exe"
Private Const browserExeSubPathEdge As String = "\Microsoft\Edge\Application\msedge.exe"

Private Const processNameChrome As String = "chrome.exe"
Private Const browserExeSubPathChrome As String = "\Google\Chrome\Application\chrome.exe"

Private Const processNameFirefox As String = "firefox.exe"
Private Const browserExeSubPathFirefox As String = "\Mozilla Firefox\firefox.exe"


' enumeration if browser is installed and whether running or not
Public Enum BrowserAvailability
    Not_Installed   ' browser not found
    Running         ' browser is installed but currently in use
    Available       ' browser is installed and not currently in use
End Enum


' returns availability of requested browser
Public Function browserStatus(ByVal oneBrowser As browserType) As BrowserAvailability
   browserStatus = BrowserAvailability.Available
   If getProcessPathAndName(oneBrowser) = vbNullString Then browserStatus = BrowserAvailability.Not_Installed: Exit Function
   If IsProcessRunning(ProcessName:=getProcessName(oneBrowser)) Then browserStatus = BrowserAvailability.Running: Exit Function
End Function

' returns True if requested browser is installed (does not check if currently in use)
Private Function browserIsInstalled(ByVal oneBrowser As browserType)
   browserIsInstalled = False
   If getProcessPathAndName(oneBrowser) <> vbNullString Then browserIsInstalled = True
End Function
     
' returns first browser from selection that is installed and not currently in use
' If either no requested browser is installed or all requested browsers are in use then returns noBrowser
Public Function getFirstAvailableBrowser(Optional ByVal browserOne As browserType = browserType.noBrowser, Optional ByVal browserTwo As browserType = browserType.noBrowser, Optional ByVal browserThree As browserType = browserType.noBrowser) As browserType
   getFirstAvailableBrowser = browserType.noBrowser
   If browserStatus(browserOne) = BrowserAvailability.Available Then getFirstAvailableBrowser = browserOne: Exit Function
   If browserStatus(browserTwo) = BrowserAvailability.Available Then getFirstAvailableBrowser = browserTwo: Exit Function
   If browserStatus(browserThree) = BrowserAvailability.Available Then getFirstAvailableBrowser = browserThree: Exit Function
End Function

' returns first browser from selection that is installed (regardless if in use or not)
' If no requested browser is installed then returns noBrowser
Public Function getFirstInstalledBrowser(Optional ByVal browserOne As browserType = browserType.noBrowser, Optional ByVal browserTwo As browserType = browserType.noBrowser, Optional ByVal browserThree As browserType = browserType.noBrowser) As browserType
   getFirstInstalledBrowser = browserType.noBrowser
   If browserIsInstalled(browserOne) Then getFirstInstalledBrowser = browserOne: Exit Function
   If browserIsInstalled(browserTwo) Then getFirstInstalledBrowser = browserTwo: Exit Function
   If browserIsInstalled(browserThree) Then getFirstInstalledBrowser = browserThree: Exit Function
End Function
                                 
' returns just the process name for matching browser
Private Function getProcessName(ByVal whichBrowser As browserType)
    If (whichBrowser And browserType.Edge) = browserType.Edge Then
        getProcessName = processNameEdge
    ElseIf (whichBrowser And browserType.Chrome) = browserType.Chrome Then
        getProcessName = processNameChrome
    ElseIf (whichBrowser And browserType.firefox) = browserType.firefox Then
        getProcessName = processNameFirefox
    Else
        Debug.Print "Unknown browserType!"
        Stop
    End If
End Function


'
' Return the file path of supported browser if browser is installed on computer
'
Private Function getProcessPathAndName(ByVal oneBrowser As browserType) As String
checkNextBrowser:
    getProcessPathAndName = vbNullString
    Dim browserExeSubPath As String
   
    browserExeSubPath = vbNullString
    ' Note: order of these checks is reverse of default match order, i.e. edge->chrome->firefox
    If (oneBrowser And browserType.firefox) = browserType.firefox Then browserExeSubPath = browserExeSubPathFirefox
    If (oneBrowser And browserType.Chrome) = browserType.Chrome Then browserExeSubPath = browserExeSubPathChrome
    If (oneBrowser And browserType.Edge) = browserType.Edge Then browserExeSubPath = browserExeSubPathEdge
    If browserExeSubPath = vbNullString Then Exit Function
   
    ' Get "program files (X86)" exact path  (sometimes there is a space before (X86) so get it with Environ$
  
    Dim programFiles As String
    Dim programFilesX86 As String
    programFilesX86 = Environ$("PROGRAMFILES(X86)")
    If Environ$("ProgramW6432") = vbNullString Then
        programFiles = Environ$("ProgramFiles")
    Else
        programFiles = Environ$("ProgramW6432")
    End If
    '
    ' Sometimes Firefox is installed in AppData/Local... (documented on internet but not tested...)
    Dim appDataLocal As String
    appDataLocal = Environ$("LocalAppData")
    
    ' note order of checks below, if browser somehow installed to multiple locations, last match is the one used
   
    ' verify if browser exists in Program Files
    If Dir(programFiles & browserExeSubPath, vbNormal) <> "" Then getProcessPathAndName = programFiles & browserExeSubPath: Exit Function
      
    ' verify if browser exists in Program Files (X86)
    If Dir(programFilesX86 & browserExeSubPath, vbNormal) <> "" Then getProcessPathAndName = programFilesX86 & browserExeSubPath: Exit Function
         
    ' verify if browser exists in AppData\Local
    If Dir(appDataLocal & browserExeSubPath, vbNormal) <> "" Then getProcessPathAndName = appDataLocal & browserExeSubPath: Exit Function

    ' if getProcessPathAndName is still "" then we didn't find browser installed, if multiple browsers requested then try others
    If getProcessPathAndName = vbNullString Then
        ' if either Chromium based browser was requested, try again with other one; also still check for firefox if requested
        If (oneBrowser And browserType.Chromium) = browserType.Chromium Then
            oneBrowser = oneBrowser And (Not browserType.Edge) ' assumes Edge checked prior to Chrome
            GoTo checkNextBrowser
        End If
        ' if any other browser and firefox also requested then assume Chromium browsers already failed, so now try firefox
        If ((oneBrowser And browserType.firefox) = browserType.firefox) And (oneBrowser <> browserType.firefox) Then
            oneBrowser = browserType.firefox
            GoTo checkNextBrowser
        End If
    End If

End Function

' To start a browser when no need to control
Public Function startBrowserNoControl(oneBrowser As browserType, Optional url As String = vbNullString) As Boolean
   Dim browserFile As String
   startBrowserNoControl = True
   browserFile = getProcessPathAndName(oneBrowser)
   If browserFile <> vbNullString Then Shell ("""" & browserFile & """ """ & url & """"): Exit Function
   startBrowserNoControl = False
End Function

' LaunchBrowser : validates if browser is installed and can be launched
'                 or ensure browser process is started and connected
'
' whichBrowser : the browser to connect to
'    Edge, Chrome, FireFox
'    Options to run first unused browser when not using WebSocket
'        AnyBrowser: Edge, Chrome, FireFox  if all running, kill Edge and launch Edge
'        Chromium: Edge, Chrome  if all running, kill Edge and launch Edge
'
' if a websocket is to be used then see if useexisting browser is requested, if
' not use same algorithm to determine browser as direct pipe case
' but if want to reuse we search for existing process and only if not found do we
' spawn, note AnyBrowser and Chromium will always spawn/use Edge first, otherwise specified browser
' Note: Edge is used first and then Chrome only because Chrome is generally in use by person
' and Firefox may not be installed along with its CDP support is not as mature
'
' at the end, whichBrowser is set to browser used for connection and connection object is returned
'
Public Function LaunchBrowser( _
    ByRef whichBrowser As browserType, _
    Optional url As String = vbNullString, _
    Optional useWebSocket As Boolean = False, _
    Optional useExistingBrowser As Boolean = False, _
    Optional additionalInlineCommands As String = vbNullString, _
    Optional killWithoutAsking As Boolean = False _
    ) As Object
    Dim browserConnectionObj As Object
    Dim objBrowser As clsProcess
    Dim wsBrowser As clsWebSocket
    
    ' check for browser(s) running
    Dim runningBrowsers As browserType
    If IsProcessRunning(ProcessName:=getProcessName(browserType.Edge)) Then runningBrowsers = runningBrowsers Or browserType.Edge
    If IsProcessRunning(ProcessName:=getProcessName(browserType.Chrome)) Then runningBrowsers = runningBrowsers Or browserType.Chrome
    If IsProcessRunning(ProcessName:=getProcessName(browserType.firefox)) Then runningBrowsers = runningBrowsers Or browserType.firefox
    
    ' select first browser within whichBrowser that is not running (if installed)
    If browserIsInstalled(browserType.Edge) And ((whichBrowser And browserType.Edge) = browserType.Edge) And ((useWebSocket And useExistingBrowser) Or ((runningBrowsers And browserType.Edge) <> browserType.Edge)) Then
        whichBrowser = browserType.Edge
    ElseIf browserIsInstalled(browserType.Chrome) And ((whichBrowser And browserType.Chrome) = browserType.Chrome) And ((useWebSocket And useExistingBrowser) Or ((runningBrowsers And browserType.Chrome) <> browserType.Chrome)) Then
        whichBrowser = browserType.Chrome
    ElseIf browserIsInstalled(browserType.firefox) And ((whichBrowser And browserType.firefox) = browserType.firefox) And ((useWebSocket And useExistingBrowser) Or ((runningBrowsers And browserType.firefox) <> browserType.firefox)) Then
        whichBrowser = browserType.firefox
    Else
        ' at this point all requested browsers are running so we need to identify first installed browser within whichBrowser
        If browserIsInstalled(browserType.Edge) And ((whichBrowser And browserType.Edge) = browserType.Edge) Then
            whichBrowser = browserType.Edge
        ElseIf browserIsInstalled(browserType.Chrome) And ((whichBrowser And browserType.Chrome) = browserType.Chrome) Then
            whichBrowser = browserType.Chrome
        ElseIf browserIsInstalled(browserType.firefox) And ((whichBrowser And browserType.firefox) = browserType.firefox) Then
            whichBrowser = browserType.firefox
        Else
            'At this point, whichBrowser contains no installed browsers
            Debug.Print "LaunchBrowser() - unable to determine browser to start!"
            MsgBox "Unable to determine browser to start! Aborting!", vbCritical Or vbOKOnly
            Exit Function
        End If
    End If
    
    ' At this point, whichBrowser contains the browser to start or to kill and restart
    ' Verification was made that whichBrowser is installed on the running computer
    
    
    ' ensure browser is not already running (kill it if it is)
    If Not (useWebSocket And useExistingBrowser) Then
        If Not TerminateProcess(ProcessName:=getProcessName(whichBrowser), PromptBefore:=Not killWithoutAsking) Then
            ' abort if browser is already running and failed to kill it or user elected not to terminate
            MsgBox "Failed to terminate and start expected browser! Aborting!", vbCritical Or vbOKOnly
            Exit Function
        End If
    End If

    Dim strCall As String
    ' --remote-debugging-pipe allow communicating with Edge
    ' --remote-debugging-port=9222 alternative for communication via TCP
    ' --enable-automation can be omitted for Edge to avoid displaying "being controlled banner", but necessary for Chrome
    ' --disable-infobars removes messages such as "being controlled by automation" banners, ignored by Edge, used by Chrome
    ' --enable-logging allows viewing additional connection details, but causes opening of console windows
    ' --user-data-dir=c:\temp\fakeEdgeUser to open in different profile, and therefore session than current user, but screws up because Admin managed user so don't use

    ' only one of objBrowser or wsBrowser can be valid, depending how we are connecting initialize the correct one
    If useWebSocket Then
        Set objBrowser = Nothing
        Set wsBrowser = New clsWebSocket
        Set LaunchBrowser = wsBrowser
        strCall = """" & getProcessPathAndName(whichBrowser) & """ --remote-debugging-port=9222 " & additionalInlineCommands & " " & url
        If (Not useExistingBrowser) Or (Not IsProcessRunning(ProcessName:=getProcessName(whichBrowser))) Then
            If Not SpawnProcess(strCall) Then
                MsgBox "Error spawning browser! Aborting!", vbCritical Or vbOKOnly
                Exit Function
            End If
        End If
        
        ' give it a bit to startup (may take a while if machine is swapping or antivirus really slow or other forced software running...)
        Do Until IsProcessRunning(ProcessName:=getProcessName(whichBrowser))
            Sleep 1
        Loop
        ' TODO also make sure process is visible (not just started)
        
        ' get the path to the browser target websocket - we hard code only connecting to localhost port 9222
        Dim wsPath As String
        
        ' we get the path (really full link) by connecting via HTTP to http://localhost:9222/json/version - note 9222 is same as specified on cmdline
        ' and extracting the information from the webSocketDebuggerUrl field of the returned JSON
        ' eg. "ws://localhost:9222/devtools/browser/f875efa4-b4ee-4c35-848b-73bc384a32bb"
        On Error Resume Next
GetWebSocketEndPoint:
        wsPath = wsBrowser.HttpGetMessage("localhost", 9222, "/json/version")
        If Err.Number <> 0 Then
            Debug.Print "Obtaining websocket endpoint - error: " & Err.description
            Sleep 0.5
            GoTo GetWebSocketEndPoint
        End If
        Dim versionInfo As Dictionary
        Set versionInfo = JsonConverter.ParseJson(wsPath)
        If Err.Number <> 0 Then
            Debug.Print "Retrying attempt to determine websocket endpoint"
            GoTo GetWebSocketEndPoint
        End If
        On Error GoTo 0
        
        wsPath = versionInfo("webSocketDebuggerUrl")
        wsPath = Right(wsPath, Len(wsPath) - Len("ws://localhost:9222"))
    
        ' connect to the browser target websocket
        With wsBrowser
            .protocol = "ws://"
            .Server = "localhost"
            .port = 9222
            .path = wsPath   ' this one changes with each time browser is started
            If Not .Connect() Then Exit Function ' False
        End With
    Else
        Set objBrowser = New clsProcess
        Set wsBrowser = Nothing
        Set LaunchBrowser = objBrowser
        strCall = """" & getProcessPathAndName(whichBrowser) & """ --remote-debugging-pipe " & additionalInlineCommands & " " & url
        
        ' create pipes and spawn browser
        Dim intRes As Integer
        intRes = objBrowser.init(strCall)
        If intRes <> 0 Then
            MsgBox "Error spawning and connecting pipe to browser! Aborting!", vbCritical Or vbOKOnly
            Exit Function
        End If
        
        ' give it a bit to startup (may take a while if machine is swapping or antivirus really slow or other forced software running...)
        Do Until IsProcessRunning(ProcessName:=getProcessName(whichBrowser))
            Sleep 1
        Loop
    End If
End Function

