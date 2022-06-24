Attribute VB_Name = "browsers"
' provides information about supported browsers
Option Explicit

Private Const processNameEdge As String = "msedge.exe"
Private Const browserExeSubPathEdge As String = "\Microsoft\Edge\Application\msedge.exe"

Private Const processNameChrome As String = "chrome.exe"
Private Const browserExeSubPathChrome As String = "\Google\Chrome\Application\chrome.exe"

Private Const processNameFirefox As String = "firefox.exe"
Private Const browserExeSubPathFirefox As String = "\Mozilla Firefox\firefox.exe"

Public Function browserStatus(ByVal oneBrowser As browserType) As String
   browserStatus = "Available"
   If getProcessPathAndName(oneBrowser) = vbNullString Then browserStatus = "Not installed": Exit Function
   If IsProcessRunning(ProcessName:=getProcessName(oneBrowser)) Then browserStatus = "Running": Exit Function
End Function

Private Function browserIsInstalled(ByVal oneBrowser As browserType)
   browserIsInstalled = False
   If getProcessPathAndName(oneBrowser) <> vbNullString Then browserIsInstalled = True
End Function

' returns just the process name for matching browser
Private Function getProcessName(ByVal whichBrowser As browserType)
    If (whichBrowser And browserType.Edge) = browserType.Edge Then
        getProcessName = processNameEdge
    ElseIf (whichBrowser And browserType.Chrome) = browserType.Chrome Then
        getProcessName = processNameChrome
    ElseIf (whichBrowser And browserType.Firefox) = browserType.Firefox Then
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

   getProcessPathAndName = vbNullString
   Dim browserExeSubPath As String
   
   browserExeSubPath = vbNullString
   If oneBrowser = browserType.Edge Then browserExeSubPath = browserExeSubPathEdge
   If oneBrowser = browserType.Chrome Then browserExeSubPath = browserExeSubPathChrome
   If oneBrowser = browserType.Firefox Then browserExeSubPath = browserExeSubPathFirefox
   If browserExeSubPath = vbNullString Then Exit Function
   
   ' Get «program files (X86)» exact path  (sometimes there is a space before (X86) so get it with Environ$
  
   Dim programFiles As String
   Dim programFilesX86 As String
   programFilesX86 = Environ$("PROGRAMFILES(X86)")
   If Environ$("ProgramW6432") = "" Then
      programFiles = Environ$("ProgramFiles")
   Else
      programFiles = Environ$("ProgramW6432")
   End If
   '
   ' Sometimes Firefox is installed in AppData/Local... (documented on internet but not tested...)
   Dim appDataLocal As String
   appDataLocal = Environ$("LocalAppData")
   
   ' verify if browser exists in Program Files
   If Dir(programFiles & browserExeSubPath, vbNormal) <> "" Then getProcessPathAndName = programFiles & browserExeSubPath: Exit Function
      
   ' verify if browser exists in Program Files (X86)
   If Dir(programFilesX86 & browserExeSubPath, vbNormal) <> "" Then getProcessPathAndName = programFilesX86 & browserExeSubPath: Exit Function
         
   ' verify if browser exists in AppData\Local
   If Dir(appDataLocal & browserExeSubPath, vbNormal) <> "" Then getProcessPathAndName = appDataLocal & browserExeSubPath: Exit Function

   
    
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
' killExistingInstance: (if all running)  true=kill, don't ask     false=Prompt before killing
' browserInlineCommands: options like «--enable-automation» to add to the call string
'
' if a websocket is to be used then see if useExistingBrowser is requested,
' if not use same algorithm to determine browser as direct pipe case
' but if want to reuse we search for existing process and only if not found do we
' spawn, not AnyBrowser and Chromium will always spawn/use Edge, otherwise specified browser used
' Note: Edge is used first and Chrome only because Chrome is generally in use by person
' and Firefox may not be installed and its CDP support is not as mature
'
' at the end, whichBrowser is set to browser used for connection and connection is returned
'
' Note: AutomateBrowser.launch calls clsCDP.launch that calls LaunchBrowser
'
Public Function LaunchBrowser( _
    ByRef whichBrowser As browserType, _
    Optional url As String = vbNullString, _
    Optional useWebSocket As Boolean = False, _
    Optional useExistingBrowser As Boolean = False, _
    Optional browserInlineCommands As String = "", _
    Optional killExistingInstance As Boolean = False _
    ) As Object
    Dim browserConnectionObj As Object
    Dim objBrowser As clsProcess
    Dim wsBrowser As clsWebSocket
    
    ' check for browser running
    Dim runningBrowsers As browserType
    If IsProcessRunning(ProcessName:=getProcessName(browserType.Edge)) Then runningBrowsers = runningBrowsers Or browserType.Edge
    If IsProcessRunning(ProcessName:=getProcessName(browserType.Chrome)) Then runningBrowsers = runningBrowsers Or browserType.Chrome
    If IsProcessRunning(ProcessName:=getProcessName(browserType.Firefox)) Then runningBrowsers = runningBrowsers Or browserType.Firefox
    
    ' select first requested browser that is not running if found
    If browserIsInstalled(browserType.Edge) And ((whichBrowser And browserType.Edge) = browserType.Edge) And ((useWebSocket And useExistingBrowser) Or ((runningBrowsers And browserType.Edge) <> browserType.Edge)) Then
        whichBrowser = browserType.Edge
    ElseIf browserIsInstalled(browserType.Chrome) And ((whichBrowser And browserType.Chrome) = browserType.Chrome) And ((useWebSocket And useExistingBrowser) Or ((runningBrowsers And browserType.Chrome) <> browserType.Chrome)) Then
        whichBrowser = browserType.Chrome
    ElseIf browserIsInstalled(browserType.Firefox) And ((whichBrowser And browserType.Firefox) = browserType.Firefox) And ((useWebSocket And useExistingBrowser) Or ((runningBrowsers And browserType.Firefox) <> browserType.Firefox)) Then
        whichBrowser = browserType.Firefox
    Else
        ' all requested browsers are running so we need to kill the existing browser
        If browserIsInstalled(browserType.Edge) And ((whichBrowser And browserType.Edge) = browserType.Edge) Then
            whichBrowser = browserType.Edge
        ElseIf browserIsInstalled(browserType.Chrome) And ((whichBrowser And browserType.Chrome) = browserType.Chrome) Then
            whichBrowser = browserType.Chrome
        ElseIf browserIsInstalled(browserType.Firefox) And ((whichBrowser And browserType.Firefox) = browserType.Firefox) Then
            whichBrowser = browserType.Firefox
        Else
            If IsProcessRunning(ProcessName:=getProcessName(whichBrowser)) Then   'And killExistingInstance
            Else
               Debug.Print "LaunchBrowser() - unable to determine browser to start!"
               Stop
               Exit Function
            End If
        End If
    End If
    
    ' At this point, whichBrowser contains the browser to start or kill and restart
    ' Verification was made that whichBrowser is installed on the running computer
    
    ' ensure browser is not already running (kill it if it is)
    If Not (useWebSocket And useExistingBrowser) Then
        If Not TerminateProcess(ProcessName:=getProcessName(whichBrowser), PromptBefore:=Not (killExistingInstance)) Then
            ' abort if browser is already running and failed to kill it or user elected not to terminate
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
        strCall = """" & getProcessPathAndName(whichBrowser) & """ --remote-debugging-port=9222 " & browserInlineCommands & " " & url
        If (Not useExistingBrowser) Or (Not IsProcessRunning(ProcessName:=getProcessName(whichBrowser))) Then
            If Not SpawnProcess(strCall) Then Exit Function
        End If
        
        ' give it a bit to startup
        Do Until IsProcessRunning(ProcessName:=getProcessName(whichBrowser))
            Sleep 1
        Loop
        ' TODO also make sure process is visible
        
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
            .server = "localhost"
            .port = 9222
            .path = wsPath   ' this one changes with each time browser is started
            If Not .Connect() Then Exit Function ' False
        End With
    Else
        Set objBrowser = New clsProcess
        Set wsBrowser = Nothing
        Set LaunchBrowser = objBrowser
        strCall = """" & getProcessPathAndName(whichBrowser) & """ --remote-debugging-pipe " & browserInlineCommands & " " & url
        
        Dim intRes As Integer
        intRes = objBrowser.init(strCall)
        If intRes <> 0 Then
           'Call Err.Raise(-99, , "error start browser")
        Exit Function ' False
        End If
        
        ' give it a bit to startup
        Sleep 1
    End If
End Function

