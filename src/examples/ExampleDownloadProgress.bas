Attribute VB_Name = "ExampleDownloadProgress"
Option Explicit


Private Sub test_download()
    Dim browser As AutomateBrowser: Set browser = new_automateBrowser
    browser.launch "https://www.google.com/"
    
    Dim two As Double: two = browser.jsEval("1+1")
    browser.alert "You're on Google! Click 'ok' to close the browser. [1+1=" & two & "]"
    
    Dim doc As cdpDOMNode
    Set doc = browser.Document
    PrintNode doc
    
    Dim winState As WindowState, bounds As cdpBrowserBounds
    Set bounds = browser.getWindowBounds(state:=winState) ' get both the bounds and current window state
    
    Debug.Print "left=" & bounds.Left & " top=" & bounds.Top & " width=" & bounds.Width & " height=" & bounds.Height
    browser.resizeWindow Width:=bounds.Width / 2
    sleep 1
    browser.resizeWindow Width:=bounds.Width
    sleep 1
    browser.resizeWindow state:=WindowState.WS_maximized
    sleep 1
    browser.showNormal
    sleep 1
    browser.ShowMinimized
    sleep 1
    browser.showNormal
    
    Dim dlProgressBar As ehDownloadProgress: Set dlProgressBar = New ehDownloadProgress
    browser.navigate "https://www.ibiblio.org/pub/micro/pc-stuff/freedos/files/distributions/1.3/previews/1.3-rc5/"
    browser.cdp.Page.enable ' we need to enable so we can get the download events
    'browser.CDP.registerEventHandler "Browser.downloadWillBegin", dlProgressBar ' currently not supported with Edge but will replace deprecated Page events
    'browser.CDP.registerEventHandler "Browser.downloadProgress", dlProgressBar
    browser.cdp.registerEventHandler "Page.downloadWillBegin", dlProgressBar
    browser.cdp.registerEventHandler "Page.downloadProgress", dlProgressBar
    browser.alert "Download a zip file from here to see progress bar"
    dlProgressBar.Status = vbNullString ' reset
    browser.cdp.clearMessageQueue ' we don't want to wait while a backlog of messages are processed as user may do all sorts of navigation to get to here
    
    Do While dlProgressBar.Status <> "completed"
        'browser.jsEval ("1+1") ' just to drain CDP messages so events processed, could also be javascript that initiates the download
        browser.cdp.peakMessage
    Loop
    
    'This closes the browser cleanly
    browser.Quit
    Set browser = Nothing
End Sub

