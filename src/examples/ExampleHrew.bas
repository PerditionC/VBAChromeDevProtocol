Attribute VB_Name = "ExampleHrew"
' example on page that has frames and loads new windows and has shadow dom
Option Explicit


Private Sub testNewWindows()
    Dim v As Variant
    Dim dict As Dictionary
    Dim node As cdpDOMNode
    
    Dim browser As AutomateBrowser
    Set browser = new_automateBrowser
    browser.launch
    ' Note: we need a wait to ensure page loaded, events aren't fired during launch so a navigate or timing delay is needed to ensure page is ready
    browser.navigate "https://www.hrew-ces.com/AppsBtnContent.aspx"
    
    Set node = Nothing
    Do While node Is Nothing
        Set node = browser.querySelector("#ctl00_HyperLink9", waitUntilFound:=False)
        Debug.Print "."
        sleep 1
    Loop
    
    ' show currently active targets, this will change when a new window is opened
    'Set v = browser.cdp.Target.getTargets
    'Debug.Print "Targets: " & JsonConverter.ConvertToJson(v)
    
    ' demonstrate getting frame we want to access, not needed as long as we know expected frame name
    Dim frame As cdpPageFrame
    Set frame = browser.mainSession.frames.Item(browser.mainSession.frameNames.Item("trackerContent"))
    
    ' enable DOM tracking and get document root - most be done so nodes returned are tracked, i.e. nodeId is set to nonzero value
    browser.cdp.DOM.enable
    Dim root As cdpDOMNode
    Set root = browser.Document() ' pierce is optional, if calling browser.cdp.DOM.Document directly, ensure Depth:=-1
        
    ' get search box and set value
    Set node = browser.QuerySelectorOnFrame("#txtTrakContNum", "trackerContent")
    ' set search value
    ' Note: can not use browser.cdp.DOM.setNodeValue node.nodeId, "AMFU8920246" as this won't do as expected since value is stored in an attribute not directly
    'browser.cdp.DOM.setAttributeValue node.nodeId, "value", "AMFU8920246"
    node.elementValue = "AMFU8920246"
    
    ' get status button
    Set node = browser.QuerySelectorOnFrame("#btnStatus", "trackerContent")
    ' and do the search
    browser.Click node.nodeId, strategy:=WindowOpen ' click, can't wait for page to load as it opens in a new window/new target that we need to attach to first
    
    ' new window open, so targets updated
    'Set v = browser.cdp.Target.getTargets
    'Debug.Print "Targets: " & JsonConverter.ConvertToJson(v)
    
    ' attach to the new window and wait for it to finish loading
    ' but first we must wait a small bit for the browser to initiate the opening
    browser.attachToChildTargets navigationStrategy.NetworkIdle
    
    ' switch to newly opened window so we can process and close it
    Dim winSessionId As String
    For Each v In browser.sessions.Keys
        If v <> browser.mainSession.sessionId Then
            browser.switchTo v
            winSessionId = v
            Exit For
        End If
    Next v
    
    browser.cdp.Page.BringToFront
    Debug.Print JsonConverter.ConvertToJson(browser.Document.asDictionary())
    
    ' get TR element starting with tracking number and traverse its TD elements
    Dim tr As cdpDOMNode
    Set tr = browser.querySelector("#VisibleReportContentMainRptViewerID_ctl09 table table table table > tbody > tr:nth-child(3)")
    Dim td As cdpDOMNode
    Set td = tr.Children(1)
    Do While Not td Is Nothing
        Debug.Print td.Value & " -- " & td.innerText
        Set td = td.NextSibling
    Loop
    
    Dim nodes As Dictionary
    Set nodes = browser.querySelectorAll("#VisibleReportContentMainRptViewerID_ctl09 table table table table > tbody > tr:nth-child(3) td")
    Dim columnCount As Long
    columnCount = 7
    For Each v In nodes.Items
        Set td = v
        Debug.Print td.Value & " -- " & td.innerText
        columnCount = columnCount - 1: If columnCount < 1 Then Exit For
    Next v
    
    ' return to original/main window
    browser.switchToMainWindow
    browser.cdp.Page.BringToFront
    
    Set tr = browser.querySelector("#VisibleReportContentMainRptViewerID_ctl09 table table table table > tbody > tr:nth-child(3)", waitUntilFound:=False)
    If tr Is Nothing Then
        Debug.Print "Success"
    Else
        Debug.Print "Failure"
    End If
    
    ' close the window
    browser.switchTo winSessionId
    browser.cdp.Page.closePage
    browser.switchToMainWindow
    
    ' search for next container
    Set node = browser.QuerySelectorOnFrame("#txtTrakContNum", "trackerContent")
    browser.cdp.DOM.setAttributeValue node.nodeId, "value", "BADU12345678"
    
    Stop
    ' and wait a few seconds before closing browser
    sleep (5)
    browser.Quit
End Sub


