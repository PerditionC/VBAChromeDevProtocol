Attribute VB_Name = "ExampleBasic"
' minimal tests for basic functionality
Option Explicit

Sub test1a()
    Dim cdp As clsCDP
    Set cdp = New_clsCDP
    cdp.launch "https://www.google.com/"
    Debug.Print JsonConverter.ConvertToJson(cdp.browser.getWindowForTarget)
    cdp.browser.closeBrowser
End Sub

Sub test1a2()
    Dim cdp As clsCDP
    Set cdp = New_clsCDP
    cdp.launch "https://www.google.com/", useWebSocket:=True, useExistingBrowser:=False
    Debug.Print JsonConverter.ConvertToJson(cdp.browser.getWindowForTarget)
    cdp.browser.closeBrowser
End Sub


Sub test1b()
    Dim browser As AutomateBrowser
    Set browser = new_automateBrowser
    browser.launch "https://www.google.com/"
    Debug.Print JsonConverter.ConvertToJson(browser.cdp.browser.getWindowForTarget)
    browser.Quit
End Sub


Sub test2()
    Dim browser As AutomateBrowser
    Set browser = new_automateBrowser
    browser.launch
    browser.navigate "https://www.google.com/"
    Debug.Print JsonConverter.ConvertToJson(browser.cdp.browser.getWindowForTarget)
    browser.resizeWindow , , , , , WindowState.WS_maximized
    Stop
    
    Dim hWnd As Long
    hWnd = findBrowserHWnd(limitTo:=Edge)
    ShowWindowAsync hWnd, ShowWindowCmd.SW_HIDE
    ShowWindowAsync hWnd, ShowWindowCmd.SW_NORMAL

    browser.showNormal
    browser.Quit
End Sub


Sub test3()
    Dim b As AutomateBrowser: Set b = new_automateBrowser
    'b.launch "http://www.google.com/"
    b.launch "www.google.com"
    Debug.Print JsonConverter.ConvertToJson(b.cdp.browser.getWindowForTarget)
    
    
    'via CDP
    b.resizeWindow , 0, 0, 572, 572
    If b.cdp.ErrorCode <> 0 Then Exit Sub
    sleep 2
    b.resizeWindow , , , 800, 800
    
    b.alert "Hello from google. Let's do a search!"
    
    b.hide
    
    Dim doc As cdpDOMNode
    Set doc = b.Document
    
    Dim SearchBar As cdpDOMNode
    Set SearchBar = b.querySelector("[class~=gLFyf][class~=gsfi]", doc.nodeId)
    
    'SearchBar.value = "kittens"
    b.cdp.DOM.setAttributeValue SearchBar.nodeId, "value", "kittens"
    
    Dim Search As cdpDOMNode
    Set Search = b.querySelector(".gNO89b", doc.nodeId)
    
    'Search.Click
    b.Click Search.nodeId, strategy:=Normal
    
    b.showNormal
    b.resizeWindow , , , , , WindowState.WS_minimized
    b.showNormal
    b.alert "Ohhhh such cool search results!"
    Stop
    b.Quit
End Sub


Private Sub test4()
    Dim browser As AutomateBrowser
    Set browser = new_automateBrowser
    ' goto Google and attach session
    browser.cdp.launch "http://www.google.com/"
    ' page is usually loaded by the time attach has completed if loaded during launch
    ' but if we navigate to a page, we need to wait for it to finish loading
    browser.cdp.Page.enable
    browser.cdp.Page.navigate "http://www.google.com/"
    Debug.Print "page loaded?"
    browser.cdp.Page.navigate "http://www.fdos.org/"
    Debug.Print "page loaded?"
    browser.cdp.Page.navigate "http://www.google.com/"
    Debug.Print "page loaded?"
    browser.cdp.Page.disable
    ' get document root
    Dim root As Dictionary, rootId As Integer
    Set root = browser.cdp.DOM.getDocument
    rootId = root("nodeId")
    ' get search box and set value
    Dim nodeId As Integer
    nodeId = browser.cdp.DOM.querySelector(rootId, "body > div.L3eUgb > div.o3j99.ikrT4e.om7nvf > form > div:nth-child(1) > div.A8SBwf > div.RNNXgb > div.SDkEP > div.a4bIc > input")
    ' set search value
    browser.cdp.DOM.setAttributeValue nodeId, "value", "happy kitten"
    ' get search button
    Dim btnId As Long
    btnId = browser.cdp.DOM.querySelector(rootId, "body > div.L3eUgb > div.o3j99.ikrT4e.om7nvf > form > div:nth-child(1) > div.A8SBwf > div.FPdoLc.lJ9FBc > center > input.gNO89b")
    ' and do the search
    'cdp.DOM.focus nodeId:=btnId
    'cdp.Overlay.highlightNode
    browser.Click btnId, navigationStrategy.Normal ' click, but since we expect a page to load, wait for it to be loaded before continueing
    Debug.Print "click complete"
    Set root = browser.cdp.DOM.getDocument
    rootId = root("nodeId")
    Dim resultsId As Integer
    resultsId = browser.cdp.DOM.querySelector(rootId, "#result-stats")
    Debug.Print resultsId
    ' wait for page to load
    ' and wait a few seconds before closing browser
    sleep (5)
    browser.Quit
End Sub


Sub test5()
    Dim browser As AutomateBrowser, doc As cdpDOMNode
    Set browser = new_automateBrowser
    browser.launch
    browser.navigate "https://www.google.com/"
    Set doc = browser.Document(pierce:=True)
    PrintNode doc
    Stop
    Debug.Print doc.textContent
    browser.Quit
End Sub


Private Sub test6()
    Dim browser As AutomateBrowser
retry:
    Set browser = new_automateBrowser
    ' this will occasionally fail as the url will periodically append "?gws_rd=ssl" to the end of the requested url
    If Not browser.launch("http://www.google.com/", partialMatch:=False) Then Exit Sub
    browser.Quit
    sleep (2)
    DoEvents
    GoTo retry
End Sub



