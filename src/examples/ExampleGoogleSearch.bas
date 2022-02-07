Attribute VB_Name = "ExampleGoogleSearch"
Option Explicit


Sub Google_search()
    Dim b As AutomateBrowser: Set b = new_automateBrowser
    b.launch
    b.navigate "https://www.google.com"
    
    'find node by class
    Dim SearchBar As cdpDOMNode
    Set SearchBar = b.querySelector(".gLFyf.gsfi")
    
    'do something to/with the node; i like having the methods here vs by b.method node.nodeId, etc
    SearchBar.elementAttribute("value") = "cats"
    sleep 1
    SearchBar.elementAttribute("value") = ""
    sleep 1
    SearchBar.elementValue = "big cats"
    ' print out some of its properties
    Debug.Print SearchBar.innerText
    Debug.Print SearchBar.Title
    Debug.Print SearchBar.Value
    Debug.Print SearchBar.elementValue
    
    
    sleep 1
    Dim myInputs As Dictionary
    Dim myInput As cdpDOMNode
    
    'find nodes by tag
    Set myInputs = b.querySelectorAll("input[class~='gLFyf'][class~='gsfi']")
    Set myInput = myInputs.Items(0)
    myInput.SetValue "tiger", SNV_FakeInput
    
    sleep 1
    
    myInput.SetValue "lion", SNV_Clipboard
    ' clear popup list
    b.cdp.SimulateInput.dispatchKeyEvent "keyDown", code:="Escape", windowsVirtualKeyCode:=27
    b.cdp.SimulateInput.dispatchKeyEvent "keyUp"
    
    'find a node by class name
    'Set myInputs = b.getNodesByClassName("gLFyf gsfi")
    'Set myInput = myInputs.Items(0)
    myInput.elementValue = "baby cats"
    Debug.Print myInput.elementValue
    
    sleep 1
    
    'loop through nodes to find a specific node
    Dim v As Variant
    For Each v In myInputs.Items
        Set myInput = v
        If myInput.elementAttribute("class") = "gLFyf gsfi" Then
            myInput.SetValue "kittens"
            Exit For
        End If
    Next v
    
    ' clear popup list
    sleep 1 ' needed before we press escape to give time for popup to pop up
    b.cdp.SimulateInput.dispatchKeyEvent "keyDown", code:="Escape", windowsVirtualKeyCode:=27
    b.cdp.SimulateInput.dispatchKeyEvent "keyUp"
    
    'find node by exact selector path as copied from debugger
    Dim SearchButton As cdpDOMNode
    Set SearchButton = b.querySelector("body > div.L3eUgb > div.o3j99.ikrT4e.om7nvf > form > div:nth-child(1) > div.A8SBwf > div.FPdoLc.lJ9FBc > center > input.gNO89b")
    SearchButton.Click
    
    sleep 1
    b.alert "weeee it worked. click 'ok' to close"
    
    b.Quit
End Sub

