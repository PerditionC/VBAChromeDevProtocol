Attribute VB_Name = "ExampleLaunch"

Option Explicit

Dim browser As Object
Dim url As String
Dim inlineCommand As String

Const kill = True


'
' Sometimes, you don't need to control the Browser
'  Like opening an help file address on a new browser
'
Sub runBrowsers() ' This example will open a new instance of each of the installed browsers without control of HTML

   If startBrowserNoControl(browserType.Chrome, "https://github.com/PerditionC/VBAChromeDevProtocol#readme") = False Then MsgBox ("Browser Chrome not found, did not execute")
   If startBrowserNoControl(browserType.Edge, "https://www.google.com/search?q=github+perditionc") = False Then MsgBox ("Browser Edge not found, did not execute")
   If startBrowserNoControl(browserType.firefox, "https://www.google.com/search?q=github+help") = False Then MsgBox "Browser Firefox not found, did not execute", vbSystemModal

End Sub
 

Sub showStatusOfAllSupportedBrowsers()
   
   Dim msg As String
   msg = "Edge:" & vbTab & browserStatus(Edge) & vbCrLf & "Chrome:" & vbTab & browserStatus(Chrome) & vbCrLf & "Firefox:" & vbTab & browserStatus(firefox)
   MsgBox msg, vbOKOnly, "Status of supported browsers"
   
End Sub

Public Sub forceLaunchChrome()
   
   Set browser = New_AutomateBrowser
   url = "https://www.google.com/search?q=github+perditionc"
   inlineCommand = "--disable-print-preview"
   
   browser.launch whichBrowser:=browserType.Chrome, url:=url, additionnalInlineCommands:=inlineCommand, killWithoutAsking:=kill
    
   Call doYourThings
    
   ' Close browser
   On Error Resume Next
   browser.Quit
   Set browser = Nothing
   On Error GoTo 0
   
End Sub

Public Sub forceLaunchChromeIfInstalled()
   
   Set browser = New_AutomateBrowser
   url = "https://www.google.com/search?q=github+perditionc"
   inlineCommand = "--disable-print-preview"
   
   If browserStatus(Chrome) = "Not installed" Then
      MsgBox ("Chrome : " & browserStatus(Chrome) & vbLf & _
             "Requets aborted")
      Set browser = Nothing
   End If
   
   browser.launch whichBrowser:=browserType.Chrome, url:=url, additionnalInlineCommands:=inlineCommand, killWithoutAsking:=kill
    
   Call doYourThings
   
   ' Close browser
   On Error Resume Next
   browser.Quit
   Set browser = Nothing
   On Error GoTo 0

End Sub
Public Sub forceLaunchChromeDefaultAsk()
   
   Set browser = New_AutomateBrowser
   url = "https://www.google.com/search?q=github+perditionc"
   inlineCommand = "--disable-print-preview"
   
   If browserStatus(Chrome) = "Not installed" Then
      MsgBox ("Chrome : " & browserStatus(Chrome) & vbLf & _
             "Requets aborted")
      Set browser = Nothing
   End If
   
   browser.launch whichBrowser:=browserType.Chrome, url:=url, additionnalInlineCommands:=inlineCommand

    
   Call doYourThings
   
   
   ' Close browser
   On Error Resume Next
   browser.Quit
   Set browser = Nothing
   On Error GoTo 0

End Sub
Public Sub launchChromeCustomAsk()
   
   Set browser = New_AutomateBrowser
   url = "https://www.google.com/search?q=github+perditionc"
   inlineCommand = "--disable-print-preview"
   
   If browserStatus(Chrome) = "Not installed" Then
      MsgBox ("Chrome : " & browserStatus(Chrome) & vbLf & _
             "Requets aborted")
      Set browser = Nothing
   End If
   
   If browserStatus(Chrome) = "Running" Then
      If MsgBox("Browser Chrome already running and needs to be killed to proceed with current action. Is it OK to close Chrome and loose current pages in all tabs ?", vbYesNo) = vbNo Then
         MsgBox ("Request cancelled")
      End If
   End If
      
   browser.launch whichBrowser:=browserType.Chrome, url:=url, additionnalInlineCommands:=inlineCommand, killWithoutAsking:=kill

    
   Call doYourThings
   
   ' Close browser
   On Error Resume Next
   browser.Quit
   Set browser = Nothing
   On Error GoTo 0

End Sub
Public Sub forceLaunchAnyBrowser()
   
   Set browser = New_AutomateBrowser
   url = "https://www.google.com/search?q=github+perditionc"
   inlineCommand = ""
   
   browser.launch url:=url, killWithoutAsking:=kill

    
   Call doYourThings
   
   
   ' Close browser
   browser.Quit
   Set browser = Nothing

End Sub

Public Sub launchReorderedPriorityBrowser()

   Dim BT As browserType
   Set browser = New_AutomateBrowser
   url = "https://www.google.com/search?q=github+perditionc"
   inlineCommand = ""
      
   '
   ' This example shows possibilities, don't question why I used such order, make it your own
   '
   BT = getFirstAvailableBrowser(firefox, Chrome, Edge)
   
   If BT = noBrowser Then ' No browser is available (installed and not running)
      
      BT = getFirstInstalledBrowser(Chrome, firefox, noBrowser)
      If BT = noBrowser Then ' No browser is installed, most likely never to happen
         MsgBox ("Logic selected no browser, request aborted")
         Call showStatusOfAllSupportedBrowsers
         Exit Sub
      End If
   End If
   
   '
   ' At this point, browserType is chosen,
   '
   browser.launch whichBrowser:=BT, url:=url, killWithoutAsking:=kill
    
   Call doYourThings
   
   ' Close browser
   On Error Resume Next
   browser.Quit
   Set browser = Nothing
   On Error GoTo 0
   
End Sub

Private Sub doYourThings()

   '
   ' At this point, you need to verify if «browser» is at the right html page
   ' See other examples in this example.xlsm file
   '
   MsgBox (browser.document.Title)
   
End Sub
