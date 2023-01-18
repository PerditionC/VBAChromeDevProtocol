Attribute VB_Name = "Example_iFrame"
'
' This is my understanding of how cdp handles iFrame(s)
' Feel free to adjust my explanations
'
'
' CDP recognizes <iFrame> nodes but it can't access the iFrame node.
'
' The iFrame(s) are evaluated and converted to frames and their names are stored in .mainsession.framenames
'
' You need the text value of the framenames in order to access its nodes in .mainsession.frames
' For an unknown reason, I did not succeed in using .mainsession.framenames(i) as a string value, it doesn't return a string value.
' I used JsonConverter to extract the string value of the names

Dim navigator As Object
Dim useBrowser As browsertype
Dim oneNode As cdpDOMNode
Dim allNodes As Object


'
' When you know the iFrame name, you can call querySelectorOnFrame or querySelectorAllOnFrame with the HARD-CODED value
'
'    set oneNode = navigator.querySelectorOnFrame("a[name='serialNumber']","exactframename")
'    set allNodes= navigator.querySelectorAllOnFrame("a","exactframename")
'

'
' When you don't know the frame name, you need to convert .mainsession.framenames to string ( using json )
' and parse the string to retrieve frame name
'
'

Sub exampleUsing_iFrame()

   Set navigator = new_automateBrowser
   useBrowser = browsertype.Chrome
   navigator.launch whichBrowser:=useBrowser, additionnalInlineCommands:="--allow-file-access-from-files", killWithoutAsking:=True
    
   navigator.navigate "https://nunzioweb.com/iframes-example.htm"  'any web page that contains <iframe
   Application.Wait (Now + TimeSerial(0, 0, 3))
   
   If navigator.mainsession.framenames.Count = 0 Then
      Stop ' The specified web page does not contain iFrame
      GoTo noFrames
   End If
   
  
   Dim jsonTxt As String
   Set allNodes = navigator.queryselectorall("iframe", pierce:=True)
   Debug.Print ""
   Debug.Print "List of <iFrame (s)"
   
   ' «set oneNode = allNodes(i)»   will not work for iFrames so I display them using JsonConverter
   jsonTxt = JsonConverter.ConvertToJson(allNodes)
   jsonTxt = Mid(jsonTxt, 3)
   For i = 1 To allNodes.Count
      Debug.Print "<iframe " & Left(jsonTxt, InStr(1, jsonTxt, """") - 1)
      jsonTxt = Mid(jsonTxt, InStr(1, jsonTxt, """") + 4)
   Next
   
   ' Get framenames list - needs to convert to text via JsonConverter
   jsonTxt = JsonConverter.ConvertToJson(navigator.mainsession.framenames)
   ' Separate names
   Dim frameNameTxt(10) As String
   Dim nuFrame As Integer
   nuFrame = 0
   
   Debug.Print ""
   Debug.Print "List of frames from .mainsession.framenames"
   jsonTxt = Mid(jsonTxt, 2)
   Do While Len(jsonTxt) > 1
      jsonTxt = Mid(jsonTxt, 2)
      nuFrame = nuFrame + 1
      frameNameTxt(nuFrame) = Left(jsonTxt, InStr(1, jsonTxt, """") - 1)
      jsonTxt = Mid(jsonTxt, InStr(1, jsonTxt, """") + 3) ' skip ending quote and comma 
      jsonTxt = Mid(jsonTxt, InStr(1, jsonTxt, """") + 2) ' skip frame id ending quote and comma
      Debug.Print "Frame: " & frameNameTxt(nuFrame)
   Loop
   
   '
   ' Loop thru frames to search for <a  (by example)
   '
  
   For i = 1 To nuFrame
      Debug.Print
      Debug.Print "Content of frame: " & frameNameTxt(i)
      
      ' access directly all tage <a
      Set allNodes = navigator.queryselectorallonframe("a", frameNameTxt(i), pierce:=True)
      
      ' or access tag <body and search within its node
      Set allNodes = navigator.queryselectorallonframe("body", frameNameTxt(i), pierce:=True)
      Set oneNode = allNodes.Item(0)
      Set allNodes = oneNode.queryselectorall("a")
      
      If allNodes.Count = 0 Then
         Debug.Print "No html tag <a found"
      Else
         For j = 1 To allNodes.Count
            Set oneNode = allNodes.Item(j - 1)
            Debug.Print oneNode.outerHTML
         Next
      End If
      
      'Set allNodes = oneNode.querySelectorAll("a", pierce:=True)
      'Set allNodes = oneNode.querySelectorAll("div", pierce:=True)
      

   Next
   
   Stop
noFrames:
   navigator.Quit
   

End Sub
