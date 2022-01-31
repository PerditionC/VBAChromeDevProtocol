VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public nodeId, parentId, backendNodeId, nodeType, childNodeCount As Double

Public frameId As String
Public error As Dictionary
Public sessionId As String

Public nodeName, localName, nodeValue, documentURL, xmlVersion, baseURL, publicId, systemId, internalSubset, name, CompatibilityMode As String
Public nth_child, n_siblings As Integer
Public parentSelector, QuerySelector, ShadowRootType As String
Public isSVG, pierce As Boolean
Public root, templateContent As Dictionary
Public contentDocument_collection, children_collection, attributes_collection, pseudoElements_collection, distributedNodes_collection, shadowRoots_collection As Collection
Public objBrowser As AutomateBrowser

Private root_children As New Collection

Private Type Download
    File_Name As String
    File_Path As String
End Type

' sets value of element
Public Property Let value(ByVal value_str As String)
    Me.objBrowser.jsEval Me.QuerySelector & ".value = '" & value_str & "'"
End Property

' returns current value of element
Public Property Get value() As String
    value = Me.objBrowser.jsEval(Me.QuerySelector & ".value;")
End Property


'Return the InnerText
'Only works if element is childless
Public Property Get InnerText() As String
    InnerText = Me.objBrowser.jsEval(Me.QuerySelector & ".innerText")
End Property

'Returns the Outer
Public Property Get outerHTML() As String
    outerHTML = ConvertToJson(Me.objBrowser.cdp.DOM.getOuterHTML(Me.nodeId))
End Property

Public Property Get getAttribute(ByVal return_value_of_this_attr As String) As String
    
    Dim index As Integer
    Dim AttributeType, attributeValue As String
    
    For index = 1 To Me.attributes_collection.Count Step 2
        AttributeType = Me.attributes_collection(index)
        attributeValue = Me.attributes_collection(index + 1)
        
        If AttributeType = return_value_of_this_attr Then
            getAttribute = attributeValue
            Exit Property
        End If
    Next index
    
    getAttribute = ""
End Property

'Return a clsElements object that is limited to the children of Me
'Similar to objBrowser.Document
Property Get children() As clsElements
    Dim d As New clsElements
        
    d.Add_Element Me.root, Me.objBrowser, Me.parentSelector, Me.nth_child, Me.n_siblings, Me.pierce
 
    root_children.Add d

    Set children = root_children(1)

End Property

Public Sub setAttribute(ByVal AttributeType As String, ByVal attributeValue As String)
    Me.objBrowser.cdp.DOM.setAttributesAsText Me.nodeId, AttributeType, "'" & attributeValue & "'"
    Sleep 0.1
End Sub

Public Sub Click()
    Me.objBrowser.jsEval Me.QuerySelector & ".click();"
End Sub

Public Function Click_Download() As String
    Dim fso As Object
    Dim Downloads_Folder As Scripting.Folder
    Dim Downloads As New Dictionary
    Dim file As Scripting.file
    
    Set Downloads = Nothing
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Downloads_Folder = fso.GetFolder(Environ("USERPROFILE") & "\downloads")
    
    For Each file In Downloads_Folder.files
        Downloads.Add file.name, 1
    Next
    
    Me.Click
    
    Dim s As Single: s = Timer
    Dim periods As String: periods = "."
    Dim Downloading As String
    Dim Elapsed As Single
    
    Do
        Downloading = "unknown"
        
        For Each file In Downloads_Folder.files
            If Downloads.Exists(file.name) = False Then
                If InStr(file.name, ".tmp") > 0 Or InStr(file.name, ".crdownload") > 0 Then
                    'File is downloading... please wait
                    Downloading = "downloading"
                Else
                    Downloading = "complete"
                    Click_Download = file.path
                    Exit For
                End If
            End If
        Next file
        
        DoEvents
        
        Elapsed = Round(Timer - s, 1)
        
        'update progress once per 2.5 seconds
        If Elapsed Mod 2.5 = 0 Then
            If Downloading = "unknown" Then
                Debug.Print "No new files detected. Trying again" & periods
            ElseIf Downloading = "downloading" Then
                Debug.Print "Status: Downloading" & periods
            ElseIf Downloading = "complete" Then
                Debug.Print "File successfully downloaded: " & Click_Download
                Exit Do
            End If
            
            periods = periods & "."
            If periods = "....." Then periods = "."
        End If
        
        Sleep 0.5
        
        If Elapsed > 1000 Then
            MsgBox "There was a problem downloading the file! Please try again or contact the macro administrator for help."
            Me.objBrowser.quit
            End
        End If
        
    Loop Until Elapsed > 30 And Downloading = "unknown"
    
End Function
