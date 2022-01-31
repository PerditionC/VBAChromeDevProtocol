VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ElementsCollection As New Collection

Property Get index(ByVal ElementIndex As Integer) As clsElement
    Set index = ElementsCollection(ElementIndex)
End Property

Property Get Count() As Integer
    Count = ElementsCollection.Count
End Property

Property Get asCollection() As Collection
    Set asCollection = ElementsCollection
End Property

'Extract a dictionary that counts the types of tags present and then convert to Json for comparision
Property Get DescribePage() As String
    
    Dim k As Variant
    Dim v, c As Integer
    Dim d As New Dictionary
    Dim element As clsElement
    
    For Each element In ElementsCollection
        k = element.nodeName
        If d.Exists(k) Then
            v = d(k)
            d(k) = v + 1
        Else
            d(k) = 1
        End If
        c = c + 1
    Next element
    
    If c > 25 Then
        DescribePage = ConvertToJson(d)
    'Else
        'Debug.Print c, DescribePage
    End If
End Property

'Scour each element in ElementsCollection and build a new clsElements collection
'based on clsElements ClassName = str_ClassName
'Returns a clsElements collection of clsElement objects
Public Function getElementsByClassName(ByVal str_ClassName As String) As clsElements
    
    Dim element As clsElement
    Dim elements As New clsElements
    
    For Each element In ElementsCollection
        If element.attributes_collection.Count > 0 Then
            
            Dim index As Integer
            Dim AttributeType, attributeValue As String
            
            For index = 1 To element.attributes_collection.Count Step 2
                AttributeType = element.attributes_collection(index)
                attributeValue = element.attributes_collection(index + 1)
                
                If AttributeType = "class" And attributeValue = str_ClassName Then
                    
                    elements.Add_Element_asCopy element
                End If
            Next index
        End If
    Next element
    
    Set getElementsByClassName = elements
End Function


'Scour each element in ElementsCollection and return the first item matching on class name
Public Function getElementByClassName(ByVal str_ClassName As String) As clsElement
    Set getElementByClassName = Me.getElementByAttribute("class", str_ClassName)
End Function


'Scour each element in ElementsCollection and build a new clsElements collection
'based on clsElement's TagName = str_TagName
'Returns a clsElements collection of clsElement objects
Public Function getElementsByTagName(ByVal str_TagName As String) As clsElements
    
    Dim element As clsElement
    Dim elements As New clsElements
    
    For Each element In ElementsCollection
        If element.nodeName = UCase(str_TagName) Or str_TagName = "*" Then
            elements.Add_Element_asCopy element
        End If
    Next element
    
    Set getElementsByTagName = elements

End Function


'Scour each element in ElementsCollection for AttributeType: id, value: str_id
'Returns a clsElement object
Public Function getElementById(ByVal str_id As String) As clsElement
    Set getElementById = Me.getElementByAttribute("id", str_id)
End Function


'Scour each element in ElementsCollection for AttributeType: str_AttributeType,  value: str_AttributeValue
'Returns the first matching clsElement object
Public Function getElementByAttribute(ByVal str_AttributeType As String, ByVal str_AttributeValue As String) As clsElement
    
    Dim element As clsElement
    Dim child As clsElement
    Dim children As clsElements

    For Each child In ElementsCollection
    'For Each element In ElementsCollection
        'For Each child In element.Children.asCollection
            If child.attributes_collection.Count > 0 Then
                
                Dim index As Integer
                Dim AttributeType, attributeValue As String
                
                For index = 1 To child.attributes_collection.Count Step 2
                    AttributeType = child.attributes_collection(index)
                    attributeValue = child.attributes_collection(index + 1)
                    'If AttributeType = str_AttributeType Then Debug.Print AttributeValue
                    If InStr(str_AttributeValue, "*") > 0 Then
                        'wildcard match
                        If AttributeType = str_AttributeType And attributeValue Like str_AttributeValue Then
                            Set getElementByAttribute = child
                            Exit Function
                        End If
                    Else
                        'exact Match
                        If AttributeType = str_AttributeType And attributeValue = str_AttributeValue Then
                            Set getElementByAttribute = child
                            Exit Function
                        End If
                    End If
                Next index
            End If
        Next child
    'Next element
End Function


'html tree >> flat

'build ElementsCollection by selecting specific clsElement objects from the
Public Sub Add_Element_asCopy(ByRef element As clsElement)
    ElementsCollection.Add element
End Sub


'Build ElementsCollection by reading through the supplied dictionary and it's children
'Children are recursively added until all branches have been traversed
Public Sub Add_Element(ByVal ElementDictionary As Dictionary, ByRef objBrowser As AutomateBrowser, _
                        ByVal parentSelector As String, ByVal nth_child As Integer, ByVal n_siblings As Integer, _
                        ByVal pierce As Boolean)
    
    Dim element As New clsElement
    Dim dKey As Variant
    
    With element
        Set .objBrowser = objBrowser
    
        Set .children_collection = New Collection
        Set .attributes_collection = New Collection
        Set .shadowRoots_collection = New Collection
        Set .pseudoElements_collection = New Collection
        Set .distributedNodes_collection = New Collection
        Set .contentDocument_collection = New Collection
        
        Set .root = ElementDictionary
        Set .error = New Dictionary
        
        .nth_child = nth_child
        .n_siblings = n_siblings
        .parentSelector = parentSelector
        .pierce = pierce
        
        For Each dKey In ElementDictionary.Keys

            Select Case dKey
                Case "attributes"
                    Set .attributes_collection = ElementDictionary(dKey)
                
                Case "backendNodeId"
                    .backendNodeId = UCase(ElementDictionary(dKey))
                
                Case "baseURL"
                    .baseURL = UCase(ElementDictionary(dKey))
                
                Case "childNodeCount"
                    .childNodeCount = UCase(ElementDictionary(dKey))
                
                Case "children"
                    Set .children_collection = ElementDictionary(dKey)
                
                Case "compatibilityMode"
                    .CompatibilityMode = UCase(ElementDictionary(dKey))
                
                Case "contentDocument"
                    'contentDocument is a dictionary
                    .contentDocument_collection.Add ElementDictionary(dKey)
                    
                Case "localName"
                    .localName = UCase(ElementDictionary(dKey))
                
                Case "distributedNodes"
                    Set .distributedNodes_collection = ElementDictionary(dKey)
                
                Case "documentURL"
                    .documentURL = UCase(ElementDictionary(dKey))
                
                Case "error"
                    Set .error = ElementDictionary(dKey)
                
                Case "frameId"
                    .frameId = UCase(ElementDictionary(dKey))
                
                'Case "importedDocument"
                
                Case "internalSubset"
                    .internalSubset = UCase(ElementDictionary(dKey))
                
                Case "isSVG"
                    .isSVG = UCase(ElementDictionary(dKey))
                
                'Case "name"
                '    .name = UCase(ElementDictionary(dKey))
                
                Case "nodeId"
                    .nodeId = UCase(ElementDictionary(dKey))
                
                Case "nodeName"
                    .nodeName = UCase(ElementDictionary(dKey))
                
                Case "nodeType"
                    .nodeType = UCase(ElementDictionary(dKey))
                
                Case "nodeValue"
                    .nodeValue = UCase(ElementDictionary(dKey))
                
                Case "parentId"
                    .parentId = UCase(ElementDictionary(dKey))
                
                Case "pseudoElements"
                    Set .pseudoElements_collection = ElementDictionary(dKey)
                
                'Case "pseudoType"
                
                Case "publicId"
                    .publicId = UCase(ElementDictionary(dKey))
                
                Case "sessionId"
                    .sessionId = ElementDictionary(dKey)
                
                Case "shadowRootType"
                    .ShadowRootType = UCase(ElementDictionary(dKey))
                
                Case "shadowRoots"
                    Set .shadowRoots_collection = ElementDictionary(dKey)
                
                Case "systemId"
                    .systemId = UCase(ElementDictionary(dKey))
                
                'Case "value"
                    'Debug.Print .value
                    '.value = UCase(ElementDictionary(dKey))
                
                Case "templateContent"
                    Set .templateContent = ElementDictionary(dKey)
                
                Case "xmlVersion"
                    .xmlVersion = UCase(ElementDictionary(dKey))
                
                Case Else
                    Debug.Print dKey, typeName(ElementDictionary(dKey))
                    Stop
            End Select
        Next dKey
        
        Call getQuerySelector(element, .parentSelector)                     'jsEval selector
        Call AddChildren(element, "children", .pierce)                       'regular DOM elements
        If pierce Then Call AddChildren(element, "shadowRoots", .pierce)     'shadowDOM (DOMgetDocument - Pierce TRUE required)
        If pierce Then Call AddChildren(element, "contentDocument", .pierce) 'iFrames (DOMgetDocument - Pierce TRUE required)
        
    End With
    
    ElementsCollection.Add element
        
End Sub

'Here we iterate through each child-type collections or dictionaries to fully explore the DOM and store it
Private Sub AddChildren(ByRef element As clsElement, ByVal WhichSetOfChildren As String, ByVal pierce As Boolean)
    
    Dim col As Collection
    Set col = New Collection
    Dim child As Dictionary
    Dim child_counter As Integer
    child_counter = 1
    
    Select Case WhichSetOfChildren
        Case "children"
            Set col = element.children_collection
        Case "shadowRoots"
            Set col = element.shadowRoots_collection
        Case "contentDocument"
            Set col = element.contentDocument_collection
        Case Else
            Debug.Print WhichSetOfChildren
            Stop
    End Select
    
    With element
        For Each child In col
            If InStr(child("nodeName"), "DOCTYPE html") > 0 Then Stop
            If child("nodeName") <> "#comment" And child("nodeName") <> "#text" Then
                Me.Add_Element child, .objBrowser, .QuerySelector, child_counter, col.Count, .pierce
                child_counter = child_counter + 1
            End If
        Next child
    End With
End Sub


'Here we create a querySelector path to the element for jsEval calls, see clsElement.click method for an example
'the querySelector string is of the following varieties:
'document.querySelector('body > DIV:nth-child(2) > DIV > FORM') or similar full path from body using n-child to distinguish
'document.querySelector('#elementid > DIV:nth-child(2) > A') or similar shortened path from element id
'document.querySelector('as above').shadowRoot.querySelector('as above') or similar to access shadowRoot children
'document.querySelector('something > IFRAME').contentDocument.querySelector('as above') or similar to access iframe elements
Private Sub getQuerySelector(ByRef element As clsElement, ByVal parentSelector As String)
    'Debug.Print parentSelector
    If Len(parentSelector) > 1 Then
        'trim parent selector for new addition
        parentSelector = left(parentSelector, Len(parentSelector) - 2)
    End If
    
    'Debug.Print element.nodeName
    'Stop

    'If element.nodeName = "TITLE" Then Stop
    
    'If element.nodeName = "BODY" And parentSelector = "" Then
    If parentSelector = "" Then
        element.QuerySelector = "document.querySelector('')"
        'start tracking for querySelector
        'element.querySelector = "document.querySelector('body')"
    ElseIf element.nodeName = "BODY" And parentSelector <> "" Then
        'clear most recently querySelector & restart with 'body'
        'Debug.Print parentSelector
        element.QuerySelector = left(parentSelector, InStrRev(parentSelector, ".querySelector")) & "querySelector('body')"
        'Debug.Print element.querySelector
        'Stop
    ElseIf element.nodeName = "#DOCUMENT-FRAGMENT" Then
        'shadowRoot
        element.QuerySelector = parentSelector & "').shadowRoot.querySelector('')"
    ElseIf element.nodeName = "#DOCUMENT" And parentSelector <> "" Then
        'likely start of an iFrame
        element.QuerySelector = parentSelector & "').contentDocument.querySelector('')"
    ElseIf parentSelector <> "" And InStr(element.nodeName, "#") = 0 Then
        'Create a querySelector path to the element for jsEval calls, such as clsElement.click method
        If Right(parentSelector, 2) = "('" Then
            'start of a new querySelector
            'parentSelector = ...querySelector('
            element.QuerySelector = parentSelector & element.nodeName
        Else
            'Continuation of a querySelector
            'parentSelector ~ ...querySelector('DIV:nth-child(2) > DIV
            element.QuerySelector = parentSelector & " > " & element.nodeName
        End If
        
        Dim str_id As String: str_id = element.getAttribute("id")
        
        If str_id <> "" Then
            'Shortcut to #id
            element.QuerySelector = left(parentSelector, InStrRev(parentSelector, "'")) & "#" & str_id & "')"
        ElseIf element.n_siblings > 1 And element.nodeName <> "HTML" Then
            'indicate tag + child #
            element.QuerySelector = element.QuerySelector & ":nth-child(" & element.nth_child & ")')"
        Else
            'indicate tag
            element.QuerySelector = element.QuerySelector & "')"
        End If
    End If
    
End Sub
