' ************* Convenience functions *********************
Public Sub Click(Optional ByVal strategy As NavigationStrategy = NavigationStrategy.None)
    Browser.Click nodeId:=nodeId, backendNodeId:=backendNodeId, strategy:=strategy
End Sub


Public Function querySelector(ByVal selector As String, Optional ByVal frameId As String = vbNullString, _
                              Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False, _
                              Optional waitUntilFound As Boolean = True, Optional waitUntilVisible As Boolean = True) As cdpDOMNode
    Set querySelector = Browser.querySelector(selector, nodeId, frameId, depth, pierce, waitUntilFound, waitUntilVisible)
End Function

Public Function querySelectorAll(ByVal selector As String, Optional ByVal rootId As Long = 0, Optional ByVal frameId As String = vbNullString, _
                                 Optional ByVal depth As Long = -1, Optional ByVal pierce As Boolean = False) As Dictionary ' of cdpDomNode
    Set querySelectorAll = Browser.querySelectorAll(selector, nodeId, frameId, depth, pierce)
End Function


Public Sub focus()
    Browser.focus nodeId:=nodeId, backendNodeId:=backendNodeId
End Sub

Public Sub getCenter(ByRef x As Double, ByRef y As Double)
    Browser.getNodeCenter x, y, nodeId:=nodeId, backendNodeId:=backendNodeId
End Sub


Public Property Get elementValue() As String
    Dim child As cdpDOMNode, v As Variant
    
    ' look for shadowDom for value
    If shadowRoots.Count > 0 Then
        ' we assume textContent is our value
        Set child = shadowRoots(1)
        elementValue = child.textContent
    End If
    
    ' if we didn't find a value (or its "")
    If elementValue = vbNullString Then
        'is the value itself set? use it
        If value <> vbNullString Then
            elementValue = value
        ' try nodeValue, which should correspond with value attribute but directly accessible
        ElseIf nodeValue <> vbNullString Then
            elementValue = nodeValue
        ' not yet then see if there is a child (not shadow) element that is a #Text node, e.g. for textarea like controls
        ElseIf childNodeCount > 0 Then
            For Each v In children
                Set child = v
                If child.nodeType = 3 Then
                    elementValue = child.elementValue
                    Exit Property
                End If
            Next v
        End If
    End If
    
    ' if all else fails, just return value attribute's value
    If elementValue = vbNullString Then
        elementValue = elementAttribute("value")
    End If
End Property

Public Property Let elementValue(ByVal newValue As String)
    setValue newValue, SetNodeValueMode.SNV_Automatic
End Property

Public Sub setValue(ByVal newValue As String, Optional mode As SetNodeValueMode = SetNodeValueMode.SNV_Automatic)
    Dim child As cdpDOMNode, v As Variant
    
    ' first focus the element (its ok if this fails usually)
    If nodeId <> 0 Then
        Browser.cdp.DOM.focus nodeId:=nodeId
        Browser.cdp.DOM.scrollIntoViewIfNeeded nodeId:=nodeId
    Else
        Browser.cdp.DOM.focus backendNodeId:=backendNodeId
        Browser.cdp.DOM.scrollIntoViewIfNeeded backendNodeId:=backendNodeId
    End If
    
    ' select anything currently set - don't do this unless direct input has already been done, as it simulates direct input and that breaks setting the value explicitly
    'Browser.cdp.SimulateInput.dispatchKeyEvent "keyDown", modifiers:=KbModifier.Ctrl, Text:="a", windowsVirtualKeyCode:=65
    'Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp"
    
    ' guess proper mode if automatic selected
    If mode = SetNodeValueMode.SNV_Automatic Then
        If autoSetMode <> SetNodeValueMode.SNV_Automatic Then
            mode = autoSetMode
        ElseIf nodeName = "INPUT" Then
            ' if node is INPUT then assume value stored in attribute
            mode = SetNodeValueMode.SNV_Attribute
        ElseIf childNodeCount > 0 Then
            ' set first Text node found's node value, e.g. assume like a TEXTAREA
            mode = SetNodeValueMode.SNV_NodeValue
        Else
            ' ???
            mode = SetNodeValueMode.SNV_Attribute
        End If
    End If
    
    ' set the value using requested method
    Select Case mode
        Case SetNodeValueMode.SNV_Attribute
            elementAttribute("value") = newValue
        Case SetNodeValueMode.SNV_Value
            Browser.cdp.DOM.setNodeValue nodeId, newValue ' ???
            ' update cached value as well
            'value = newValue
        Case SetNodeValueMode.SNV_NodeValue
            If nodeType = 3 Then ' is this a text node?
                Browser.cdp.DOM.setNodeValue nodeId, newValue
                ' update cached value as well
                'nodeValue = newValue
            Else
                ' if has children see if has Text node (type=3) and directly set its value
                For Each v In children
                    Set child = v
                    If child.nodeType = 3 Then
                        Browser.cdp.DOM.setNodeValue child.nodeId, newValue
                        ' update cached value as well
                        'child.nodeValue = newValue
                        Exit For
                    End If
                Next v
            End If
            Browser.cdp.DOM.setNodeValue nodeId, newValue
        Case SetNodeValueMode.SNV_Clipboard
            autoSetMode = SetNodeValueMode.SNV_Clipboard
            'elementAttribute("value") = vbNullString
            ' put string in clipboard
            SetClipboard newValue
            ' select all current text so replaced with new value, via select all
            Browser.cdp.SimulateInput.dispatchKeyEvent "rawKeyDown", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA", windowsVirtualKeyCode:=65
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA"
            ' issue Ctrl-V to paste
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyDown", modifiers:=KbModifier.Ctrl, code:="KeyV", windowsVirtualKeyCode:=86
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp"
        Case SetNodeValueMode.SNV_FakeInput
            autoSetMode = SetNodeValueMode.SNV_FakeInput
            'elementAttribute("value") = vbNullString
            ' we need to clear node first
            ' select all current text so replaced with new value, via select all
            Browser.cdp.SimulateInput.dispatchKeyEvent "rawKeyDown", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA", windowsVirtualKeyCode:=65
            Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp", modifiers:=KbModifier.Ctrl, Text:="a", unmodifiedText:="a", code:="KeyA", key:="KeyA"
            ' then dispatch string as a set of fake key events
            Dim c As String, ndx As Long
            For ndx = 1 To Len(newValue)
                c = Mid$(newValue, ndx, 1)
                Browser.cdp.SimulateInput.dispatchKeyEvent "keyDown", text:=c
                ' doing all three can cause duplicate characters input, char doesn't always work though, so use keyDown/keyUp
                'Browser.cdp.SimulateInput.dispatchKeyEvent "char", Text:=c
                Browser.cdp.SimulateInput.dispatchKeyEvent "keyUp"
            Next ndx
    End Select
    
    ' to avoid stale data, we need to update all attributes, values, shadowdom, etc
    Dim dict As Dictionary
    If nodeId <> 0 Then
        Set dict = Browser.cdp.DOM.describeNode(nodeId:=nodeId, depth:=-1, pierce:=True)
    ElseIf backendNodeId <> 0 Then
        Set dict = Browser.cdp.DOM.describeNode(backendNodeId:=backendNodeId, depth:=-1, pierce:=True)
    End If
    If Browser.cdp.ErrorCode = 0 Then
        Me.init dict, Browser, nodeId, parentNode
    End If
End Sub


Public Property Get elementAttribute(ByVal name As String) As String
    elementAttribute = Browser.getNodeAttribute(Me, name)
End Property

Public Property Let elementAttribute(ByVal name As String, ByVal value As String)
    Browser.cdp.DOM.setAttributeValue nodeId, name, value
    
    ' update cached value
    If Browser.cdp.ErrorCode = 0 Then
        Dim v As Variant, col As Collection
        Set attributes = New Collection
        Set col = Browser.cdp.DOM.getAttributes(nodeId)
        For Each v In col
            attributes.Add v
        Next v
    End If
End Property


' clears our cache so next call to getProperty (or related) retrieves [updated] property values
Public Sub refreshProperties()
    Set objectProperties = Nothing
End Sub

' returns specified javascript property associated with this node
' see Runtime.getProperties
' Note: this call may take a noticable amount of time, especially when called on multiple nodes with many properties
Public Function getProperty(ByVal propertyName As String) As cdpRuntimePropertyDescriptor
    Dim index As Integer
    Dim objectProperty As Dictionary
    
    If objectProperties Is Nothing Then Set objectProperties = Browser.GetProperties(Me)
    If Not objectProperties Is Nothing Then
        If objectProperties.Exists(propertyName) Then
            Dim prop As cdpRuntimePropertyDescriptor
            Set prop = objectProperties(propertyName)
            Set getProperty = prop
        End If
    End If
End Function

' returns value of specified javascript property associated with this node
' see Runtime.getProperties
' Note: this call may take a noticable amount of time, especially when called on multiple nodes with many properties
Public Function getPropertyValue(ByVal propertyName As String) As String
    Dim index As Integer
    Dim objectProperty As Dictionary
    
    If objectProperties Is Nothing Then Set objectProperties = Browser.GetProperties(Me)
    If Not objectProperties Is Nothing Then
        If objectProperties.Exists(propertyName) Then
            Dim prop As cdpRuntimePropertyDescriptor
            Set prop = objectProperties(propertyName)
            If IsNull(prop.value.value) Then
                getPropertyValue = vbNullString
            Else
                getPropertyValue = CStr(prop.value.value)
            End If
        End If
    End If
End Function

Public Function getIdOrName() As String
    getIdOrName = Browser.getNodeIdOrName(Me)
End Function

Public Property Get id() As String
    id = getPropertyValue("id")
End Property

Public Property Get className() As String
    className = getPropertyValue("className")
End Property

Public Property Get title() As String
    title = getPropertyValue("title")
End Property

Public Property Get textContent() As String
    textContent = getPropertyValue("textContent")
End Property

Public Property Get tagName() As String
    tagName = getPropertyValue("tagName")
End Property

Public Property Get outerText() As String
    outerText = getPropertyValue("outerText")
End Property

Public Property Get InnerText() As String
    InnerText = getPropertyValue("innerText")
End Property

Public Property Get innerHTML() As String
    innerHTML = getPropertyValue("innerHTML")
End Property

Public Property Get outerHTML() As String
    'outerHTML = getPropertyValue("outerHTML")
    outerHTML = Browser.cdp.DOM.getOuterHTML(objectId:=resolveNode.objectId)
End Property


' returns True only if element meets both criteria:
'   has a positive width and height and
'   CSS display property is not "none" or "hidden"
Public Function isVisible() As Boolean
    'isVisible = False
    
    ' first check if has a positive width and height
    Dim x As Double, y As Double
    getCenter x, y
    If (Abs(x) > 0#) And (Abs(y) > 0#) Then
        ' now get computed CSS display value
        Dim style As cdpCSSCSSComputedStyleProperty
        Set style = getCSSStyle("display")
        If Not style Is Nothing Then
            If (style.value <> "none") And (style.value <> "hidden") Then
                isVisible = True
            End If
        End If
    End If
End Function


' retrieves current CSS computed style for specified property
Public Function getCSSStyle(ByVal name As String) As cdpCSSCSSComputedStyleProperty
    Dim styles As Collection
    Browser.cdp.css.enable ' must be enabled or call to computer style will fail
    Set styles = Browser.cdp.css.getComputedStyleForNode(nodeId)
    Browser.cdp.css.disable
    Dim v As Variant
    For Each v In styles
        If v("name") = name Then
            Dim style As cdpCSSCSSComputedStyleProperty
            Set style = New cdpCSSCSSComputedStyleProperty
            style.init v
            Set getCSSStyle = style
            Exit Function
        End If
    Next v
End Function


' returns neighboring child node of parent
' returns Nothing if no next neighbor or parent unknown
Public Function NextSibling() As cdpDOMNode
    If Not parentNode Is Nothing Then
        Dim v As Variant, lastNode As cdpDOMNode
        For Each v In parentNode.children
            Dim node As cdpDOMNode
            Set node = v
            ' was previous child me, if so then return this one
            If lastNode Is Me Then
                Set NextSibling = node
                Exit Function
            End If
            Set lastNode = node
        Next v
    End If
End Function

' returns neighboring child node of parent
' returns Nothing if no previous neighbor or parent unknown
Public Function PrevSibling() As cdpDOMNode
    If Not parentNode Is Nothing Then
        Dim v As Variant, lastNode As cdpDOMNode
        For Each v In parentNode.children
            Dim node As cdpDOMNode
            Set node = v
            ' is current child me, if so then return last one
            If node Is Me Then
                Set PrevSibling = lastNode
                Exit Function
            End If
            Set lastNode = node
        Next v
    End If
End Function
