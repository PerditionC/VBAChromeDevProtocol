' ************* Convenience functions *********************
Public Sub click(Optional ByVal strategy As NavigationStrategy = NavigationStrategy.none)
    browser.click nodeId:=nodeId, backendNodeId:=backendNodeId, strategy:=strategy
End Sub

Public Sub getCenter(ByRef x As Double, ByRef y As Double)
    browser.getNodeCenter x, y, nodeId:=nodeId, backendNodeId:=backendNodeId
End Sub


Public Property Get elementValue() As String
    If value <> vbNullString Then
        elementValue = value
    ElseIf nodeValue <> vbNullString Then
        elementValue = nodeValue
    Else
        elementValue = elementAttribute("value")
    End If
End Property

Public Property Let elementValue(ByVal newvalue As String)
    setValue newvalue, SetNodeValueMode.SNV_AsAttribute
End Property

Public Function setValue(ByVal newvalue As String, Optional mode As SetNodeValueMode = SetNodeValueMode.SNV_AsAttribute)
    Select Case mode
        Case SetNodeValueMode.SNV_AsAttribute
            elementAttribute("value") = newvalue
        Case SetNodeValueMode.SNV_AsValue
            value = newvalue
        Case SetNodeValueMode.SNV_AsNodeValue
            nodeValue = newvalue
        Case SetNodeValueMode.SNV_AsFakeInput
            browser.cdp.SimulateInput.dispatchKeyEvent newvalue
    End Select
End Function


Public Property Get elementAttribute(ByVal name As String) As String
    elementAttribute = browser.getNodeAttribute(Me, name)
End Property

Public Property Let elementAttribute(ByVal name As String, ByVal value As String)
    browser.cdp.DOM.setAttributeValue nodeId, name, value
End Property


'nodeId >> resolveNode >> objectId >> objectProperties
'see Runtime.getProperties
Private Function getProperty(ByVal propertyName As String) As String
    Dim index As Integer
    Dim objectProperty As Dictionary
    
    For index = 1 To objectProperties("result").count
        Set objectProperty = objectProperties("result")(index)
        
        If objectProperty("name") = propertyName Then
            getProperty = objectProperty("value")("value")
            Exit For
        End If
    Next index
End Function

Public Function getIdOrName() As String
    getIdOrName = browser.getNodeIdOrName(Me)
End Function

Public Property Get id() As String
    id = getProperty("id")
End Property

Public Property Get className() As String
    className = getProperty("className")
End Property

Public Property Get title() As String
    title = getProperty("title")
End Property

Public Property Get textContent() As String
    textContent = getProperty("textContent")
End Property

Public Property Get tagName() As String
    tagName = getProperty("tagName")
End Property

Public Property Get outerText() As String
    outerText = getProperty("outerText")
End Property

Public Property Get innerText() As String
    innerText = getProperty("innerText")
End Property

Public Property Get innerHTML() As String
    innerHTML = getProperty("innerHTML")
End Property

Public Property Get outerHTML() As String
    outerHTML = getProperty("outerHTML")
End Property


'Parent
'NextSibling
'PrevSibling
