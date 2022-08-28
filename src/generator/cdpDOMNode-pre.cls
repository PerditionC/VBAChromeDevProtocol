' to allow convenience form of calling node.fn(...) instead of browser.fn(node,...)
Private Browser As AutomateBrowser
' resolve javascript reference to node, only done if browser is provided
Public resolveNode As cdpRuntimeRemoteObject
' cache object properties, initialized on demand not during init as takes time to serialize and deserialize to do automatically
Private objectProperties As Dictionary ' of cdpRuntimePropertyDescriptor
' cache of parent node, to allow traversing next/prev sibling
Private parentNode As cdpDOMNode

' which method to use to set a node's "value"
Public Enum SetNodeValueMode
    SNV_Automatic ' guess best method to set value, may not work correctly
    SNV_Attribute ' e.g. value attribute of an <input> element, i.e. input's value
    SNV_Value     ' ???
    SNV_NodeValue ' e.g. text of a <p> or <textarea> node not input's value
    SNV_FakeInput ' simulated typing out input
    SNV_Clipboard ' simulated copy and paste in value
End Enum

Private autoSetMode As SetNodeValueMode

' modifiers to simulating key input
Private Enum KbModifier
    None = 0
    Alt = 1
    Ctrl = 2
    MetaCommand = 4
    Shift = 8
End Enum

' valid node types
Public Enum NodeTypeValue
    NT_ELEMENT = 1              ' An element node such as <p> or <div>
    NT_ATTRIBUTE                ' An attribute of an element
    NT_TEXT                     ' The textual content (actual "text") inside an element or attribute
    NT_CDATA_SECTION            ' A CDATASection, such as <!CDATA[[ … ]]>
    NT_ENTITY_REFERENCE         ' deprecated / unused
    NT_ENTITY                   ' deprecated / unused
    NT_PROCESSING_INSTRUCTION   ' A processingInstruction of an XML document, such as <?xml-stylesheet … ?>
    NT_COMMENT                  ' A comment node, such as <!-- … -->
    NT_DOCUMENT                 ' A document node, the root node of the DOM tree
    NT_DOCUMENT_TYPE            ' A documentType node, such as <!DOCTYPE html>
    NT_DOCUMENT_FRAGMENT        ' A documentFragment node
    NT_NOTATION                 ' deprecated / unused
End Enum

