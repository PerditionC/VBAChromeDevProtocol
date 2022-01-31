' to allow convenience form of calling node.fn(...) instead of browser.fn(node,...)
Private browser As AutomateBrowser
Private resolveNode As Dictionary
Private objectProperties As Dictionary

' which method to use to set a node's "value"
Public Enum SetNodeValueMode
    SNV_AsAttribute
    SNV_AsValue
    SNV_AsNodeValue
    SNV_AsFakeInput
End Enum

