Attribute VB_Name = "PrintDomTree"
Option Explicit

Public Sub PrintNode(ByVal node As cdpDOMNode, Optional ByVal pad As String = "")
    Debug.Print pad & "<" & node.nodeName & "(" & node.nodeType & ") " & "name='" & node.localName & "' attributes:" & Concat(node.attributes) & IIf(node.childNodeCount > 0, ">", " />")
    If node.childNodeCount > 0 Then
        Dim v As Variant
        For Each v In node.children
            PrintNode v, pad & "  "
        Next v
        Debug.Print pad & "</" & node.nodeName & ">"
    End If
End Sub

Private Function Concat(ByVal attrib As Collection)
    Dim v As Variant
    Dim key As String, Value As String
    Dim counter As Integer
    If attrib.count = 0 Then Concat = "''"
    For Each v In attrib
        If (counter Mod 2) = 0 Then
            key = CStr(v)
        Else
            Value = "'" & CStr(v) & "'"
            If counter > 1 Then Concat = Concat & ","
            Concat = Concat & key & "=" & Value
        End If
        counter = counter + 1
    Next v
End Function


