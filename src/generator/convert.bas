Attribute VB_Name = "convert"
Option Explicit

Private fOut As Integer
Private types As Dictionary

Private Function FileToString(SourceFile As String) As String
    Dim i As Integer
    Dim s As String

    i = FreeFile
    Open SourceFile For Input As #i
    s = Input(LOF(i), i)
    Close #i

    FileToString = s
End Function


Private Function WriteStr(line As String)
    Print #fOut, line
End Function

Private Sub NewClass(ByVal className As String)
    fOut = FreeFile
    Open ThisWorkbook.Path & "\..\src\cdp\" & className & ".cls" For Output As #fOut
    WriteStr "VERSION 1.0 CLASS"
    WriteStr "BEGIN"
    WriteStr "  MultiUse = -1  'True"
    WriteStr "END"
    className = "cdp" & className ' prefix with cdp instead of the more common cls to group together and differeniate
    If Len(className) > 31 Then className = left(className, 31) ' VBA classes can't exceed 31 characters
    WriteStr "Attribute VB_Name = """ & className & """"
    WriteStr "Attribute VB_GlobalNameSpace = False"
    WriteStr "Attribute VB_Creatable = False"
    WriteStr "Attribute VB_PredeclaredId = False"
    WriteStr "Attribute VB_Exposed = True"
    WriteStr "Attribute VB_Description = """ & className & """"
    ' code visible in VBA editor starts here
End Sub


Private Function OutputTypeInitializers(opt As Boolean, rawName As String, name As String, cType As String, dType As String, cSubType As String, dSubType As String, ByRef initFromDictStatements As Dictionary, ByRef initToDictStatements As Dictionary)
    Dim needsV As Boolean: needsV = False
        Select Case dType
            Case "Collection"
                ' TODO this needs to recursively process each type
                ' for each in collection
                '    if type is blah then do blah
                ' next
                needsV = True
                Dim c As Collection
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "        For each v in obj.Item(""" & rawName & """)"
                    ' based on type add value or convert from collection, currently doesn't support nesting, i.e. a Collections of Collections [array of array]
                    If dSubType = "Dictionary" Then
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "            Dim obj_" & name & " As " & cSubType & ": Set obj_" & name & " = New " & cSubType
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "            obj_" & name & ".init v"
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "            " & name & ".Add obj_" & name
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "            Set obj_" & name & " = Nothing"
                    Else
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "            " & name & ".Add v"
                    End If
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Next v"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    End If"
                    
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    Set dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    For each v in obj.Item(""" & rawName & """)"
                    ' based on type add value or convert from collection, currently doesn't support nesting, i.e. a Collections of Collections [array of array]
                    If dSubType = "Dictionary" Then
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Dim obj_" & name & " As " & cSubType & ": Set obj_" & name & " = New " & cSubType
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        obj_" & name & ".init v"
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        " & name & ".Add obj_" & name & ""
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Set obj_" & name & " = Nothing"
                    Else
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        " & name & ".Add v"
                    End If
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Next v"
                    
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    Set dict(""" & rawName & """) = " & name
                End If
            Case "Dictionary"
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then"
                    
                    If cType = "object" Then
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Set " & name & " = obj.Item(""" & rawName & """)"
                    Else
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Set " & name & " = New " & cType
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        " & name & ".init obj.Item(""" & rawName & """)"
                    End If
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    End If"
                    
                    If cType = "object" Then
                        initToDictStatements.Add CStr(initToDictStatements.Count), "    If Not " & name & " Is Nothing Then Set dict(""" & rawName & """) = " & name
                    Else
                        initToDictStatements.Add CStr(initToDictStatements.Count), "    If Not " & name & " Is Nothing Then Set dict(""" & rawName & """) = " & name & ".asDictionary()"
                    End If
                Else
                    If cType = "object" Then
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Set " & name & " = obj.Item(""" & rawName & """)"
                    Else
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Set " & name & " = New " & cType
                        initFromDictStatements.Add CStr(initFromDictStatements.Count), "    " & name & ".init obj.Item(""" & rawName & """)"
                    End If
                    
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    Set dict(""" & rawName & """) = " & name & ".asDictionary()"
                End If
            Case "boolean"
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then Let " & name & " = CBool(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    if Not IsEmpty(" & name & ") Then dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Let " & name & " = CBool(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                End If
            Case "Long"
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then Let " & name & " = CLng(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Let " & name & " = CLng(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                End If
            Case "Double"
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then Let " & name & " = CDbl(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Let " & name & " = CDbl(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                End If
            Case "string"
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then Let " & name & " = CStr(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    if " & name & " <> vbNullString Then dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Let " & name & " = CStr(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                End If
            Case "String" ' enum
                ' TODO convert to/from actual enum
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then Let " & name & " = CStr(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    if " & name & " <> vbNullString Then dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Let " & name & " = CStr(obj.Item(""" & rawName & """))"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                End If
            Case "Variant" ' any
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "        If IsObject(obj.Item(""" & rawName & """)) Then"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "            Set " & name & " = obj.Item(""" & rawName & """)"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Else"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "            Let " & name & " = obj.Item(""" & rawName & """)"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "        End If"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    End If"
                    
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If IsObject(obj.Item(""" & rawName & """)) Then"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Set " & name & " = obj.Item(""" & rawName & """)"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Else"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "        Let " & name & " = obj.Item(""" & rawName & """)"
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    End If"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                End If
            Case Else
                Stop
                If opt Then
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    If obj.Exists(""" & rawName & """) Then Let " & name & " = obj.Item(""" & rawName & """)"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                Else
                    initFromDictStatements.Add CStr(initFromDictStatements.Count), "    Let " & name & " = obj.Item(""" & rawName & """)"
                    initToDictStatements.Add CStr(initToDictStatements.Count), "    dict(""" & rawName & """) = " & name
                End If
        End Select

    OutputTypeInitializers = needsV
End Function

Private Sub injectFile(ByVal fTypeOut As Integer, ByVal sourceFilename As String)
    Dim fIn As Integer
    fIn = FreeFile
    
    ' attempt to open the file, on error assume doesn't exist and we are done
    On Error Resume Next
    Open sourceFilename For Input Access Read Shared As #fIn
    If Err.Number <> 0 Then GoTo cleanup
    On Error GoTo 0
    
    Dim line As String
    While Not EOF(fIn)
        Line Input #fIn, line
        Print #fTypeOut, line
    Wend
    
cleanup:
    Close fIn
End Sub

Private Sub NewType(ByVal typeName As String, ByVal domainName As String, typeDesc As String, props As Collection)
    'If typeName = "AdFrameStatus" Then Stop
    Dim fTypeOut As Integer
    fTypeOut = FreeFile
    Open ThisWorkbook.Path & "\..\src\cdp\" & domainName & "_" & typeName & ".cls" For Output As #fTypeOut
    Print #fTypeOut, "VERSION 1.0 CLASS"
    Print #fTypeOut, "BEGIN"
    Print #fTypeOut, "  MultiUse = -1  'True"
    Print #fTypeOut, "END"
    Dim className As String
    className = "cdp" & domainName & typeName ' types in form of cdpDomainTypename
    If Len(className) > 31 Then className = left(className, 31) ' VBA classes can't exceed 31 characters
    Print #fTypeOut, "Attribute VB_Name = """ & className & """"
    Print #fTypeOut, "Attribute VB_GlobalNameSpace = False"
    Print #fTypeOut, "Attribute VB_Creatable = False"
    Print #fTypeOut, "Attribute VB_PredeclaredId = False"
    Print #fTypeOut, "Attribute VB_Exposed = True"
    Print #fTypeOut, "Attribute VB_Description = """ & typeName & """"
    ' code visible in VBA editor starts here
    Print #fTypeOut, "' " & domainName & "." & typeName
    Print #fTypeOut, "' " & typeDesc
    Print #fTypeOut, "' This class is automatically generated, please make changes to generator and not this class directly."
    Print #fTypeOut, "Option Explicit"
    Print #fTypeOut, ""
    
    ' insert prefix code if exists for this type
    injectFile fTypeOut, ThisWorkbook.Path & "\..\src\generator\cdp" & domainName & typeName & "-pre.cls"
    
    Dim typeObj As Dictionary
    Dim initStatements As Dictionary: Set initStatements = New Dictionary
    Dim initFromDictStatements As Dictionary: Set initFromDictStatements = New Dictionary
    Dim initToDictStatements As Dictionary: Set initToDictStatements = New Dictionary
    Dim needsV As Boolean: needsV = False
    For Each typeObj In props
        Dim name As String, rawName As String, desc As String, opt As Boolean, ref As String, rType As String, cType As String, dType As String, cSubType As String, dSubType As String
        name = getName(typeObj, domainName, "field")
        rawName = typeObj("name")
        desc = typeObj("description")
        desc = Replace(desc, vbLf, vbLf & "    '   ")
        If typeObj.Exists("optional") Then
            opt = CBool(typeObj("optional"))
        Else
            opt = False ' i.e. required
        End If
        If typeObj.Exists("$ref") Then
            rType = typeObj("$ref")
            Dim refType As Dictionary
            If InStr(1, rType, ".", vbBinaryCompare) > 0 Then
                cType = "cdp" & Replace(rType, ".", "")
            Else
                cType = "cdp" & domainName & rType
            End If
        Else
            rType = typeObj("type")
            cType = typeObj("type")
        End If
        If Len(cType) > 31 Then cType = left(cType, 31) ' max class name in VBA
        dType = getType(typeObj, domainName)
        If dType = "Collection" Then
            Dim itemsDict As Dictionary
            If typeObj.Exists("items") Then
getSubTypeInfo:
                Set itemsDict = typeObj("items")
                If itemsDict.Exists("$ref") Then
                    cSubType = itemsDict("$ref")
                    If InStr(1, cSubType, ".") > 0 Then
                        cSubType = Replace(cSubType, ".", "_")
                    Else
                        cSubType = domainName & "_" & cSubType
                    End If
                    If types.Exists(cSubType) Then
                        Set refType = types(cSubType)
                        If refType.Exists("enum") Then
                            cSubType = "String"
                        Else ' assume object
                            cSubType = "cdp" & Replace(cSubType, "_", "")
                        End If
                    Else
                        If InStr(1, cSubType, ".", vbBinaryCompare) > 0 Then
                            cSubType = "cdp" & Replace(cSubType, ".", "")
                        Else
                            cSubType = "cdp" & domainName & cSubType
                        End If
                    End If
                Else ' assume simple type
                    cSubType = itemsDict("type")
                End If
                dSubType = getType(itemsDict, domainName)
            Else
                cSubType = typeObj("$ref")
                ' if refers to object in current domain then protocol doesn't include domain, so we need to add it
                If Not (InStr(1, cSubType, ".", vbBinaryCompare) > 0) Then
                    cSubType = domainName & "." & cSubType
                End If
                cSubType = Replace(cSubType, ".", "_") ' instead of returning typeType, return full typeObj
                If types.Exists(cSubType) Then
                    Set typeObj = types(cSubType)
                    GoTo getSubTypeInfo
                Else
                    Stop
                    dSubType = vbNullString
                End If
                cSubType = "cdp" & Replace(cSubType, "_", "") ' convert to class name
            End If
            If Len(cSubType) > 31 Then cSubType = left(cSubType, 31) ' max class name in VBA
        Else
            cSubType = vbNullString
            dSubType = vbNullString
        End If
        
        If rType = "array" Then
            Print #fTypeOut, "' " & rawName & " : " & rType & " of " & cSubType
        Else
            Print #fTypeOut, "' " & rawName & " : " & rType
        End If
        If opt Then
            Print #fTypeOut, "' Optional"
        End If
        Print #fTypeOut, "' " & desc
        
        ' to simplify converting Dictionary returned for object to a concrete class and back
        needsV = needsV Or OutputTypeInitializers(opt, rawName, name, cType, dType, cSubType, dSubType, initFromDictStatements, initToDictStatements)
        
        ' we init these objects so don't have to test against Nothing, just an empty collection/dictionary
        If dType = "Collection" Then
            initStatements.Add CStr(initStatements.Count), "    Set " & name & " = New " & dType
        End If
        'If dType = "Dictionary" Then
        '    initStatements.Add CStr(initStatements.Count), "    Set " & name & " = New " & cType
        'End If
        If opt Then
            ' optional booleans we treat as variant so can tell if true, valse, or not set
            ' optional strings are left as vbNullString
            If dType = "boolean" Then
                Print #fTypeOut, "Public " & name & " AS Variant ' " & dType
            ElseIf dType = "Dictionary" Then
                If cType = "object" Then
                    Print #fTypeOut, "Public " & name & " AS " & dType  ' if generic "object" then keep as Dictionary
                Else
                    Print #fTypeOut, "Public " & name & " AS " & cType  ' else convert to specific object type
                End If
            Else
                Print #fTypeOut, "Public " & name & " AS " & dType
            End If
        Else
            If dType = "Dictionary" Then
                If cType = "object" Then
                    Print #fTypeOut, "Public " & name & " AS " & dType  ' if generic "object" then keep as Dictionary
                Else
                    Print #fTypeOut, "Public " & name & " AS " & cType  ' else convert to specific object type
                End If
            Else
                Print #fTypeOut, "Public " & name & " AS " & dType
            End If
        End If
        Print #fTypeOut, ""
    Next typeObj

    Dim v As Variant
    Dim classNameTrunc As String
    classNameTrunc = "cdp" & domainName & typeName
    If Len(classNameTrunc) > 31 Then classNameTrunc = left(classNameTrunc, 31)
    If initFromDictStatements.Count > 0 Then
        Dim optionalInitArgs As String: optionalInitArgs = vbNullString
        If "DOMNode" = domainName & typeName Then
            optionalInitArgs = ", Optional ByVal b As AutomateBrowser"
        End If
        Print #fTypeOut, "Public Function init(ByVal obj as Dictionary" & optionalInitArgs & ") As " & classNameTrunc
        Print #fTypeOut, "Attribute Item.VB_Description = ""Initialize class from Dictionary returned by CDP method."""
        If needsV Then Print #fTypeOut, "    Dim v as Variant"
        Print #fTypeOut, ""
        For Each v In initFromDictStatements.Items
            Print #fTypeOut, CStr(v)
        Next v
        Print #fTypeOut, ""
        
        ' insert extra init code if exists for this type
        injectFile fTypeOut, ThisWorkbook.Path & "\..\src\generator\cdp" & domainName & typeName & "-init.cls"
        
        Print #fTypeOut, "    Set init = Me"
        Print #fTypeOut, "End Function"
        Print #fTypeOut, ""
    End If

    If initToDictStatements.Count > 0 Then
        Print #fTypeOut, "Public Function asDictionary() As Dictionary"
        Print #fTypeOut, "    Dim dict as Dictionary: Set dict = New Dictionary"
        Print #fTypeOut, ""
        For Each v In initToDictStatements.Items
            Print #fTypeOut, CStr(v)
        Next v
        Print #fTypeOut, ""
        Print #fTypeOut, "    Set asDictionary = dict"
        Print #fTypeOut, "End Function"
        Print #fTypeOut, ""
    End If

    If initStatements.Count > 0 Then
        Print #fTypeOut, "Private Sub Class_Initialize()"
        For Each v In initStatements.Items
            Print #fTypeOut, CStr(v)
        Next v
        Print #fTypeOut, "End Sub"
        Print #fTypeOut, ""
    End If
    
    ' insert suffix code if exists for this type
    injectFile fTypeOut, ThisWorkbook.Path & "\..\src\generator\cdp" & domainName & typeName & "-post.cls"
    
    Close #fTypeOut
End Sub

Private Function getType(paramObj As Variant, domainName As String) As String
    Dim ref As String
    If paramObj.Exists("$ref") Then
        ref = paramObj("$ref")
        ' if refers to object in current domain then protocol doesn't include domain, so we need to add it
        If Not (InStr(1, ref, ".", vbBinaryCompare) > 0) Then
            ref = domainName & "." & ref
        End If
        If types.Exists(ref) Then
            ref = types(ref)
        Else
            ' just leave asis and figure it out later
            Stop
            'ref = "Unknown"
        End If
    Else ' assume it is a simple type: boolean, string, ...
        ref = paramObj("type")
    End If
    
    ' convert some types to VBA equivalent
    If ref = "array" Then ref = "Collection"
    If ref = "object" Then ref = "Dictionary"
    If ref = "number" Then ref = "Double"
    If ref = "binary" Then ref = "String" ' base64 encoded data
    If ref = "any" Then ref = "Variant"
    If ref = "integer" Then ref = "Long" ' we need to support 32bit numbers not 16bit
    
    getType = ref
End Function

Private Function pName(ByVal name As String, ByVal prefix As String) As String
    pName = prefix & UCase(left(name, 1)) & Mid(name, 2)
End Function

Private Function getName(ByVal obj As Dictionary, domainName As String, Optional prefix As String = "p") As String
    If obj.Exists("name") Then
        getName = obj("name")
        ' replace names that conflict with VBA ***
        If getName = "close" Then getName = getName & domainName
        If getName = "resume" Then getName = getName & domainName
        If getName = "stop" Then getName = getName & domainName
        
        If getName = "get" Then getName = "getObject"
        If getName = "set" Then getName = "setObject"
        
        If getName = "scale" Then getName = pName(getName, prefix) ' scale (and circle) are undocumented reserved words from BASIC heritage
        If getName = "end" Then getName = pName(getName, prefix)
        If getName = "type" Then getName = pName(getName, prefix)
        If getName = "attribute" Then getName = pName(getName, prefix)
    Else
        Stop ' missing name property
        getName = "NoName"
    End If
End Function

Private Function setParam(rawName As String, cookedName As String, cType As String, opt As Boolean) As String
    If opt Then
        setParam = "    If Not IsMissing(" & cookedName & ") Then params(""" & rawName & """) = " & cType & "(" & cookedName & ")"
    Else
        setParam = "    params(""" & rawName & """) = " & cType & "(" & cookedName & ")"
    End If
End Function


' params is a Collection of parameter information
' paramCnt is how many total parameters are left, used to determine if comma should be omitted
' outputOptional determines if optional or required parameters are printed
Private Function WriteParameter(ByVal params As Collection, paramCnt As Integer, ByVal domainName As String, outputOptional As Boolean) As Integer
    Dim paramObj As Variant
    Dim pName As String, pDesc As String, pType As String, opt As Boolean
    For Each paramObj In params
        pName = getName(paramObj, domainName)
        pDesc = paramObj("description")
        If paramObj.Exists("optional") Then
            opt = CBool(paramObj("optional"))
        Else
            opt = False ' i.e. required
        End If
        pType = getType(paramObj, domainName)
        Dim comma As String
        comma = vbNullString
        If paramCnt > 0 Then comma = ","
        If opt Then
            If outputOptional Then
                WriteStr ("    Optional ByVal " & pName & " AS Variant" & comma & " _") ' optional fields are passed as Variants so we can see if IsMissing() or IsEmpty
                paramCnt = paramCnt - 1
            End If
        Else
            If Not outputOptional Then
                WriteStr ("    ByVal " & pName & " AS " & pType & comma & " _")
                paramCnt = paramCnt - 1
            End If
        End If
    Next paramObj
                
    WriteParameter = paramCnt
End Function


' returns true if command on object
Private Function hasCommand(ByVal cmdName As String, ByVal commands As Collection, ByVal domainName As String) As Boolean
    Dim commandObj As Variant
    For Each commandObj In commands
        Dim name As String
        name = getName(commandObj, domainName)
        If name = cmdName Then
            hasCommand = True
            Exit Function
        End If
    Next commandObj
End Function


' given a name returns just the initials (1st and capital letters)
Private Function getInitials(ByVal name As String) As String
    Dim i As Integer
    For i = 1 To Len(name)
        If i = 1 Then
            getInitials = left(name, 1)
        Else
            Dim c As String
            c = Mid(name, i, 1)
            If c = UCase(c) Then ' only if capital
                getInitials = getInitials & c
            End If
        End If
    Next i
    getInitials = UCase$(getInitials)
End Function

' removes any invalid characters from value
Private Function fixValue(ByVal value As String) As String
    fixValue = Replace(value, "-", vbNullString)
End Function

' output enum and conversion routines
Private Sub outputEnum(ByVal id As String, ByVal typeType As String, ByVal values As Collection, ByRef c As Collection)
    Dim idAbbr As String
    idAbbr = getInitials(id)
    ' MouseButton AS string
    WriteStr "Public Enum " & id
    Dim v As Variant
    For Each v In values
        WriteStr "    " & idAbbr & "_" & fixValue(v)
    Next v
    WriteStr "End Enum"
    
    Dim lidAbbr As String
    lidAbbr = LCase(idAbbr)
    If lidAbbr = "is" Or lidAbbr = "as" Or lidAbbr = "to" Then lidAbbr = "e" ' check for VBA identifiers
    
    c.Add "Public Function " & id & "ToString(ByVal " & lidAbbr & " As " & id & ") As String"
    c.Add "    Dim retVal As String"
    c.Add "    Select Case " & lidAbbr
    For Each v In values
        c.Add "        Case " & idAbbr & "_" & fixValue(v)
        c.Add "            retVal = """ & CStr(v) & """"
    Next v
    c.Add "        Case Else"
    c.Add "            Debug.Print ""Warning, unknown value "" & " & lidAbbr
    c.Add "    End Select"
    c.Add "    " & id & "ToString = retVal"
    c.Add "End Function"
    c.Add ""
    c.Add "Public Function StringTo" & id & "(ByVal s As String) As " & id
    c.Add "    Dim retVal As " & id
    c.Add "    Select Case s"
    For Each v In values
        c.Add "        Case """ & CStr(v) & """"
        c.Add "            retVal = " & idAbbr & "_" & fixValue(v)
    Next v
    c.Add "        Case Else"
    c.Add "            Debug.Print ""Warning, unknown value "" & s"
    c.Add "    End Select"
    c.Add "    StringTo" & id & " = retVal"
    c.Add "End Function"
    c.Add ""
    c.Add ""
End Sub


' ***** run me to generate cdp* class files from protocol JSON
Sub convert()
    Dim content As String
    content = FileToString(ThisWorkbook.Path & "\..\src\generator\protocol.txt")
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(content)
    'Stop
    Debug.Print "protocol loaded, generating class structure"
    
    Dim v As Variant, v2 As Variant
    Set v = o("domains")
        
    Dim domain As Variant, typeObj As Variant
    Dim id As String, desc As String, typeType As String, name As String, opt As Boolean, experimental As Boolean, domainName As String
    ' get all types 1st
    Set types = New Dictionary
    For Each domain In v
        If domain.Exists("types") Then
            For Each typeObj In domain("types")
                id = typeObj("id")
                typeType = typeObj("type")
                'enumValues = typeObj("enum")
                types.Add CStr(domain("domain") & "_" & id), typeObj
                types.Add CStr(domain("domain") & "." & id), typeType  ' quick way to get type
            Next typeObj
        End If
    Next domain
    
    ' now generate class files
    For Each domain In v
        domainName = domain("domain")
        NewClass domainName
        If CBool(domain("experimental")) Then
            WriteStr ("' " & domainName & " [Experimental]")
        Else
            WriteStr ("' " & domainName)
        End If
        WriteStr ("' This class is automatically generated, please make changes to generator and not this class directly.")
        WriteStr ("Option Explicit")
        WriteStr ("")
        WriteStr ("Private cdp As clsCDP")
        WriteStr ("")
        #If False Then ' remove this as this does not correctly track across sessions
        If hasCommand("enable", domain("commands"), domainName) Or hasCommand("disable", domain("commands"), domainName) Then
            WriteStr ("' keeps track of event notification state, defaults to disable")
            WriteStr ("Public isEnabled As Boolean")
        End If
        #End If
        WriteStr ("")
        WriteStr ("")
        
        WriteStr ("' *** Types:")
        Dim c_enumConverters As Collection: Set c_enumConverters = New Collection
        If domain.Exists("types") Then
            For Each typeObj In domain("types")
                id = typeObj("id")
                desc = typeObj("description")
                desc = Replace(desc, vbLf, vbLf & "'   ")
                typeType = typeObj("type")
                'enumValues = typeObj("enum")
                WriteStr ("' " & desc)
                WriteStr ("' " & id & " AS " & typeType)
                
                ' output enum converters
                If typeObj.Exists("enum") Then
                    outputEnum id, typeType, typeObj("enum"), c_enumConverters
                End If
                
                ' generate wrapper objects for "object" types
                If typeType = "object" Then
                    If typeObj.Exists("properties") Then ' if object lacks any properties then don't create an empty object
                        NewType id, domainName, desc, typeObj.Item("properties")
                    End If
                End If
            
                WriteStr ("")
            Next typeObj
        Else
            WriteStr ("' None")
        End If
        Dim tmp As Variant
        If c_enumConverters.Count > 0 Then
            WriteStr ("")
            For Each tmp In c_enumConverters
                WriteStr CStr(tmp)
            Next
        Else
            WriteStr ("")
            WriteStr ("")
        End If
        WriteStr ("Public Sub init(ByRef cdpObj As clsCDP)")
        WriteStr ("    Set cdp = cdpObj")
        WriteStr ("End Sub")
        WriteStr ("")
        WriteStr ("'Private Sub Class_Initialize()")
        WriteStr ("    ' add any needed initialization logic here")
        WriteStr ("'End Sub")
        WriteStr ("")
        WriteStr ("Private Sub Class_Terminate()")
        WriteStr ("    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance")
        WriteStr ("    Set cdp = Nothing")
        WriteStr ("End Sub")
        WriteStr ("")
        WriteStr ("")
        WriteStr ("' *** Commands:")
        WriteStr ("")
        Dim commandObj As Variant
        For Each commandObj In domain("commands")
            name = getName(commandObj, domainName)
            desc = commandObj("description")
            desc = Replace(desc, vbLf, vbLf & "' ") ' handle descriptions that span multiple lines
            Dim rName As String, rDesc As String, rType As String, IsFunction As Boolean, SubOrFunction As String, rCount As Integer
            If commandObj.Exists("returns") Then
                IsFunction = True: SubOrFunction = "Function"
                Dim retObj As Variant
                Set retObj = commandObj("returns") ' returns a Collection of N items, each item a Dictionary describing the returned information
                ' *** some commands, e.g. getWindowForTarget return a Collection of multiple items, if only one then we unwrap it
                rCount = retObj.Count
                If rCount > 1 Then
                    rName = vbNullString
                    rDesc = vbNullString
                    rType = "Dictionary"
                Else
                    Set retObj = retObj.Item(1)
                    rName = retObj("name")
                    rDesc = retObj("description")
                    rDesc = Replace(rDesc, vbLf, vbLf & " '   ")
                    rType = getType(retObj, domainName)
                    ' arrays also have "items" for what the array contains
                End If
            Else
                IsFunction = False: SubOrFunction = "Sub"
                rCount = 0
                rName = vbNullString
                rDesc = vbNullString
                rType = vbNullString
            End If
            If commandObj.Exists("experimental") Then
                experimental = CBool(commandObj("experimental"))
            Else
                experimental = False
            End If
            WriteStr ("' " & desc)
            If IsFunction Then WriteStr ("' Returns: " & rName & " - " & rDesc)
            If experimental Then WriteStr ("' Experimental")
            If commandObj.Exists("parameters") Then
                WriteStr ("Public " & SubOrFunction & " " & name & "( _")
                        
                Dim paramObj As Variant, paramCnt As Integer
                Dim pName As String, pDesc As String, pType As String
                paramCnt = commandObj("parameters").Count - 1
                ' first all non-Optional parameters
                paramCnt = WriteParameter(commandObj("parameters"), paramCnt, domainName, outputOptional:=False)
                ' now all the optional ones
                Call WriteParameter(commandObj("parameters"), paramCnt, domainName, outputOptional:=True)
                
                ' return appropriate type or for Sub then no return
                ' result InvokeMethod always return a dictionary with "result" or "error" field depending on success or failure of command along with "session"
                ' which we unwrap prior to returning
                If IsFunction Then
                    WriteStr (") AS " & rType) ' was Dictionary, now unwrapped type
                Else
                    WriteStr (")")
                End If
            Else
                If IsFunction Then
                    WriteStr ("Public " & SubOrFunction & " " & name & "() AS " & rType)
                Else
                    WriteStr ("Public " & SubOrFunction & " " & name & "()")
                End If
            End If
                If commandObj.Exists("parameters") Then
                For Each paramObj In commandObj("parameters")
                    pName = getName(paramObj, domainName)
                    pDesc = paramObj("description")
                    pDesc = Replace(pDesc, vbLf, vbLf & "    '   ")
                    If paramObj.Exists("optional") Then
                        opt = CBool(paramObj("optional"))
                    Else
                        opt = False ' i.e. required
                    End If
                    ' get type but as described in protocol for comment
                    If paramObj.Exists("$ref") Then
                        pType = paramObj("$ref")
                    Else
                        pType = paramObj("type")
                    End If
                    Dim optS As String
                    optS = vbNullString
                    If opt Then optS = "(optional)"
                    WriteStr ("    ' " & pName & ": " & pType & optS & " " & pDesc)
                Next paramObj
                    WriteStr ("")
                End If
            WriteStr ("    Dim params As New Dictionary")
            If commandObj.Exists("parameters") Then
                'WriteStr ("")
                For Each paramObj In commandObj("parameters")
                    pName = getName(paramObj, domainName)
                    If paramObj.Exists("optional") Then
                        opt = CBool(paramObj("optional"))
                    Else
                        opt = False ' i.e. required
                    End If
                    pType = getType(paramObj, domainName)
                    
                    Select Case pType ' note we have already converted pType to VBA types compatible value
                        Case "string"
                            WriteStr (setParam(paramObj.Item("name"), pName, "CStr", opt))
                        Case "String" ' note this is binary value encoded as base64, not a plain 'string'
                            WriteStr (setParam(paramObj.Item("name"), pName, "CStr", opt))
                        Case "Long"
                            WriteStr (setParam(paramObj.Item("name"), pName, "CLng", opt))
                        Case "Double"
                            WriteStr (setParam(paramObj.Item("name"), pName, "CDbl", opt))
                        Case "boolean"
                            WriteStr (setParam(paramObj.Item("name"), pName, "CBool", opt))
                        Case "Collection", "Dictionary" ' pass directly any array / object as Collection / Dictionary and converted to JSON later in send process
                            ' is this correct for Collection, works for Dictionary
                            If opt Then
                                WriteStr ("    If Not IsMissing(" & pName & ") Then Set params(""" & paramObj.Item("name") & """) = " & pName)
                            Else
                                WriteStr ("    Set params(""" & paramObj.Item("name") & """) = " & pName)
                            End If
                        Case Else
                            Stop ' we don't know what this is? or how to handle it!
                    End Select
                Next paramObj
                WriteStr ("")
            Else
                'WriteStr ("    ' no parameters")
            End If
            If IsFunction Then
                WriteStr ("    Dim results as Dictionary")
                ' we need command name unmodified to pass to CDP, so get it directly, but function name remains adjusted
                WriteStr ("    Set results = cdp.InvokeMethod(""" & domainName & "." & commandObj.Item("name") & """, params)")
                WriteStr ("    If cdp.ErrorCode = 0 Then")
                Dim needSet As String
                If (rType = "Collection") Or (rType = "Dictionary") Then ' array or object returned
                    needSet = "Set "
                Else ' assume direct type, e.g. string, integer, ...
                    needSet = vbNullString
                End If
                If rCount = 1 Then ' unwrap and return directly
                    WriteStr ("        If results.Exists(""" & rName & """) Then " & needSet & name & " = results(""" & rName & """)")
                Else ' rCount > 1, return dictonary with all returned items
                    WriteStr ("        Set " & name & " = results")
                End If
                'WriteStr ("    'Else return default value for type, e.g. 0, """", Nothing, ...")
                WriteStr ("    End If")
                WriteStr ("End Function")
            Else ' Sub so don't expect anything returned from InvokeMethod
                WriteStr ("    cdp.InvokeMethod """ & domainName & "." & commandObj.Item("name") & """, params")
                ' custom handling
                #If False Then ' remove this as this does not correctly track across sessions
                If name = "enable" Then
                    WriteStr ("    If cdp.ErrorCode = 0 Then isEnabled = True")
                ElseIf name = "disable" Then
                    WriteStr ("    If cdp.ErrorCode = 0 Then isEnabled = False")
                End If
                #End If
                WriteStr ("End Sub")
            End If
            
            WriteStr ("")
        Next commandObj
        
        Close #fOut
    Next domain

End Sub
