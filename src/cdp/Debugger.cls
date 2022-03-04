VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpDebugger"
' Debugger
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Breakpoint identifier.
' BreakpointId AS string

' Call frame identifier.
' CallFrameId AS string

' Location in the source code.
' Location AS object

' Location in the source code.
' ScriptPosition AS object

' Location range within one script.
' LocationRange AS object

' JavaScript call frame. Array of call frames form the call stack.
' CallFrame AS object

' Scope description.
' Scope AS object

' Search match for resource.
' SearchMatch AS object

' 
' BreakLocation AS object

' Enum of possible script languages.
' ScriptLanguage AS string
Public Enum ScriptLanguage
    SL_JavaScript
    SL_WebAssembly
End Enum

' Debug symbols available for a wasm script.
' DebugSymbols AS object


Public Function ScriptLanguageToString(ByVal sl As ScriptLanguage) As String
    Dim retVal As String
    Select Case sl
        Case SL_JavaScript
            retVal = "JavaScript"
        Case SL_WebAssembly
            retVal = "WebAssembly"
        Case Else
            Debug.Print "Warning, unknown value " & sl
    End Select
    ScriptLanguageToString = retVal
End Function

Public Function StringToScriptLanguage(ByVal s As String) As ScriptLanguage
    Dim retVal As ScriptLanguage
    Select Case s
        Case "JavaScript"
            retVal = SL_JavaScript
        Case "WebAssembly"
            retVal = SL_WebAssembly
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToScriptLanguage = retVal
End Function


Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Continues execution until specific location is reached.
Public Sub continueToLocation( _
    ByVal location AS Dictionary, _
    Optional ByVal targetCallFrames AS Variant _
)
    ' location: Location Location to continue to.
    ' targetCallFrames: string(optional) 

    Dim params As New Dictionary
    Set params("location") = location
    If Not IsMissing(targetCallFrames) Then params("targetCallFrames") = CStr(targetCallFrames)

    cdp.InvokeMethod "Debugger.continueToLocation", params
End Sub

' Disables debugger for given page.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Debugger.disable", params
End Sub

' Enables debugger for the given page. Clients should not assume that the debugging has been
' enabled until the result for this command is received.
' Returns: debuggerId - Unique identifier of the debugger.
Public Function enable( _
    Optional ByVal maxScriptsCacheSize AS Variant _
) AS string
    ' maxScriptsCacheSize: number(optional) The maximum size in bytes of collected scripts (not referenced by other heap objects)
    '   the debugger can hold. Puts no limit if parameter is omitted.

    Dim params As New Dictionary
    If Not IsMissing(maxScriptsCacheSize) Then params("maxScriptsCacheSize") = CDbl(maxScriptsCacheSize)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.enable", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("debuggerId") Then enable = results("debuggerId")
    End If
End Function

' Evaluates expression on a given call frame.
' Returns:  - 
Public Function evaluateOnCallFrame( _
    ByVal callFrameId AS string, _
    ByVal expression AS string, _
    Optional ByVal objectGroup AS Variant, _
    Optional ByVal includeCommandLineAPI AS Variant, _
    Optional ByVal silent AS Variant, _
    Optional ByVal returnByValue AS Variant, _
    Optional ByVal generatePreview AS Variant, _
    Optional ByVal throwOnSideEffect AS Variant, _
    Optional ByVal timeout AS Variant _
) AS Dictionary
    ' callFrameId: CallFrameId Call frame identifier to evaluate on.
    ' expression: string Expression to evaluate.
    ' objectGroup: string(optional) String object group name to put result into (allows rapid releasing resulting object handles
    '   using `releaseObjectGroup`).
    ' includeCommandLineAPI: boolean(optional) Specifies whether command line API should be available to the evaluated expression, defaults
    '   to false.
    ' silent: boolean(optional) In silent mode exceptions thrown during evaluation are not reported and do not pause
    '   execution. Overrides `setPauseOnException` state.
    ' returnByValue: boolean(optional) Whether the result is expected to be a JSON object that should be sent by value.
    ' generatePreview: boolean(optional) Whether preview should be generated for the result.
    ' throwOnSideEffect: boolean(optional) Whether to throw an exception if side effect cannot be ruled out during evaluation.
    ' timeout: Runtime.TimeDelta(optional) Terminate execution after timing out (number of milliseconds).

    Dim params As New Dictionary
    params("callFrameId") = CStr(callFrameId)
    params("expression") = CStr(expression)
    If Not IsMissing(objectGroup) Then params("objectGroup") = CStr(objectGroup)
    If Not IsMissing(includeCommandLineAPI) Then params("includeCommandLineAPI") = CBool(includeCommandLineAPI)
    If Not IsMissing(silent) Then params("silent") = CBool(silent)
    If Not IsMissing(returnByValue) Then params("returnByValue") = CBool(returnByValue)
    If Not IsMissing(generatePreview) Then params("generatePreview") = CBool(generatePreview)
    If Not IsMissing(throwOnSideEffect) Then params("throwOnSideEffect") = CBool(throwOnSideEffect)
    If Not IsMissing(timeout) Then params("timeout") = CDbl(timeout)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.evaluateOnCallFrame", params)
    If cdp.ErrorCode = 0 Then
        Set evaluateOnCallFrame = results
    End If
End Function

' Returns possible locations for breakpoint. scriptId in start and end range locations should be
' the same.
' Returns: locations - List of the possible breakpoint locations.
Public Function getPossibleBreakpoints( _
    ByVal start AS Dictionary, _
    Optional ByVal pEnd AS Variant, _
    Optional ByVal restrictToFunction AS Variant _
) AS Collection
    ' start: Location Start of range to search possible breakpoint locations in.
    ' pEnd: Location(optional) End of range to search possible breakpoint locations in (excluding). When not specified, end
    '   of scripts is used as end of range.
    ' restrictToFunction: boolean(optional) Only consider locations which are in the same (non-nested) function as start.

    Dim params As New Dictionary
    Set params("start") = start
    If Not IsMissing(pEnd) Then Set params("end") = pEnd
    If Not IsMissing(restrictToFunction) Then params("restrictToFunction") = CBool(restrictToFunction)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.getPossibleBreakpoints", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("locations") Then Set getPossibleBreakpoints = results("locations")
    End If
End Function

' Returns source for the script with given id.
' Returns:  - 
Public Function getScriptSource( _
    ByVal scriptId AS string _
) AS Dictionary
    ' scriptId: Runtime.ScriptId Id of the script to get source for.

    Dim params As New Dictionary
    params("scriptId") = CStr(scriptId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.getScriptSource", params)
    If cdp.ErrorCode = 0 Then
        Set getScriptSource = results
    End If
End Function

' This command is deprecated. Use getScriptSource instead.
' Returns: bytecode - Script source.
Public Function getWasmBytecode( _
    ByVal scriptId AS string _
) AS String
    ' scriptId: Runtime.ScriptId Id of the Wasm script to get source for.

    Dim params As New Dictionary
    params("scriptId") = CStr(scriptId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.getWasmBytecode", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("bytecode") Then getWasmBytecode = results("bytecode")
    End If
End Function

' Returns stack trace with given `stackTraceId`.
' Returns: stackTrace - 
' Experimental
Public Function getStackTrace( _
    ByVal stackTraceId AS Dictionary _
) AS Dictionary
    ' stackTraceId: Runtime.StackTraceId 

    Dim params As New Dictionary
    Set params("stackTraceId") = stackTraceId

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.getStackTrace", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("stackTrace") Then Set getStackTrace = results("stackTrace")
    End If
End Function

' Stops on the next JavaScript statement.
Public Sub pause()
    Dim params As New Dictionary
    cdp.InvokeMethod "Debugger.pause", params
End Sub

' 
' Experimental
Public Sub pauseOnAsyncCall( _
    ByVal parentStackTraceId AS Dictionary _
)
    ' parentStackTraceId: Runtime.StackTraceId Debugger will pause when async call with given stack trace is started.

    Dim params As New Dictionary
    Set params("parentStackTraceId") = parentStackTraceId

    cdp.InvokeMethod "Debugger.pauseOnAsyncCall", params
End Sub

' Removes JavaScript breakpoint.
Public Sub removeBreakpoint( _
    ByVal breakpointId AS string _
)
    ' breakpointId: BreakpointId 

    Dim params As New Dictionary
    params("breakpointId") = CStr(breakpointId)

    cdp.InvokeMethod "Debugger.removeBreakpoint", params
End Sub

' Restarts particular call frame from the beginning.
' Returns:  - 
Public Function restartFrame( _
    ByVal callFrameId AS string _
) AS Dictionary
    ' callFrameId: CallFrameId Call frame identifier to evaluate on.

    Dim params As New Dictionary
    params("callFrameId") = CStr(callFrameId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.restartFrame", params)
    If cdp.ErrorCode = 0 Then
        Set restartFrame = results
    End If
End Function

' Resumes JavaScript execution.
Public Sub resumeDebugger( _
    Optional ByVal terminateOnResume AS Variant _
)
    ' terminateOnResume: boolean(optional) Set to true to terminate execution upon resuming execution. In contrast
    '   to Runtime.terminateExecution, this will allows to execute further
    '   JavaScript (i.e. via evaluation) until execution of the paused code
    '   is actually resumed, at which point termination is triggered.
    '   If execution is currently not paused, this parameter has no effect.

    Dim params As New Dictionary
    If Not IsMissing(terminateOnResume) Then params("terminateOnResume") = CBool(terminateOnResume)

    cdp.InvokeMethod "Debugger.resume", params
End Sub

' Searches for given string in script content.
' Returns: result - List of search matches.
Public Function searchInContent( _
    ByVal scriptId AS string, _
    ByVal query AS string, _
    Optional ByVal caseSensitive AS Variant, _
    Optional ByVal isRegex AS Variant _
) AS Collection
    ' scriptId: Runtime.ScriptId Id of the script to search in.
    ' query: string String to search for.
    ' caseSensitive: boolean(optional) If true, search is case sensitive.
    ' isRegex: boolean(optional) If true, treats string parameter as regex.

    Dim params As New Dictionary
    params("scriptId") = CStr(scriptId)
    params("query") = CStr(query)
    If Not IsMissing(caseSensitive) Then params("caseSensitive") = CBool(caseSensitive)
    If Not IsMissing(isRegex) Then params("isRegex") = CBool(isRegex)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.searchInContent", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("result") Then Set searchInContent = results("result")
    End If
End Function

' Enables or disables async call stacks tracking.
Public Sub setAsyncCallStackDepth( _
    ByVal maxDepth AS Long _
)
    ' maxDepth: integer Maximum depth of async call stacks. Setting to `0` will effectively disable collecting async
    '   call stacks (default).

    Dim params As New Dictionary
    params("maxDepth") = CLng(maxDepth)

    cdp.InvokeMethod "Debugger.setAsyncCallStackDepth", params
End Sub

' Replace previous blackbox patterns with passed ones. Forces backend to skip stepping/pausing in
' scripts with url matching one of the patterns. VM will try to leave blackboxed script by
' performing 'step in' several times, finally resorting to 'step out' if unsuccessful.
' Experimental
Public Sub setBlackboxPatterns( _
    ByVal patterns AS Collection _
)
    ' patterns: array Array of regexps that will be used to check script url for blackbox state.

    Dim params As New Dictionary
    Set params("patterns") = patterns

    cdp.InvokeMethod "Debugger.setBlackboxPatterns", params
End Sub

' Makes backend skip steps in the script in blackboxed ranges. VM will try leave blacklisted
' scripts by performing 'step in' several times, finally resorting to 'step out' if unsuccessful.
' Positions array contains positions where blackbox state is changed. First interval isn't
' blackboxed. Array should be sorted.
' Experimental
Public Sub setBlackboxedRanges( _
    ByVal scriptId AS string, _
    ByVal positions AS Collection _
)
    ' scriptId: Runtime.ScriptId Id of the script.
    ' positions: array 

    Dim params As New Dictionary
    params("scriptId") = CStr(scriptId)
    Set params("positions") = positions

    cdp.InvokeMethod "Debugger.setBlackboxedRanges", params
End Sub

' Sets JavaScript breakpoint at a given location.
' Returns:  - 
Public Function setBreakpoint( _
    ByVal location AS Dictionary, _
    Optional ByVal condition AS Variant _
) AS Dictionary
    ' location: Location Location to set breakpoint in.
    ' condition: string(optional) Expression to use as a breakpoint condition. When specified, debugger will only stop on the
    '   breakpoint if this expression evaluates to true.

    Dim params As New Dictionary
    Set params("location") = location
    If Not IsMissing(condition) Then params("condition") = CStr(condition)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.setBreakpoint", params)
    If cdp.ErrorCode = 0 Then
        Set setBreakpoint = results
    End If
End Function

' Sets instrumentation breakpoint.
' Returns: breakpointId - Id of the created breakpoint for further reference.
Public Function setInstrumentationBreakpoint( _
    ByVal instrumentation AS string _
) AS string
    ' instrumentation: string Instrumentation name.

    Dim params As New Dictionary
    params("instrumentation") = CStr(instrumentation)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.setInstrumentationBreakpoint", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("breakpointId") Then setInstrumentationBreakpoint = results("breakpointId")
    End If
End Function

' Sets JavaScript breakpoint at given location specified either by URL or URL regex. Once this
' command is issued, all existing parsed scripts will have breakpoints resolved and returned in
' `locations` property. Further matching script parsing will result in subsequent
' `breakpointResolved` events issued. This logical breakpoint will survive page reloads.
' Returns:  - 
Public Function setBreakpointByUrl( _
    ByVal lineNumber AS Long, _
    Optional ByVal url AS Variant, _
    Optional ByVal urlRegex AS Variant, _
    Optional ByVal scriptHash AS Variant, _
    Optional ByVal columnNumber AS Variant, _
    Optional ByVal condition AS Variant _
) AS Dictionary
    ' lineNumber: integer Line number to set breakpoint at.
    ' url: string(optional) URL of the resources to set breakpoint on.
    ' urlRegex: string(optional) Regex pattern for the URLs of the resources to set breakpoints on. Either `url` or
    '   `urlRegex` must be specified.
    ' scriptHash: string(optional) Script hash of the resources to set breakpoint on.
    ' columnNumber: integer(optional) Offset in the line to set breakpoint at.
    ' condition: string(optional) Expression to use as a breakpoint condition. When specified, debugger will only stop on the
    '   breakpoint if this expression evaluates to true.

    Dim params As New Dictionary
    params("lineNumber") = CLng(lineNumber)
    If Not IsMissing(url) Then params("url") = CStr(url)
    If Not IsMissing(urlRegex) Then params("urlRegex") = CStr(urlRegex)
    If Not IsMissing(scriptHash) Then params("scriptHash") = CStr(scriptHash)
    If Not IsMissing(columnNumber) Then params("columnNumber") = CLng(columnNumber)
    If Not IsMissing(condition) Then params("condition") = CStr(condition)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.setBreakpointByUrl", params)
    If cdp.ErrorCode = 0 Then
        Set setBreakpointByUrl = results
    End If
End Function

' Sets JavaScript breakpoint before each call to the given function.
' If another function was created from the same source as a given one,
' calling it will also trigger the breakpoint.
' Returns: breakpointId - Id of the created breakpoint for further reference.
' Experimental
Public Function setBreakpointOnFunctionCall( _
    ByVal objectId AS string, _
    Optional ByVal condition AS Variant _
) AS string
    ' objectId: Runtime.RemoteObjectId Function object id.
    ' condition: string(optional) Expression to use as a breakpoint condition. When specified, debugger will
    '   stop on the breakpoint if this expression evaluates to true.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)
    If Not IsMissing(condition) Then params("condition") = CStr(condition)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.setBreakpointOnFunctionCall", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("breakpointId") Then setBreakpointOnFunctionCall = results("breakpointId")
    End If
End Function

' Activates / deactivates all breakpoints on the page.
Public Sub setBreakpointsActive( _
    ByVal active AS boolean _
)
    ' active: boolean New value for breakpoints active state.

    Dim params As New Dictionary
    params("active") = CBool(active)

    cdp.InvokeMethod "Debugger.setBreakpointsActive", params
End Sub

' Defines pause on exceptions state. Can be set to stop on all exceptions, uncaught exceptions or
' no exceptions. Initial pause on exceptions state is `none`.
Public Sub setPauseOnExceptions( _
    ByVal state AS string _
)
    ' state: string Pause on exceptions mode.

    Dim params As New Dictionary
    params("state") = CStr(state)

    cdp.InvokeMethod "Debugger.setPauseOnExceptions", params
End Sub

' Changes return value in top frame. Available only at return break position.
' Experimental
Public Sub setReturnValue( _
    ByVal newValue AS Dictionary _
)
    ' newValue: Runtime.CallArgument New return value.

    Dim params As New Dictionary
    Set params("newValue") = newValue

    cdp.InvokeMethod "Debugger.setReturnValue", params
End Sub

' Edits JavaScript source live.
' Returns:  - 
Public Function setScriptSource( _
    ByVal scriptId AS string, _
    ByVal scriptSource AS string, _
    Optional ByVal dryRun AS Variant _
) AS Dictionary
    ' scriptId: Runtime.ScriptId Id of the script to edit.
    ' scriptSource: string New content of the script.
    ' dryRun: boolean(optional) If true the change will not actually be applied. Dry run may be used to get result
    '   description without actually modifying the code.

    Dim params As New Dictionary
    params("scriptId") = CStr(scriptId)
    params("scriptSource") = CStr(scriptSource)
    If Not IsMissing(dryRun) Then params("dryRun") = CBool(dryRun)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Debugger.setScriptSource", params)
    If cdp.ErrorCode = 0 Then
        Set setScriptSource = results
    End If
End Function

' Makes page not interrupt on any pauses (breakpoint, exception, dom exception etc).
Public Sub setSkipAllPauses( _
    ByVal skip AS boolean _
)
    ' skip: boolean New value for skip pauses state.

    Dim params As New Dictionary
    params("skip") = CBool(skip)

    cdp.InvokeMethod "Debugger.setSkipAllPauses", params
End Sub

' Changes value of variable in a callframe. Object-based scopes are not supported and must be
' mutated manually.
Public Sub setVariableValue( _
    ByVal scopeNumber AS Long, _
    ByVal variableName AS string, _
    ByVal newValue AS Dictionary, _
    ByVal callFrameId AS string _
)
    ' scopeNumber: integer 0-based number of scope as was listed in scope chain. Only 'local', 'closure' and 'catch'
    '   scope types are allowed. Other scopes could be manipulated manually.
    ' variableName: string Variable name.
    ' newValue: Runtime.CallArgument New variable value.
    ' callFrameId: CallFrameId Id of callframe that holds variable.

    Dim params As New Dictionary
    params("scopeNumber") = CLng(scopeNumber)
    params("variableName") = CStr(variableName)
    Set params("newValue") = newValue
    params("callFrameId") = CStr(callFrameId)

    cdp.InvokeMethod "Debugger.setVariableValue", params
End Sub

' Steps into the function call.
Public Sub stepInto( _
    Optional ByVal breakOnAsyncCall AS Variant, _
    Optional ByVal skipList AS Variant _
)
    ' breakOnAsyncCall: boolean(optional) Debugger will pause on the execution of the first async task which was scheduled
    '   before next pause.
    ' skipList: array(optional) The skipList specifies location ranges that should be skipped on step into.

    Dim params As New Dictionary
    If Not IsMissing(breakOnAsyncCall) Then params("breakOnAsyncCall") = CBool(breakOnAsyncCall)
    If Not IsMissing(skipList) Then Set params("skipList") = skipList

    cdp.InvokeMethod "Debugger.stepInto", params
End Sub

' Steps out of the function call.
Public Sub stepOut()
    Dim params As New Dictionary
    cdp.InvokeMethod "Debugger.stepOut", params
End Sub

' Steps over the statement.
Public Sub stepOver( _
    Optional ByVal skipList AS Variant _
)
    ' skipList: array(optional) The skipList specifies location ranges that should be skipped on step over.

    Dim params As New Dictionary
    If Not IsMissing(skipList) Then Set params("skipList") = skipList

    cdp.InvokeMethod "Debugger.stepOver", params
End Sub

