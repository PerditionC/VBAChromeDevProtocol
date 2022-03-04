VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpRuntime"
' Runtime
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Unique script identifier.
' ScriptId AS string

' Unique object identifier.
' RemoteObjectId AS string

' Primitive value which cannot be JSON-stringified. Includes values `-0`, `NaN`, `Infinity`,
'   `-Infinity`, and bigint literals.
' UnserializableValue AS string

' Mirror object referencing original JavaScript object.
' RemoteObject AS object

' 
' CustomPreview AS object

' Object containing abbreviated remote object value.
' ObjectPreview AS object

' 
' PropertyPreview AS object

' 
' EntryPreview AS object

' Object property descriptor.
' PropertyDescriptor AS object

' Object internal property descriptor. This property isn't normally visible in JavaScript code.
' InternalPropertyDescriptor AS object

' Object private field descriptor.
' PrivatePropertyDescriptor AS object

' Represents function call argument. Either remote object id `objectId`, primitive `value`,
'   unserializable primitive value or neither of (for undefined) them should be specified.
' CallArgument AS object

' Id of an execution context.
' ExecutionContextId AS integer

' Description of an isolated world.
' ExecutionContextDescription AS object

' Detailed information about exception (or error) that was thrown during script compilation or
'   execution.
' ExceptionDetails AS object

' Number of milliseconds since epoch.
' Timestamp AS number

' Number of milliseconds.
' TimeDelta AS number

' Stack entry for runtime errors and assertions.
' CallFrame AS object

' Call frames for assertions or error messages.
' StackTrace AS object

' Unique identifier of current debugger.
' UniqueDebuggerId AS string

' If `debuggerId` is set stack trace comes from another debugger and can be resolved there. This
'   allows to track cross-debugger calls. See `Runtime.StackTrace` and `Debugger.paused` for usages.
' StackTraceId AS object



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

' Add handler to promise with given promise object id.
' Returns:  - 
Public Function awaitPromise( _
    ByVal promiseObjectId AS string, _
    Optional ByVal returnByValue AS Variant, _
    Optional ByVal generatePreview AS Variant _
) AS Dictionary
    ' promiseObjectId: RemoteObjectId Identifier of the promise.
    ' returnByValue: boolean(optional) Whether the result is expected to be a JSON object that should be sent by value.
    ' generatePreview: boolean(optional) Whether preview should be generated for the result.

    Dim params As New Dictionary
    params("promiseObjectId") = CStr(promiseObjectId)
    If Not IsMissing(returnByValue) Then params("returnByValue") = CBool(returnByValue)
    If Not IsMissing(generatePreview) Then params("generatePreview") = CBool(generatePreview)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.awaitPromise", params)
    If cdp.ErrorCode = 0 Then
        Set awaitPromise = results
    End If
End Function

' Calls function with given declaration on the given object. Object group of the result is
' inherited from the target object.
' Returns:  - 
Public Function callFunctionOn( _
    ByVal functionDeclaration AS string, _
    Optional ByVal objectId AS Variant, _
    Optional ByVal arguments AS Variant, _
    Optional ByVal silent AS Variant, _
    Optional ByVal returnByValue AS Variant, _
    Optional ByVal generatePreview AS Variant, _
    Optional ByVal userGesture AS Variant, _
    Optional ByVal awaitPromise AS Variant, _
    Optional ByVal executionContextId AS Variant, _
    Optional ByVal objectGroup AS Variant, _
    Optional ByVal throwOnSideEffect AS Variant _
) AS Dictionary
    ' functionDeclaration: string Declaration of the function to call.
    ' objectId: RemoteObjectId(optional) Identifier of the object to call function on. Either objectId or executionContextId should
    '   be specified.
    ' arguments: array(optional) Call arguments. All call arguments must belong to the same JavaScript world as the target
    '   object.
    ' silent: boolean(optional) In silent mode exceptions thrown during evaluation are not reported and do not pause
    '   execution. Overrides `setPauseOnException` state.
    ' returnByValue: boolean(optional) Whether the result is expected to be a JSON object which should be sent by value.
    ' generatePreview: boolean(optional) Whether preview should be generated for the result.
    ' userGesture: boolean(optional) Whether execution should be treated as initiated by user in the UI.
    ' awaitPromise: boolean(optional) Whether execution should `await` for resulting value and return once awaited promise is
    '   resolved.
    ' executionContextId: ExecutionContextId(optional) Specifies execution context which global object will be used to call function on. Either
    '   executionContextId or objectId should be specified.
    ' objectGroup: string(optional) Symbolic group name that can be used to release multiple objects. If objectGroup is not
    '   specified and objectId is, objectGroup will be inherited from object.
    ' throwOnSideEffect: boolean(optional) Whether to throw an exception if side effect cannot be ruled out during evaluation.

    Dim params As New Dictionary
    params("functionDeclaration") = CStr(functionDeclaration)
    If Not IsMissing(objectId) Then params("objectId") = CStr(objectId)
    If Not IsMissing(arguments) Then Set params("arguments") = arguments
    If Not IsMissing(silent) Then params("silent") = CBool(silent)
    If Not IsMissing(returnByValue) Then params("returnByValue") = CBool(returnByValue)
    If Not IsMissing(generatePreview) Then params("generatePreview") = CBool(generatePreview)
    If Not IsMissing(userGesture) Then params("userGesture") = CBool(userGesture)
    If Not IsMissing(awaitPromise) Then params("awaitPromise") = CBool(awaitPromise)
    If Not IsMissing(executionContextId) Then params("executionContextId") = CLng(executionContextId)
    If Not IsMissing(objectGroup) Then params("objectGroup") = CStr(objectGroup)
    If Not IsMissing(throwOnSideEffect) Then params("throwOnSideEffect") = CBool(throwOnSideEffect)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.callFunctionOn", params)
    If cdp.ErrorCode = 0 Then
        Set callFunctionOn = results
    End If
End Function

' Compiles expression.
' Returns:  - 
Public Function compileScript( _
    ByVal expression AS string, _
    ByVal sourceURL AS string, _
    ByVal persistScript AS boolean, _
    Optional ByVal executionContextId AS Variant _
) AS Dictionary
    ' expression: string Expression to compile.
    ' sourceURL: string Source url to be set for the script.
    ' persistScript: boolean Specifies whether the compiled script should be persisted.
    ' executionContextId: ExecutionContextId(optional) Specifies in which execution context to perform script run. If the parameter is omitted the
    '   evaluation will be performed in the context of the inspected page.

    Dim params As New Dictionary
    params("expression") = CStr(expression)
    params("sourceURL") = CStr(sourceURL)
    params("persistScript") = CBool(persistScript)
    If Not IsMissing(executionContextId) Then params("executionContextId") = CLng(executionContextId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.compileScript", params)
    If cdp.ErrorCode = 0 Then
        Set compileScript = results
    End If
End Function

' Disables reporting of execution contexts creation.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Runtime.disable", params
End Sub

' Discards collected exceptions and console API calls.
Public Sub discardConsoleEntries()
    Dim params As New Dictionary
    cdp.InvokeMethod "Runtime.discardConsoleEntries", params
End Sub

' Enables reporting of execution contexts creation by means of `executionContextCreated` event.
' When the reporting gets enabled the event will be sent immediately for each existing execution
' context.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Runtime.enable", params
End Sub

' Evaluates expression on global object.
' Returns:  - 
Public Function evaluate( _
    ByVal expression AS string, _
    Optional ByVal objectGroup AS Variant, _
    Optional ByVal includeCommandLineAPI AS Variant, _
    Optional ByVal silent AS Variant, _
    Optional ByVal contextId AS Variant, _
    Optional ByVal returnByValue AS Variant, _
    Optional ByVal generatePreview AS Variant, _
    Optional ByVal userGesture AS Variant, _
    Optional ByVal awaitPromise AS Variant, _
    Optional ByVal throwOnSideEffect AS Variant, _
    Optional ByVal timeout AS Variant, _
    Optional ByVal disableBreaks AS Variant, _
    Optional ByVal replMode AS Variant, _
    Optional ByVal allowUnsafeEvalBlockedByCSP AS Variant, _
    Optional ByVal uniqueContextId AS Variant _
) AS Dictionary
    ' expression: string Expression to evaluate.
    ' objectGroup: string(optional) Symbolic group name that can be used to release multiple objects.
    ' includeCommandLineAPI: boolean(optional) Determines whether Command Line API should be available during the evaluation.
    ' silent: boolean(optional) In silent mode exceptions thrown during evaluation are not reported and do not pause
    '   execution. Overrides `setPauseOnException` state.
    ' contextId: ExecutionContextId(optional) Specifies in which execution context to perform evaluation. If the parameter is omitted the
    '   evaluation will be performed in the context of the inspected page.
    '   This is mutually exclusive with `uniqueContextId`, which offers an
    '   alternative way to identify the execution context that is more reliable
    '   in a multi-process environment.
    ' returnByValue: boolean(optional) Whether the result is expected to be a JSON object that should be sent by value.
    ' generatePreview: boolean(optional) Whether preview should be generated for the result.
    ' userGesture: boolean(optional) Whether execution should be treated as initiated by user in the UI.
    ' awaitPromise: boolean(optional) Whether execution should `await` for resulting value and return once awaited promise is
    '   resolved.
    ' throwOnSideEffect: boolean(optional) Whether to throw an exception if side effect cannot be ruled out during evaluation.
    '   This implies `disableBreaks` below.
    ' timeout: TimeDelta(optional) Terminate execution after timing out (number of milliseconds).
    ' disableBreaks: boolean(optional) Disable breakpoints during execution.
    ' replMode: boolean(optional) Setting this flag to true enables `let` re-declaration and top-level `await`.
    '   Note that `let` variables can only be re-declared if they originate from
    '   `replMode` themselves.
    ' allowUnsafeEvalBlockedByCSP: boolean(optional) The Content Security Policy (CSP) for the target might block 'unsafe-eval'
    '   which includes eval(), Function(), setTimeout() and setInterval()
    '   when called with non-callable arguments. This flag bypasses CSP for this
    '   evaluation and allows unsafe-eval. Defaults to true.
    ' uniqueContextId: string(optional) An alternative way to specify the execution context to evaluate in.
    '   Compared to contextId that may be reused across processes, this is guaranteed to be
    '   system-unique, so it can be used to prevent accidental evaluation of the expression
    '   in context different than intended (e.g. as a result of navigation across process
    '   boundaries).
    '   This is mutually exclusive with `contextId`.

    Dim params As New Dictionary
    params("expression") = CStr(expression)
    If Not IsMissing(objectGroup) Then params("objectGroup") = CStr(objectGroup)
    If Not IsMissing(includeCommandLineAPI) Then params("includeCommandLineAPI") = CBool(includeCommandLineAPI)
    If Not IsMissing(silent) Then params("silent") = CBool(silent)
    If Not IsMissing(contextId) Then params("contextId") = CLng(contextId)
    If Not IsMissing(returnByValue) Then params("returnByValue") = CBool(returnByValue)
    If Not IsMissing(generatePreview) Then params("generatePreview") = CBool(generatePreview)
    If Not IsMissing(userGesture) Then params("userGesture") = CBool(userGesture)
    If Not IsMissing(awaitPromise) Then params("awaitPromise") = CBool(awaitPromise)
    If Not IsMissing(throwOnSideEffect) Then params("throwOnSideEffect") = CBool(throwOnSideEffect)
    If Not IsMissing(timeout) Then params("timeout") = CDbl(timeout)
    If Not IsMissing(disableBreaks) Then params("disableBreaks") = CBool(disableBreaks)
    If Not IsMissing(replMode) Then params("replMode") = CBool(replMode)
    If Not IsMissing(allowUnsafeEvalBlockedByCSP) Then params("allowUnsafeEvalBlockedByCSP") = CBool(allowUnsafeEvalBlockedByCSP)
    If Not IsMissing(uniqueContextId) Then params("uniqueContextId") = CStr(uniqueContextId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.evaluate", params)
    If cdp.ErrorCode = 0 Then
        Set evaluate = results
    End If
End Function

' Returns the isolate id.
' Returns: id - The isolate id.
' Experimental
Public Function getIsolateId() AS string
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.getIsolateId", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("id") Then getIsolateId = results("id")
    End If
End Function

' Returns the JavaScript heap usage.
' It is the total usage of the corresponding isolate not scoped to a particular Runtime.
' Returns:  - 
' Experimental
Public Function getHeapUsage() AS Dictionary
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.getHeapUsage", params)
    If cdp.ErrorCode = 0 Then
        Set getHeapUsage = results
    End If
End Function

' Returns properties of a given object. Object group of the result is inherited from the target
' object.
' Returns:  - 
Public Function getProperties( _
    ByVal objectId AS string, _
    Optional ByVal ownProperties AS Variant, _
    Optional ByVal accessorPropertiesOnly AS Variant, _
    Optional ByVal generatePreview AS Variant, _
    Optional ByVal nonIndexedPropertiesOnly AS Variant _
) AS Dictionary
    ' objectId: RemoteObjectId Identifier of the object to return properties for.
    ' ownProperties: boolean(optional) If true, returns properties belonging only to the element itself, not to its prototype
    '   chain.
    ' accessorPropertiesOnly: boolean(optional) If true, returns accessor properties (with getter/setter) only; internal properties are not
    '   returned either.
    ' generatePreview: boolean(optional) Whether preview should be generated for the results.
    ' nonIndexedPropertiesOnly: boolean(optional) If true, returns non-indexed properties only.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)
    If Not IsMissing(ownProperties) Then params("ownProperties") = CBool(ownProperties)
    If Not IsMissing(accessorPropertiesOnly) Then params("accessorPropertiesOnly") = CBool(accessorPropertiesOnly)
    If Not IsMissing(generatePreview) Then params("generatePreview") = CBool(generatePreview)
    If Not IsMissing(nonIndexedPropertiesOnly) Then params("nonIndexedPropertiesOnly") = CBool(nonIndexedPropertiesOnly)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.getProperties", params)
    If cdp.ErrorCode = 0 Then
        Set getProperties = results
    End If
End Function

' Returns all let, const and class variables from global scope.
' Returns: names - 
Public Function globalLexicalScopeNames( _
    Optional ByVal executionContextId AS Variant _
) AS Collection
    ' executionContextId: ExecutionContextId(optional) Specifies in which execution context to lookup global scope variables.

    Dim params As New Dictionary
    If Not IsMissing(executionContextId) Then params("executionContextId") = CLng(executionContextId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.globalLexicalScopeNames", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("names") Then Set globalLexicalScopeNames = results("names")
    End If
End Function

' 
' Returns: objects - Array with objects.
Public Function queryObjects( _
    ByVal prototypeObjectId AS string, _
    Optional ByVal objectGroup AS Variant _
) AS Dictionary
    ' prototypeObjectId: RemoteObjectId Identifier of the prototype to return objects for.
    ' objectGroup: string(optional) Symbolic group name that can be used to release the results.

    Dim params As New Dictionary
    params("prototypeObjectId") = CStr(prototypeObjectId)
    If Not IsMissing(objectGroup) Then params("objectGroup") = CStr(objectGroup)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.queryObjects", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("objects") Then Set queryObjects = results("objects")
    End If
End Function

' Releases remote object with given id.
Public Sub releaseObject( _
    ByVal objectId AS string _
)
    ' objectId: RemoteObjectId Identifier of the object to release.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)

    cdp.InvokeMethod "Runtime.releaseObject", params
End Sub

' Releases all remote objects that belong to a given group.
Public Sub releaseObjectGroup( _
    ByVal objectGroup AS string _
)
    ' objectGroup: string Symbolic object group name.

    Dim params As New Dictionary
    params("objectGroup") = CStr(objectGroup)

    cdp.InvokeMethod "Runtime.releaseObjectGroup", params
End Sub

' Tells inspected instance to run if it was waiting for debugger to attach.
Public Sub runIfWaitingForDebugger()
    Dim params As New Dictionary
    cdp.InvokeMethod "Runtime.runIfWaitingForDebugger", params
End Sub

' Runs script with given id in a given context.
' Returns:  - 
Public Function runScript( _
    ByVal scriptId AS string, _
    Optional ByVal executionContextId AS Variant, _
    Optional ByVal objectGroup AS Variant, _
    Optional ByVal silent AS Variant, _
    Optional ByVal includeCommandLineAPI AS Variant, _
    Optional ByVal returnByValue AS Variant, _
    Optional ByVal generatePreview AS Variant, _
    Optional ByVal awaitPromise AS Variant _
) AS Dictionary
    ' scriptId: ScriptId Id of the script to run.
    ' executionContextId: ExecutionContextId(optional) Specifies in which execution context to perform script run. If the parameter is omitted the
    '   evaluation will be performed in the context of the inspected page.
    ' objectGroup: string(optional) Symbolic group name that can be used to release multiple objects.
    ' silent: boolean(optional) In silent mode exceptions thrown during evaluation are not reported and do not pause
    '   execution. Overrides `setPauseOnException` state.
    ' includeCommandLineAPI: boolean(optional) Determines whether Command Line API should be available during the evaluation.
    ' returnByValue: boolean(optional) Whether the result is expected to be a JSON object which should be sent by value.
    ' generatePreview: boolean(optional) Whether preview should be generated for the result.
    ' awaitPromise: boolean(optional) Whether execution should `await` for resulting value and return once awaited promise is
    '   resolved.

    Dim params As New Dictionary
    params("scriptId") = CStr(scriptId)
    If Not IsMissing(executionContextId) Then params("executionContextId") = CLng(executionContextId)
    If Not IsMissing(objectGroup) Then params("objectGroup") = CStr(objectGroup)
    If Not IsMissing(silent) Then params("silent") = CBool(silent)
    If Not IsMissing(includeCommandLineAPI) Then params("includeCommandLineAPI") = CBool(includeCommandLineAPI)
    If Not IsMissing(returnByValue) Then params("returnByValue") = CBool(returnByValue)
    If Not IsMissing(generatePreview) Then params("generatePreview") = CBool(generatePreview)
    If Not IsMissing(awaitPromise) Then params("awaitPromise") = CBool(awaitPromise)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Runtime.runScript", params)
    If cdp.ErrorCode = 0 Then
        Set runScript = results
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

    cdp.InvokeMethod "Runtime.setAsyncCallStackDepth", params
End Sub

' 
' Experimental
Public Sub setCustomObjectFormatterEnabled( _
    ByVal enabled AS boolean _
)
    ' enabled: boolean 

    Dim params As New Dictionary
    params("enabled") = CBool(enabled)

    cdp.InvokeMethod "Runtime.setCustomObjectFormatterEnabled", params
End Sub

' 
' Experimental
Public Sub setMaxCallStackSizeToCapture( _
    ByVal size AS Long _
)
    ' size: integer 

    Dim params As New Dictionary
    params("size") = CLng(size)

    cdp.InvokeMethod "Runtime.setMaxCallStackSizeToCapture", params
End Sub

' Terminate current or next JavaScript execution.
' Will cancel the termination when the outer-most script execution ends.
' Experimental
Public Sub terminateExecution()
    Dim params As New Dictionary
    cdp.InvokeMethod "Runtime.terminateExecution", params
End Sub

' If executionContextId is empty, adds binding with the given name on the
' global objects of all inspected contexts, including those created later,
' bindings survive reloads.
' Binding function takes exactly one argument, this argument should be string,
' in case of any other input, function throws an exception.
' Each binding function call produces Runtime.bindingCalled notification.
' Experimental
Public Sub addBinding( _
    ByVal name AS string, _
    Optional ByVal executionContextId AS Variant, _
    Optional ByVal executionContextName AS Variant _
)
    ' name: string 
    ' executionContextId: ExecutionContextId(optional) If specified, the binding would only be exposed to the specified
    '   execution context. If omitted and `executionContextName` is not set,
    '   the binding is exposed to all execution contexts of the target.
    '   This parameter is mutually exclusive with `executionContextName`.
    '   Deprecated in favor of `executionContextName` due to an unclear use case
    '   and bugs in implementation (crbug.com/1169639). `executionContextId` will be
    '   removed in the future.
    ' executionContextName: string(optional) If specified, the binding is exposed to the executionContext with
    '   matching name, even for contexts created after the binding is added.
    '   See also `ExecutionContext.name` and `worldName` parameter to
    '   `Page.addScriptToEvaluateOnNewDocument`.
    '   This parameter is mutually exclusive with `executionContextId`.

    Dim params As New Dictionary
    params("name") = CStr(name)
    If Not IsMissing(executionContextId) Then params("executionContextId") = CLng(executionContextId)
    If Not IsMissing(executionContextName) Then params("executionContextName") = CStr(executionContextName)

    cdp.InvokeMethod "Runtime.addBinding", params
End Sub

' This method does not remove binding function from global object but
' unsubscribes current runtime agent from Runtime.bindingCalled notifications.
' Experimental
Public Sub removeBinding( _
    ByVal name AS string _
)
    ' name: string 

    Dim params As New Dictionary
    params("name") = CStr(name)

    cdp.InvokeMethod "Runtime.removeBinding", params
End Sub

