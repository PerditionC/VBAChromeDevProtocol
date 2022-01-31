VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeStackTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "StackTrace"
' Runtime.StackTrace
' Call frames for assertions or error messages.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' description : string
' Optional
' String label of this stack trace. For async traces this may be a name of the function that
    '   initiated the async call.
Public description AS string

' callFrames : array of cdpRuntimeCallFrame
' JavaScript function name.
Public callFrames AS Collection

' parent : StackTrace
' Optional
' Asynchronous JavaScript stack trace that preceded this stack, if available.
Public parent AS cdpRuntimeStackTrace

' parentId : StackTraceId
' Optional
' Asynchronous JavaScript stack trace that preceded this stack, if available.
Public parentId AS cdpRuntimeStackTraceId

Public Function init(ByVal obj as Dictionary) As cdpRuntimeStackTrace
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("description") Then Let description = CStr(obj.Item("description"))
    For each v in obj.Item("callFrames")
        Dim obj_callFrames As cdpRuntimeCallFrame: Set obj_callFrames = New cdpRuntimeCallFrame
        obj_callFrames.init v
        callFrames.Add obj_callFrames
        Set obj_callFrames = Nothing
    Next v
    If obj.Exists("parent") Then
        Set parent = New cdpRuntimeStackTrace
        parent.init obj.Item("parent")
    End If
    If obj.Exists("parentId") Then
        Set parentId = New cdpRuntimeStackTraceId
        parentId.init obj.Item("parentId")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if description <> vbNullString Then dict("description") = description
    Set dict("callFrames") = callFrames
    If Not parent Is Nothing Then Set dict("parent") = parent.asDictionary()
    If Not parentId Is Nothing Then Set dict("parentId") = parentId.asDictionary()

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set callFrames = New Collection
End Sub

