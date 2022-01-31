VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageCompilationCacheParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "CompilationCacheParams"
' Page.CompilationCacheParams
' Per-script compilation cache parameters for `Page.produceCompilationCache`
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' url : string
' The URL of the script to produce a compilation cache entry for.
Public url AS string

' eager : boolean
' Optional
' A hint to the backend whether eager compilation is recommended.
    '   (the actual compilation mode used is upon backend discretion).
Public eager AS Variant ' boolean

Public Function init(ByVal obj as Dictionary) As cdpPageCompilationCacheParams
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let url = CStr(obj.Item("url"))
    If obj.Exists("eager") Then Let eager = CBool(obj.Item("eager"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("url") = url
    if Not IsEmpty(eager) Then dict("eager") = eager

    Set asDictionary = dict
End Function

