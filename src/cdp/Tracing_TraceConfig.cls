VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpTracingTraceConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "TraceConfig"
' Tracing.TraceConfig
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' recordMode : string
' Optional
' Controls how the trace buffer stores data.
Public recordMode AS string

' enableSampling : boolean
' Optional
' Turns on JavaScript stack sampling.
Public enableSampling AS Variant ' boolean

' enableSystrace : boolean
' Optional
' Turns on system tracing.
Public enableSystrace AS Variant ' boolean

' enableArgumentFilter : boolean
' Optional
' Turns on argument filter.
Public enableArgumentFilter AS Variant ' boolean

' includedCategories : array of string
' Optional
' Included category filters.
Public includedCategories AS Collection

' excludedCategories : array of string
' Optional
' Excluded category filters.
Public excludedCategories AS Collection

' syntheticDelays : array of string
' Optional
' Configuration to synthesize the delays in tracing.
Public syntheticDelays AS Collection

' memoryDumpConfig : MemoryDumpConfig
' Optional
' Configuration for memory dump triggers. Used only when "memory-infra" category is enabled.
Public memoryDumpConfig AS cdpTracingMemoryDumpConfig

Public Function init(ByVal obj as Dictionary) As cdpTracingTraceConfig
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("recordMode") Then Let recordMode = CStr(obj.Item("recordMode"))
    If obj.Exists("enableSampling") Then Let enableSampling = CBool(obj.Item("enableSampling"))
    If obj.Exists("enableSystrace") Then Let enableSystrace = CBool(obj.Item("enableSystrace"))
    If obj.Exists("enableArgumentFilter") Then Let enableArgumentFilter = CBool(obj.Item("enableArgumentFilter"))
    If obj.Exists("includedCategories") Then
        For each v in obj.Item("includedCategories")
            includedCategories.Add v
        Next v
    End If
    If obj.Exists("excludedCategories") Then
        For each v in obj.Item("excludedCategories")
            excludedCategories.Add v
        Next v
    End If
    If obj.Exists("syntheticDelays") Then
        For each v in obj.Item("syntheticDelays")
            syntheticDelays.Add v
        Next v
    End If
    If obj.Exists("memoryDumpConfig") Then
        Set memoryDumpConfig = New cdpTracingMemoryDumpConfig
        memoryDumpConfig.init obj.Item("memoryDumpConfig")
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if recordMode <> vbNullString Then dict("recordMode") = recordMode
    if Not IsEmpty(enableSampling) Then dict("enableSampling") = enableSampling
    if Not IsEmpty(enableSystrace) Then dict("enableSystrace") = enableSystrace
    if Not IsEmpty(enableArgumentFilter) Then dict("enableArgumentFilter") = enableArgumentFilter
    Set dict("includedCategories") = includedCategories
    Set dict("excludedCategories") = excludedCategories
    Set dict("syntheticDelays") = syntheticDelays
    If Not memoryDumpConfig Is Nothing Then Set dict("memoryDumpConfig") = memoryDumpConfig.asDictionary()

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set includedCategories = New Collection
    Set excludedCategories = New Collection
    Set syntheticDelays = New Collection
End Sub

