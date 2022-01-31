VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMemorySamplingProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "SamplingProfile"
' Memory.SamplingProfile
' Array of heap profile samples.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' samples : array of cdpMemorySamplingProfileNode
' 
Public samples AS Collection

' modules : array of cdpMemoryModule
' 
Public modules AS Collection

Public Function init(ByVal obj as Dictionary) As cdpMemorySamplingProfile
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("samples")
        Dim obj_samples As cdpMemorySamplingProfileNode: Set obj_samples = New cdpMemorySamplingProfileNode
        obj_samples.init v
        samples.Add obj_samples
        Set obj_samples = Nothing
    Next v
    For each v in obj.Item("modules")
        Dim obj_modules As cdpMemoryModule: Set obj_modules = New cdpMemoryModule
        obj_modules.init v
        modules.Add obj_modules
        Set obj_modules = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("samples") = samples
    Set dict("modules") = modules

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set samples = New Collection
    Set modules = New Collection
End Sub

