VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpRuntimeObjectPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ObjectPreview"
' Runtime.ObjectPreview
' Object containing abbreviated remote object value.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' Object type.
Public fieldType AS string

' subtype : string
' Optional
' Object subtype hint. Specified for `object` type values only.
Public subtype AS string

' description : string
' Optional
' String representation of the object.
Public description AS string

' overflow : boolean
' True iff some of the properties or entries of the original object did not fit.
Public overflow AS boolean

' properties : array of cdpRuntimePropertyPreview
' List of the properties.
Public properties AS Collection

' entries : array of cdpRuntimeEntryPreview
' Optional
' List of the entries. Specified for `map` and `set` subtype values only.
Public entries AS Collection

Public Function init(ByVal obj as Dictionary) As cdpRuntimeObjectPreview
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let fieldType = CStr(obj.Item("type"))
    If obj.Exists("subtype") Then Let subtype = CStr(obj.Item("subtype"))
    If obj.Exists("description") Then Let description = CStr(obj.Item("description"))
    Let overflow = CBool(obj.Item("overflow"))
    For each v in obj.Item("properties")
        Dim obj_properties As cdpRuntimePropertyPreview: Set obj_properties = New cdpRuntimePropertyPreview
        obj_properties.init v
        properties.Add obj_properties
        Set obj_properties = Nothing
    Next v
    If obj.Exists("entries") Then
        For each v in obj.Item("entries")
            Dim obj_entries As cdpRuntimeEntryPreview: Set obj_entries = New cdpRuntimeEntryPreview
            obj_entries.init v
            entries.Add obj_entries
            Set obj_entries = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    if subtype <> vbNullString Then dict("subtype") = subtype
    if description <> vbNullString Then dict("description") = description
    dict("overflow") = overflow
    Set dict("properties") = properties
    Set dict("entries") = entries

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set properties = New Collection
    Set entries = New Collection
End Sub

