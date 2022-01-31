VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpEmulationUserAgentMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "UserAgentMetadata"
' Emulation.UserAgentMetadata
' Used to specify User Agent Cient Hints to emulate. See https://wicg.github.io/ua-client-hints
'   Missing optional values will be filled in by the target with what it would normally use.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' brands : array of cdpEmulationUserAgentBrandVersi
' Optional
' 
Public brands AS Collection

' fullVersion : string
' Optional
' 
Public fullVersion AS string

' platform : string
' 
Public platform AS string

' platformVersion : string
' 
Public platformVersion AS string

' architecture : string
' 
Public architecture AS string

' model : string
' 
Public model AS string

' mobile : boolean
' 
Public mobile AS boolean

Public Function init(ByVal obj as Dictionary) As cdpEmulationUserAgentMetadata
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    If obj.Exists("brands") Then
        For each v in obj.Item("brands")
            Dim obj_brands As cdpEmulationUserAgentBrandVersi: Set obj_brands = New cdpEmulationUserAgentBrandVersi
            obj_brands.init v
            brands.Add obj_brands
            Set obj_brands = Nothing
        Next v
    End If
    If obj.Exists("fullVersion") Then Let fullVersion = CStr(obj.Item("fullVersion"))
    Let platform = CStr(obj.Item("platform"))
    Let platformVersion = CStr(obj.Item("platformVersion"))
    Let architecture = CStr(obj.Item("architecture"))
    Let model = CStr(obj.Item("model"))
    Let mobile = CBool(obj.Item("mobile"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("brands") = brands
    if fullVersion <> vbNullString Then dict("fullVersion") = fullVersion
    dict("platform") = platform
    dict("platformVersion") = platformVersion
    dict("architecture") = architecture
    dict("model") = model
    dict("mobile") = mobile

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set brands = New Collection
End Sub

