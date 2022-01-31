VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfoImageDecodeAcceler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ImageDecodeAcceleratorCapability"
' SystemInfo.ImageDecodeAcceleratorCapability
' Describes a supported image decoding profile with its associated minimum and
'   maximum resolutions and subsampling.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' imageType : ImageType
' Image coded, e.g. Jpeg.
Public imageType AS string

' maxDimensions : Size
' Maximum supported dimensions of the image in pixels.
Public maxDimensions AS cdpSystemInfoSize

' minDimensions : Size
' Minimum supported dimensions of the image in pixels.
Public minDimensions AS cdpSystemInfoSize

' subsamplings : array of String
' Optional array of supported subsampling formats, e.g. 4:2:0, if known.
Public subsamplings AS Collection

Public Function init(ByVal obj as Dictionary) As cdpSystemInfoImageDecodeAcceler
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let imageType = CStr(obj.Item("imageType"))
    Set maxDimensions = New cdpSystemInfoSize
    maxDimensions.init obj.Item("maxDimensions")
    Set minDimensions = New cdpSystemInfoSize
    minDimensions.init obj.Item("minDimensions")
    For each v in obj.Item("subsamplings")
        subsamplings.Add v
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("imageType") = imageType
    Set dict("maxDimensions") = maxDimensions.asDictionary()
    Set dict("minDimensions") = minDimensions.asDictionary()
    Set dict("subsamplings") = subsamplings

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set subsamplings = New Collection
End Sub

