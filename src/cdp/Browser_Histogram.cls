VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBrowserHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Browser.Histogram
' Chrome histogram.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Name.
Public name AS string

' sum : integer
' Sum of sample values.
Public sum AS integer

' count : integer
' Total number of samples.
Public count AS integer

' buckets : array of cdpBrowserBucket
' Buckets.
Public buckets AS Collection

Public Function init(ByVal obj as Dictionary) As cdpBrowserHistogram
    Dim v as Variant

    Let name = CStr(obj.Item("name"))
    Let sum = CInt(obj.Item("sum"))
    Let count = CInt(obj.Item("count"))
    For each v in obj.Item("buckets")
        Dim obj_buckets As cdpBrowserBucket: Set obj_buckets = New cdpBrowserBucket
        obj_buckets.init v
        buckets.Add obj_buckets
        Set obj_buckets = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("sum") = sum
    dict("count") = count
    Set dict("buckets") = buckets

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set buckets = New Collection
End Sub
