VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBrowserBucket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Browser.Bucket
' Chrome histogram bucket.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' low : integer
' Minimum value (inclusive).
Public low AS integer

' high : integer
' Maximum value (exclusive).
Public high AS integer

' count : integer
' Number of samples.
Public count AS integer

Public Function init(ByVal obj as Dictionary) As cdpBrowserBucket

    Let low = CInt(obj.Item("low"))
    Let high = CInt(obj.Item("high"))
    Let count = CInt(obj.Item("count"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("low") = low
    dict("high") = high
    dict("count") = count

    Set asDictionary = dict
End Function

