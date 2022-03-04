VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpBrowserBucket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Bucket"
' Browser.Bucket
' Chrome histogram bucket.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' low : integer
' Minimum value (inclusive).
Public low AS Long

' high : integer
' Maximum value (exclusive).
Public high AS Long

' count : integer
' Number of samples.
Public count AS Long

Public Function init(ByVal obj as Dictionary) As cdpBrowserBucket
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let low = CLng(obj.Item("low"))
    Let high = CLng(obj.Item("high"))
    Let count = CLng(obj.Item("count"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("low") = low
    dict("high") = high
    dict("count") = count

    Set asDictionary = dict
End Function

