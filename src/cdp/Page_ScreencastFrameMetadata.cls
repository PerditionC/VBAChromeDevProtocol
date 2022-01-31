VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageScreencastFrameMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ScreencastFrameMetadata"
' Page.ScreencastFrameMetadata
' Screencast frame metadata.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' offsetTop : number
' Top offset in DIP.
Public offsetTop AS Double

' pageScaleFactor : number
' Page scale factor.
Public pageScaleFactor AS Double

' deviceWidth : number
' Device screen width in DIP.
Public deviceWidth AS Double

' deviceHeight : number
' Device screen height in DIP.
Public deviceHeight AS Double

' scrollOffsetX : number
' Position of horizontal scroll in CSS pixels.
Public scrollOffsetX AS Double

' scrollOffsetY : number
' Position of vertical scroll in CSS pixels.
Public scrollOffsetY AS Double

' timestamp : Network.TimeSinceEpoch
' Optional
' Frame swap timestamp.
Public timestamp AS Double

Public Function init(ByVal obj as Dictionary) As cdpPageScreencastFrameMetadata
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let offsetTop = CDbl(obj.Item("offsetTop"))
    Let pageScaleFactor = CDbl(obj.Item("pageScaleFactor"))
    Let deviceWidth = CDbl(obj.Item("deviceWidth"))
    Let deviceHeight = CDbl(obj.Item("deviceHeight"))
    Let scrollOffsetX = CDbl(obj.Item("scrollOffsetX"))
    Let scrollOffsetY = CDbl(obj.Item("scrollOffsetY"))
    If obj.Exists("timestamp") Then Let timestamp = CDbl(obj.Item("timestamp"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("offsetTop") = offsetTop
    dict("pageScaleFactor") = pageScaleFactor
    dict("deviceWidth") = deviceWidth
    dict("deviceHeight") = deviceHeight
    dict("scrollOffsetX") = scrollOffsetX
    dict("scrollOffsetY") = scrollOffsetY
    dict("timestamp") = timestamp

    Set asDictionary = dict
End Function

