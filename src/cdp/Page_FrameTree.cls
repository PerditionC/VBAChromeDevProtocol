VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageFrameTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FrameTree"
' Page.FrameTree
' Information about the Frame hierarchy.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' frame : Frame
' Frame information for this tree item.
Public frame AS cdpPageFrame

' childFrames : array of cdpPageFrameTree
' Optional
' Child frames.
Public childFrames AS Collection

Public Function init(ByVal obj as Dictionary) As cdpPageFrameTree
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Set frame = New cdpPageFrame
    frame.init obj.Item("frame")
    If obj.Exists("childFrames") Then
        For each v in obj.Item("childFrames")
            Dim obj_childFrames As cdpPageFrameTree: Set obj_childFrames = New cdpPageFrameTree
            obj_childFrames.init v
            childFrames.Add obj_childFrames
            Set obj_childFrames = Nothing
        Next v
    End If

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("frame") = frame.asDictionary()
    Set dict("childFrames") = childFrames

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set childFrames = New Collection
End Sub

