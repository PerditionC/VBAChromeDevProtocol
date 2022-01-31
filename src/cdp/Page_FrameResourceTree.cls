VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageFrameResourceTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FrameResourceTree"
' Page.FrameResourceTree
' Information about the Frame hierarchy along with their cached resources.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' frame : Frame
' Frame information for this tree item.
Public frame AS cdpPageFrame

' childFrames : array of cdpPageFrameResourceTree
' Optional
' Child frames.
Public childFrames AS Collection

' resources : array of cdpPageFrameResource
' Information about frame resources.
Public resources AS Collection

Public Function init(ByVal obj as Dictionary) As cdpPageFrameResourceTree
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Set frame = New cdpPageFrame
    frame.init obj.Item("frame")
    If obj.Exists("childFrames") Then
        For each v in obj.Item("childFrames")
            Dim obj_childFrames As cdpPageFrameResourceTree: Set obj_childFrames = New cdpPageFrameResourceTree
            obj_childFrames.init v
            childFrames.Add obj_childFrames
            Set obj_childFrames = Nothing
        Next v
    End If
    For each v in obj.Item("resources")
        Dim obj_resources As cdpPageFrameResource: Set obj_resources = New cdpPageFrameResource
        obj_resources.init v
        resources.Add obj_resources
        Set obj_resources = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("frame") = frame.asDictionary()
    Set dict("childFrames") = childFrames
    Set dict("resources") = resources

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set childFrames = New Collection
    Set resources = New Collection
End Sub

