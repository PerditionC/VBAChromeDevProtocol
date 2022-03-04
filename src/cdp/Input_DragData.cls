VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpInputDragData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DragData"
' Input.DragData
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' items : array of cdpInputDragDataItem
' 
Public items AS Collection

' files : array of string
' Optional
' List of filenames that should be included when dropping
Public files AS Collection

' dragOperationsMask : integer
' Bit field representing allowed drag operations. Copy = 1, Link = 2, Move = 16
Public dragOperationsMask AS Long

Public Function init(ByVal obj as Dictionary) As cdpInputDragData
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    For each v in obj.Item("items")
        Dim obj_items As cdpInputDragDataItem: Set obj_items = New cdpInputDragDataItem
        obj_items.init v
        items.Add obj_items
        Set obj_items = Nothing
    Next v
    If obj.Exists("files") Then
        For each v in obj.Item("files")
            files.Add v
        Next v
    End If
    Let dragOperationsMask = CLng(obj.Item("dragOperationsMask"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    Set dict("items") = items
    Set dict("files") = files
    dict("dragOperationsMask") = dragOperationsMask

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set items = New Collection
    Set files = New Collection
End Sub

