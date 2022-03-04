VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageAppManifestError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "AppManifestError"
' Page.AppManifestError
' Error while paring app manifest.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' message : string
' Error message.
Public message AS string

' critical : integer
' If criticial, this is a non-recoverable parse error.
Public critical AS Long

' line : integer
' Error line.
Public line AS Long

' column : integer
' Error column.
Public column AS Long

Public Function init(ByVal obj as Dictionary) As cdpPageAppManifestError
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let message = CStr(obj.Item("message"))
    Let critical = CLng(obj.Item("critical"))
    Let line = CLng(obj.Item("line"))
    Let column = CLng(obj.Item("column"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("message") = message
    dict("critical") = critical
    dict("line") = line
    dict("column") = column

    Set asDictionary = dict
End Function

