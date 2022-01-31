VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpDatabaseError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Database.Error
' Database error.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' message : string
' Error message.
Public message AS string

' code : integer
' Error code.
Public code AS integer

Public Function init(ByVal obj as Dictionary) As cdpDatabaseError

    Let message = CStr(obj.Item("message"))
    Let code = CInt(obj.Item("code"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("message") = message
    dict("code") = code

    Set asDictionary = dict
End Function

