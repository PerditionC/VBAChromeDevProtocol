VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMemoryModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Module"
' Memory.Module
' Executable module information
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' name : string
' Name of the module.
Public name AS string

' uuid : string
' UUID of the module.
Public uuid AS string

' baseAddress : string
' Base address where the module is loaded into memory. Encoded as a decimal
    '   or hexadecimal (0x prefixed) string.
Public baseAddress AS string

' size : number
' Size of the module in bytes.
Public size AS Double

Public Function init(ByVal obj as Dictionary) As cdpMemoryModule
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let name = CStr(obj.Item("name"))
    Let uuid = CStr(obj.Item("uuid"))
    Let baseAddress = CStr(obj.Item("baseAddress"))
    Let size = CDbl(obj.Item("size"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("name") = name
    dict("uuid") = uuid
    dict("baseAddress") = baseAddress
    dict("size") = size

    Set asDictionary = dict
End Function

