VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpSystemInfoProcessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ProcessInfo"
' SystemInfo.ProcessInfo
' Represents process info.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' type : string
' Specifies process type.
Public fieldType AS string

' id : integer
' Specifies process id.
Public id AS Long

' cpuTime : number
' Specifies cumulative CPU usage in seconds across all threads of the
    '   process since the process start.
Public cpuTime AS Double

Public Function init(ByVal obj as Dictionary) As cdpSystemInfoProcessInfo
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let fieldType = CStr(obj.Item("type"))
    Let id = CLng(obj.Item("id"))
    Let cpuTime = CDbl(obj.Item("cpuTime"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("type") = fieldType
    dict("id") = id
    dict("cpuTime") = cpuTime

    Set asDictionary = dict
End Function

