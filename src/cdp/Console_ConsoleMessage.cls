VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpConsoleConsoleMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ConsoleMessage"
' Console.ConsoleMessage
' Console message.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' source : string
' Message source.
Public source AS string

' level : string
' Message severity.
Public level AS string

' text : string
' Message text.
Public text AS string

' url : string
' Optional
' URL of the message origin.
Public url AS string

' line : integer
' Optional
' Line number in the resource that generated this message (1-based).
Public line AS Long

' column : integer
' Optional
' Column number in the resource that generated this message (1-based).
Public column AS Long

Public Function init(ByVal obj as Dictionary) As cdpConsoleConsoleMessage
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let source = CStr(obj.Item("source"))
    Let level = CStr(obj.Item("level"))
    Let text = CStr(obj.Item("text"))
    If obj.Exists("url") Then Let url = CStr(obj.Item("url"))
    If obj.Exists("line") Then Let line = CLng(obj.Item("line"))
    If obj.Exists("column") Then Let column = CLng(obj.Item("column"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("source") = source
    dict("level") = level
    dict("text") = text
    if url <> vbNullString Then dict("url") = url
    dict("line") = line
    dict("column") = column

    Set asDictionary = dict
End Function

