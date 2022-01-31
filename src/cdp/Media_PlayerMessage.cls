VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpMediaPlayerMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "PlayerMessage"
' Media.PlayerMessage
' Have one type per entry in MediaLogRecord::Type
'   Corresponds to kMessage
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' level : string
' Keep in sync with MediaLogMessageLevel
    '   We are currently keeping the message level 'error' separate from the
    '   PlayerError type because right now they represent different things,
    '   this one being a DVLOG(ERROR) style log message that gets printed
    '   based on what log level is selected in the UI, and the other is a
    '   representation of a media::PipelineStatus object. Soon however we're
    '   going to be moving away from using PipelineStatus for errors and
    '   introducing a new error type which should hopefully let us integrate
    '   the error log level into the PlayerError type.
Public level AS string

' message : string
' 
Public message AS string

Public Function init(ByVal obj as Dictionary) As cdpMediaPlayerMessage
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let level = CStr(obj.Item("level"))
    Let message = CStr(obj.Item("message"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("level") = level
    dict("message") = message

    Set asDictionary = dict
End Function

