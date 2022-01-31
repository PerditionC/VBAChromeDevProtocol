VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageInstallabilityError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "InstallabilityError"
' Page.InstallabilityError
' The installability error
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' errorId : string
' The error id (e.g. 'manifest-missing-suitable-icon').
Public errorId AS string

' errorArguments : array of cdpPageInstallabilityErrorArgum
' The list of error arguments (e.g. {name:'minimum-icon-size-in-pixels', value:'64'}).
Public errorArguments AS Collection

Public Function init(ByVal obj as Dictionary) As cdpPageInstallabilityError
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."
    Dim v as Variant

    Let errorId = CStr(obj.Item("errorId"))
    For each v in obj.Item("errorArguments")
        Dim obj_errorArguments As cdpPageInstallabilityErrorArgum: Set obj_errorArguments = New cdpPageInstallabilityErrorArgum
        obj_errorArguments.init v
        errorArguments.Add obj_errorArguments
        Set obj_errorArguments = Nothing
    Next v

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("errorId") = errorId
    Set dict("errorArguments") = errorArguments

    Set asDictionary = dict
End Function

Private Sub Class_Initialize()
    Set errorArguments = New Collection
End Sub

