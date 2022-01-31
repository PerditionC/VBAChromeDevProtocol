VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageFontFamilies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "FontFamilies"
' Page.FontFamilies
' Generic font families collection.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' standard : string
' Optional
' The standard font-family.
Public standard AS string

' fixed : string
' Optional
' The fixed font-family.
Public fixed AS string

' serif : string
' Optional
' The serif font-family.
Public serif AS string

' sansSerif : string
' Optional
' The sansSerif font-family.
Public sansSerif AS string

' cursive : string
' Optional
' The cursive font-family.
Public cursive AS string

' fantasy : string
' Optional
' The fantasy font-family.
Public fantasy AS string

' pictograph : string
' Optional
' The pictograph font-family.
Public pictograph AS string

Public Function init(ByVal obj as Dictionary) As cdpPageFontFamilies
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    If obj.Exists("standard") Then Let standard = CStr(obj.Item("standard"))
    If obj.Exists("fixed") Then Let fixed = CStr(obj.Item("fixed"))
    If obj.Exists("serif") Then Let serif = CStr(obj.Item("serif"))
    If obj.Exists("sansSerif") Then Let sansSerif = CStr(obj.Item("sansSerif"))
    If obj.Exists("cursive") Then Let cursive = CStr(obj.Item("cursive"))
    If obj.Exists("fantasy") Then Let fantasy = CStr(obj.Item("fantasy"))
    If obj.Exists("pictograph") Then Let pictograph = CStr(obj.Item("pictograph"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    if standard <> vbNullString Then dict("standard") = standard
    if fixed <> vbNullString Then dict("fixed") = fixed
    if serif <> vbNullString Then dict("serif") = serif
    if sansSerif <> vbNullString Then dict("sansSerif") = sansSerif
    if cursive <> vbNullString Then dict("cursive") = cursive
    if fantasy <> vbNullString Then dict("fantasy") = fantasy
    if pictograph <> vbNullString Then dict("pictograph") = pictograph

    Set asDictionary = dict
End Function

