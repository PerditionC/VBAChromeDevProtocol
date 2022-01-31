VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpStorageTrustTokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "TrustTokens"
' Storage.TrustTokens
' Pair of issuer origin and number of available (signed, but not used) Trust
'   Tokens from that issuer.
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' issuerOrigin : string
' 
Public issuerOrigin AS string

' count : number
' 
Public count AS Double

Public Function init(ByVal obj as Dictionary) As cdpStorageTrustTokens
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let issuerOrigin = CStr(obj.Item("issuerOrigin"))
    Let count = CDbl(obj.Item("count"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("issuerOrigin") = issuerOrigin
    dict("count") = count

    Set asDictionary = dict
End Function

