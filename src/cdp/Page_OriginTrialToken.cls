VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPageOriginTrialToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "OriginTrialToken"
' Page.OriginTrialToken
' 
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

' origin : string
' 
Public origin AS string

' matchSubDomains : boolean
' 
Public matchSubDomains AS boolean

' trialName : string
' 
Public trialName AS string

' expiryTime : Network.TimeSinceEpoch
' 
Public expiryTime AS Double

' isThirdParty : boolean
' 
Public isThirdParty AS boolean

' usageRestriction : OriginTrialUsageRestriction
' 
Public usageRestriction AS string

Public Function init(ByVal obj as Dictionary) As cdpPageOriginTrialToken
Attribute Item.VB_Description = "Initialize class from Dictionary returned by CDP method."

    Let origin = CStr(obj.Item("origin"))
    Let matchSubDomains = CBool(obj.Item("matchSubDomains"))
    Let trialName = CStr(obj.Item("trialName"))
    Let expiryTime = CDbl(obj.Item("expiryTime"))
    Let isThirdParty = CBool(obj.Item("isThirdParty"))
    Let usageRestriction = CStr(obj.Item("usageRestriction"))

    Set init = Me
End Function

Public Function asDictionary() As Dictionary
    Dim dict as Dictionary: Set dict = New Dictionary

    dict("origin") = origin
    dict("matchSubDomains") = matchSubDomains
    dict("trialName") = trialName
    dict("expiryTime") = expiryTime
    dict("isThirdParty") = isThirdParty
    dict("usageRestriction") = usageRestriction

    Set asDictionary = dict
End Function

