VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpWebAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpWebAudio"
' WebAudio [Experimental]
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' An unique ID for a graph object (AudioContext, AudioNode, AudioParam) in Web Audio API
' GraphObjectId AS string

' Enum of BaseAudioContext types
' ContextType AS string
Public Enum ContextType
    CT_realtime
    CT_offline
End Enum

' Enum of AudioContextState from the spec
' ContextState AS string
Public Enum ContextState
    CS_suspended
    CS_running
    CS_closed
End Enum

' Enum of AudioNode types
' NodeType AS string

' Enum of AudioNode::ChannelCountMode from the spec
' ChannelCountMode AS string
Public Enum ChannelCountMode
    CCM_clampedmax
    CCM_explicit
    CCM_max
End Enum

' Enum of AudioNode::ChannelInterpretation from the spec
' ChannelInterpretation AS string
Public Enum ChannelInterpretation
    CI_discrete
    CI_speakers
End Enum

' Enum of AudioParam types
' ParamType AS string

' Enum of AudioParam::AutomationRate from the spec
' AutomationRate AS string
Public Enum AutomationRate
    AR_arate
    AR_krate
End Enum

' Fields in AudioContext that change in real-time.
' ContextRealtimeData AS object

' Protocol object for BaseAudioContext
' BaseAudioContext AS object

' Protocol object for AudioListener
' AudioListener AS object

' Protocol object for AudioNode
' AudioNode AS object

' Protocol object for AudioParam
' AudioParam AS object


Public Function ContextTypeToString(ByVal ct As ContextType) As String
    Dim retVal As String
    Select Case ct
        Case CT_realtime
            retVal = "realtime"
        Case CT_offline
            retVal = "offline"
        Case Else
            Debug.Print "Warning, unknown value " & ct
    End Select
    ContextTypeToString = retVal
End Function

Public Function StringToContextType(ByVal s As String) As ContextType
    Dim retVal As ContextType
    Select Case s
        Case "realtime"
            retVal = CT_realtime
        Case "offline"
            retVal = CT_offline
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToContextType = retVal
End Function


Public Function ContextStateToString(ByVal cs As ContextState) As String
    Dim retVal As String
    Select Case cs
        Case CS_suspended
            retVal = "suspended"
        Case CS_running
            retVal = "running"
        Case CS_closed
            retVal = "closed"
        Case Else
            Debug.Print "Warning, unknown value " & cs
    End Select
    ContextStateToString = retVal
End Function

Public Function StringToContextState(ByVal s As String) As ContextState
    Dim retVal As ContextState
    Select Case s
        Case "suspended"
            retVal = CS_suspended
        Case "running"
            retVal = CS_running
        Case "closed"
            retVal = CS_closed
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToContextState = retVal
End Function


Public Function ChannelCountModeToString(ByVal ccm As ChannelCountMode) As String
    Dim retVal As String
    Select Case ccm
        Case CCM_clampedmax
            retVal = "clamped-max"
        Case CCM_explicit
            retVal = "explicit"
        Case CCM_max
            retVal = "max"
        Case Else
            Debug.Print "Warning, unknown value " & ccm
    End Select
    ChannelCountModeToString = retVal
End Function

Public Function StringToChannelCountMode(ByVal s As String) As ChannelCountMode
    Dim retVal As ChannelCountMode
    Select Case s
        Case "clamped-max"
            retVal = CCM_clampedmax
        Case "explicit"
            retVal = CCM_explicit
        Case "max"
            retVal = CCM_max
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToChannelCountMode = retVal
End Function


Public Function ChannelInterpretationToString(ByVal ci As ChannelInterpretation) As String
    Dim retVal As String
    Select Case ci
        Case CI_discrete
            retVal = "discrete"
        Case CI_speakers
            retVal = "speakers"
        Case Else
            Debug.Print "Warning, unknown value " & ci
    End Select
    ChannelInterpretationToString = retVal
End Function

Public Function StringToChannelInterpretation(ByVal s As String) As ChannelInterpretation
    Dim retVal As ChannelInterpretation
    Select Case s
        Case "discrete"
            retVal = CI_discrete
        Case "speakers"
            retVal = CI_speakers
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToChannelInterpretation = retVal
End Function


Public Function AutomationRateToString(ByVal ar As AutomationRate) As String
    Dim retVal As String
    Select Case ar
        Case AR_arate
            retVal = "a-rate"
        Case AR_krate
            retVal = "k-rate"
        Case Else
            Debug.Print "Warning, unknown value " & ar
    End Select
    AutomationRateToString = retVal
End Function

Public Function StringToAutomationRate(ByVal s As String) As AutomationRate
    Dim retVal As AutomationRate
    Select Case s
        Case "a-rate"
            retVal = AR_arate
        Case "k-rate"
            retVal = AR_krate
        Case Else
            Debug.Print "Warning, unknown value " & s
    End Select
    StringToAutomationRate = retVal
End Function


Public Sub init(ByRef cdpObj As clsCDP)
    Set cdp = cdpObj
End Sub

'Private Sub Class_Initialize()
    ' add any needed initialization logic here
'End Sub

Private Sub Class_Terminate()
    ' ensure we don't have cyclic dependencies; clsCDP references this, but we also reference clsCDP instance
    Set cdp = Nothing
End Sub


' *** Commands:

' Enables the WebAudio domain and starts sending context lifetime events.
Public Sub enable()
    Dim params As New Dictionary
    cdp.InvokeMethod "WebAudio.enable", params
End Sub

' Disables the WebAudio domain.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "WebAudio.disable", params
End Sub

' Fetch the realtime data from the registered contexts.
' Returns: realtimeData - 
Public Function getRealtimeData( _
    ByVal contextId AS string _
) AS Dictionary
    ' contextId: GraphObjectId 

    Dim params As New Dictionary
    params("contextId") = CStr(contextId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("WebAudio.getRealtimeData", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("realtimeData") Then Set getRealtimeData = results("realtimeData")
    End If
End Function

