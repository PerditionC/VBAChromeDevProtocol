VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpPerformance"
' Performance
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' Run-time execution metric.
' Metric AS object



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

' Disable collecting and reporting metrics.
Public Sub disable()
    Dim params As New Dictionary
    cdp.InvokeMethod "Performance.disable", params
End Sub

' Enable collecting and reporting metrics.
Public Sub enable( _
    Optional ByVal timeDomain AS Variant _
)
    ' timeDomain: string(optional) Time domain to use for collecting and reporting duration metrics.

    Dim params As New Dictionary
    If Not IsMissing(timeDomain) Then params("timeDomain") = CStr(timeDomain)

    cdp.InvokeMethod "Performance.enable", params
End Sub

' Sets time domain to use for collecting and reporting duration metrics.
' Note that this must be called before enabling metrics collection. Calling
' this method while metrics collection is enabled returns an error.
' Experimental
Public Sub setTimeDomain( _
    ByVal timeDomain AS string _
)
    ' timeDomain: string Time domain

    Dim params As New Dictionary
    params("timeDomain") = CStr(timeDomain)

    cdp.InvokeMethod "Performance.setTimeDomain", params
End Sub

' Retrieve current values of run-time metrics.
' Returns: metrics - Current values for run-time metrics.
Public Function getMetrics() AS Collection
    Dim params As New Dictionary
    Dim results as Dictionary
    Set results = cdp.InvokeMethod("Performance.getMetrics", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("metrics") Then Set getMetrics = results("metrics")
    End If
End Function

