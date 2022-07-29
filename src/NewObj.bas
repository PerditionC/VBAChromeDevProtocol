Attribute VB_Name = "NewObj"
' Helpers to create class Objects from other Projects
Option Explicit


Public Function New_clsCDP() As clsCDP
    Set New_clsCDP = New clsCDP
End Function

Public Function New_AutomateBrowser() As AutomateBrowser
    Set New_AutomateBrowser = New AutomateBrowser
End Function

Public Function New_ufProgress() As ufProgress
    Set New_ufProgress = New ufProgress
End Function
