VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Progress Indicator"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4290
   OleObjectBlob   =   "ufProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Progress bar form
Option Explicit

' To use either create a uf = new form and uf.show(False) then set uf.Progress = value
' or use implicitly generate form via ufProgress.show(False), ufProgress.Progress = value

' indicates if should display time left or note, defaults to True - show estimated time left
Public ShowTimeLeft As Boolean

' current displayed percent
Private currentPercent As Single
' when form first shown, assumed event has began since this
Private StartTime As Single


' initialize values prior to being made visible and set defaults
Private Sub UserForm_Initialize()
    ShowTimeLeft = True
    currentPercent = 0#
    Me.Bar.width = 0 ' initially set around 20 so can see in editor, but we don't want the flash from that to 0 at startup
End Sub

' initialize progress bar to 0 progress, unknown time left
Private Sub UserForm_Activate()
    Me.Bar.width = 0 ' reset if reusing progress bar via .hide then .show
    Me.Text.Caption = "0% Completed"
    currentPercent = 0
    StartTime = Timer
End Sub

' update progress bar (where Percent is math equation to generate the %, then *100)
Public Property Let Progress(ByVal Percent As Single)
    If Percent <= 0# Then
        Percent = 0.0001
    ElseIf Percent > 100# Then
        Percent = 100#
    End If
    
    ' update caption with percent done and either estimated time left or specified string
    If ShowTimeLeft Then
        ' calculate time left
        Dim TimeLeft As Single
        TimeLeft = Round(((Timer - StartTime) / Percent * 100) - (Timer - StartTime), 2)
        
        Dim SecMinHourDay, TimeLeftStr As String
        If TimeLeft > 86400 Then
            SecMinHourDay = " day(s) left."
            TimeLeft = Round(TimeLeft / 86400, 1)
        ElseIf TimeLeft > 3600 Then
            SecMinHourDay = " hour(s) left."
            TimeLeft = Round(TimeLeft / 3600, 1)
        ElseIf TimeLeft > 60 Then
            SecMinHourDay = " minute(s) left."
            TimeLeft = Round(TimeLeft / 60, 1)
        Else
            SecMinHourDay = " second(s) left."
            TimeLeft = Round(TimeLeft, 1)
        End If
        TimeLeftStr = TimeLeft & " " & SecMinHourDay
        
        Me.Text.Caption = Round(Percent, 1) & "% Completed" & vbCrLf & TimeLeftStr
    Else
        Me.Text.Caption = Round(Percent, 1) & "% Completed" & vbCrLf ' & Left(StrInLieuOfTimeLeft, 29)
    End If
    
    ' actually update bar
    Me.Bar.width = Percent * 2
    currentPercent = Percent
    
    ' give form a chance to paint update and generally allow Windows to process messages
    DoEvents
End Property

Public Property Get Progress() As Single
    Progress = currentPercent
End Property

