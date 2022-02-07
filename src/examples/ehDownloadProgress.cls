VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ehDownloadProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Example event handler for downlod events to display progress bar during downloads
' Note: we only handle 1 download at a time, for multiple downloads the guid should be used to match up begin with progress events
Option Explicit

Public Status As String

'Private uf As ufProgress
Private filename As String

' returns True if event handled (don't add to message queue)
Public Function processEvent(ByVal eventName As String, ByVal eventData As Dictionary) As Boolean
    ' we don't need sessionId so discard everything but event parameters
    Set eventData = eventData.Item("params")
    
    If eventName = "Browser.downloadWillBegin" Or eventName = "Page.downloadWillBegin" Then
        Status = "start"
        'If uf Is Nothing Then Set uf = New ufProgress
        ufProgress.Caption = "Download Progress - Download beginning!"
        ufProgress.Show False
        filename = eventData.Item("suggestedFilename")
        processEvent = False ' say we didn't handle it (normally we would say True, but this won't hurt anything)
    ElseIf eventName = "Browser.downloadProgress" Or eventName = "Page.downloadProgress" Then
        Dim state As String
        state = eventData.Item("state")
        Status = state
        Dim Percent As Single
        Dim receivedBytes As Double, totalBytes As Double
        receivedBytes = eventData.Item("receivedBytes")
        totalBytes = eventData.Item("totalBytes")
        If state = "inProgress" Then
            If totalBytes > 0 Then
                Percent = receivedBytes / totalBytes * 100#
            Else
                Percent = 0#
            End If
        Else
            Percent = 100#
        End If
        
        'If Not uf Is Nothing Then
            If Percent = 100# Then
                ufProgress.Caption = "Download Progress - File successfully downloaded"
            Else
                ufProgress.Caption = "Download Progress - " & filename
            End If
            ufProgress.Progress = Percent
        
            If state = "completed" Then
                sleep (1) ' give user a chance to see dialog with completed message before we close it
                ufProgress.hide
                'Unload uf
                'Set uf = Nothing
            End If
        'End If
        
        processEvent = True
    Else
        Debug.Print "Received Event: " & eventName
        'processEvent = False
    End If
End Function

#If False Then
Private Sub Class_Terminate()
    If Not uf Is Nothing Then
        uf.hide
        Unload uf
        Set uf = Nothing
    End If
End Sub
#End If
