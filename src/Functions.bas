Attribute VB_Name = "Functions"
Option Explicit


'from https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit


Public Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    phReadPipe As LongPtr, _
    phWritePipe As LongPtr, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Public Declare PtrSafe Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, _
    lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As Any, _
    lpProcessInformation As Any) As Long

Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

Public Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long


Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    lpOverlapped As Long) As Long

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type


Public Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

'this is the structure to pass more than 3 fds to a child process

'see https://github.com/libuv/libuv/blob/v1.x/src/win/process-stdio.c
Public Type STDIO_BUFFER
    number_of_fds As Long
    crt_flags(0 To 4) As Byte
    os_handle(0 To 4) As LongPtr
End Type

' the fields crt_flags and os_handle must lie contigously in memory
' i.e. should not be aligned to byte boundaries
' you cannot define a packed struct in VBA
' thats why we need to have a second struct

#If Win64 Then
Public Type STDIO_BUFFER2
    number_of_fds As Long
    raw_bytes(0 To 44) As Byte
End Type
#Else
Public Type STDIO_BUFFER2
    number_of_fds As Long
    raw_bytes(0 To 24) As Byte
End Type
#End If

Public Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Const STARTF_USESTDHANDLES = &H100&
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESHOWWINDOW As Long = &H1&

' we need to move memory
'Public Declare PtrSafe Function GetAddrOf Lib "kernel32" Alias "MulDiv" (ByVal nNumber As Any, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
   ' This is the dummy function used to get the addres of a VB variable.

Public Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As LongPtr)


Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

'Custom sleep function
'change sleep period if processing is not robust
Public Const cnlngSleepPeriod As Long = 1000

Public Sub Sleep(Optional dblFrac As Double = 1)
    DoEvents
    Call sleep2(cnlngSleepPeriod * dblFrac)
    DoEvents
End Sub


' returns [current (Now)] time  as seconds since 1970, ie UnixTime
Function UnixTime(Optional localDateTime As Variant) As Long
    If IsMissing(localDateTime) Then
        UnixTime = DateDiff("S", "1/1/1970", Now())
    Else
        UnixTime = DateDiff("S", "1/1/1970", CDate(localDateTime))
    End If
End Function


' returns True if no process found or successfully terminated process; false or error or user cancels
Function TerminateProcess(ByVal ProcessName As String, Optional ByVal PromptBefore As Boolean = False) As Boolean
    On Error Resume Next ' for now ignore errors
    Dim processFound As Boolean: processFound = False
    Dim processes As Object
    Dim process As Object

    Set processes = getObject("winmgmts:").ExecQuery("SELECT * FROM win32_process")
    
    For Each process In processes
        If process.name = ProcessName Then
            processFound = True
        
            If PromptBefore Then
                If MsgBox("Browser Automation requires that " & ProcessName & " be closed before initializing. " & vbCrLf & _
                          "If you'll use Edge Automation routinely, it is suggested you use Chrome as your default browser. " & _
                          "If you need help changing your default browser contact the macro administrator for assistance." & vbCrLf & vbCrLf & _
                          "Click [OK] to terminate all Edge processes." & vbCrLf & _
                          "Click [Cancel] to abort and ", vbOKCancel, "Edge Automation") = vbCancel Then
                    'Abort runtime so user can finish up with their current Edge session(s)
                    GoTo Cleanup ' returns False
                End If
            End If
            
            TerminateProcess = (process.Terminate = 0)
            
            Exit For
        End If
    Next

Cleanup:
    Set processes = Nothing
    Set process = Nothing
    TerminateProcess = TerminateProcess Or (Not processFound) ' successfully terminated process or no process found
End Function


