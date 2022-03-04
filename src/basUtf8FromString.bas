Attribute VB_Name = "basUtf8FromString"
Option Explicit

' code from: https://www.di-mgt.com.au/basUtf8FromString64.bas.html

' basUtf8FromString

' Written by David Ireland DI Management Services Pty Limited 2015
' <http://www.di-mgt.com.au> <http://www.cryptosys.net>
' @license MIT license <http://opensource.org/licenses/MIT>
' [2015-06-30] First published.
' [2018-07-27] Updated to cope with empty or null input string.
' [2018-08-15] Added Utf8BytesToString and BytesLength functions.
' [2018-08-17] Added declaration for 64-bit Office.

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
#If VBA7 Then
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
#Else
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
#End If
    
''' Maps a character string to a UTF-16 (wide character) string
#If VBA7 Then
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
    ) As Long
#Else
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
    ) As Long
#End If

' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return length of byte array or zero if uninitialized
Public Function BytesLength(abBytes() As Byte) As Long
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function


''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function


''' Return VBA "Unicode" string from byte array encoded in UTF-8
Public Function Utf8BytesToString(abUtf8Array() As Byte, Optional ByVal nBytes As Long = 0) As String
    Dim maxBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    maxBytes = BytesLength(abUtf8Array)
    If nBytes = 0 Then nBytes = maxBytes ' use calculated value if not specified
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = left$(strOut, nChars)
End Function


' QUICK TEST
Public Sub Test_Utf8String()
    Dim b() As Byte
    Dim s As String
    
    b = Utf8BytesFromString("áéíóñ")
    s = Utf8BytesToString(b)
    Debug.Print "[" & s & "]"
End Sub

