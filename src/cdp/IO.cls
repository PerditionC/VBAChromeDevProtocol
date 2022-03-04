VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cdpIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "cdpIO"
' IO
' This class is automatically generated, please make changes to generator and not this class directly.
Option Explicit

Private cdp As clsCDP



' *** Types:
' This is either obtained from another method or specified as `blob:&lt;uuid&gt;` where
'   `&lt;uuid&gt` is an UUID of a Blob.
' StreamHandle AS string



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

' Close the stream, discard any temporary backing storage.
Public Sub closeIO( _
    ByVal handle AS string _
)
    ' handle: StreamHandle Handle of the stream to close.

    Dim params As New Dictionary
    params("handle") = CStr(handle)

    cdp.InvokeMethod "IO.close", params
End Sub

' Read a chunk of the stream
' Returns:  - 
Public Function read( _
    ByVal handle AS string, _
    Optional ByVal offset AS Variant, _
    Optional ByVal size AS Variant _
) AS Dictionary
    ' handle: StreamHandle Handle of the stream to read.
    ' offset: integer(optional) Seek to the specified offset before reading (if not specificed, proceed with offset
    '   following the last read). Some types of streams may only support sequential reads.
    ' size: integer(optional) Maximum number of bytes to read (left upon the agent discretion if not specified).

    Dim params As New Dictionary
    params("handle") = CStr(handle)
    If Not IsMissing(offset) Then params("offset") = CLng(offset)
    If Not IsMissing(size) Then params("size") = CLng(size)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("IO.read", params)
    If cdp.ErrorCode = 0 Then
        Set read = results
    End If
End Function

' Return UUID of Blob object specified by a remote object id.
' Returns: uuid - UUID of the specified Blob.
Public Function resolveBlob( _
    ByVal objectId AS string _
) AS string
    ' objectId: Runtime.RemoteObjectId Object id of a Blob object wrapper.

    Dim params As New Dictionary
    params("objectId") = CStr(objectId)

    Dim results as Dictionary
    Set results = cdp.InvokeMethod("IO.resolveBlob", params)
    If cdp.ErrorCode = 0 Then
        If results.Exists("uuid") Then resolveBlob = results("uuid")
    End If
End Function

