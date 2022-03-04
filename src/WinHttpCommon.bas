Attribute VB_Name = "WinHttpCommon"
' see https://github.com/EagleAglow/vba-websocket-class
' MIT licensed
Option Explicit

' note constant names from C header file had underscore character
' these have been replaced with lowercase "x" to avoid possible conflicts with object scripting

' flags for WinHttpOpen():
Public Const WINHTTPxFLAGxSYNC = &H0          ' session is synchronous - use this
Public Const WINHTTPxFLAGxASYNC = &H10000000  ' session is asynchronous - not implemented

' codes from GetLastError
' see: https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-
Public Const ERRORxSUCCESS = 0
Public Const ERRORxINVALIDxFUNCTION = 1
Public Const ERRORxINVALIDxHANDLE = 6
Public Const ERRORxNOTxENOUGHxMEMORY = 8

' flags for status check
Public Const WINHTTPxQUERYxSTATUSxCODE = 19  ' special: part of status line
Public Const WINHTTPxQUERYxFLAGxNUMBER = &H20000000   ' bit flag to get result as number

' WinHttpQueryHeaders constants for code readability
Public Const WINHTTPxHEADERxNAMExBYxINDEX = 0
Public Const WINHTTPxNOxOUTPUTxBUFFER = 0
Public Const WINHTTPxNOxHEADERxINDEX = 0
Public Const WINHTTPxNOxREQUESTxDATA = 0

' a few HTTP Response Status Codes
Public Const HTTPxSTATUSxCONTINUE = 100           ' OK to continue with request
Public Const HTTPxSTATUSxSWITCHxPROTOCOLS = 101   ' server has switched protocols in upgrade header
Public Const HTTPxSTATUSxOK = 200                 ' request completed

Public Const WINHTTPxWEBxSOCKETxBUFFERxTYPE = 0  ' text, not specified in original sample code

Public Const INTERNETxDEFAULTxPORT = 0           ' use the protocol-specific default
Public Const INTERNETxDEFAULTxHTTPxPORT = 80     ' use the HTTP default
Public Const INTERNETxDEFAULTxHTTPSxPORT = 443   ' use the HTTPS default

Public Const WINHTTPxACCESSxTYPExDEFAULTxPROXY = 0
Public Const WINHTTPxACCESSxTYPExNOxPROXY = 1
Public Const WINHTTPxACCESSxTYPExNAMEDxPROXY = 3
Public Const WINHTTPxACCESSxTYPExAUTOMATICxPROXY = 4

Public Const WINHTTPxOPTIONxUPGRADExTOxWEBxSOCKET = 114
Public Const WINHTTPxNOxADDITIONALxHEADERS = 0

' buffer types
Public Const WINHTTPxWEBxSOCKETxBINARYxMESSAGExBUFFERxTYPE = 0
Public Const WINHTTPxWEBxSOCKETxBINARYxFRAGMENTxBUFFERxTYPE = 1
Public Const WINHTTPxWEBxSOCKETxUTF8xMESSAGExBUFFERxTYPE = 2
Public Const WINHTTPxWEBxSOCKETxUTF8xFRAGMENTxBUFFERxTYPE = 3
Public Const WINHTTPxWEBxSOCKETxCLOSExBUFFERxTYPE = 4

'    WEBxSOCKETxUTF8xMESSAGExBUFFERxTYPE             = &H80000000,
'    WEBxSOCKETxUTF8xFRAGMENTxBUFFERxTYPE            = &H80000001,
'    WEBxSOCKETxBINARYxMESSAGExBUFFERxTYPE           = &H80000002,
'    WEBxSOCKETxBINARYxFRAGMENTxBUFFERxTYPE          = &H80000003,
'    WEBxSOCKETxCLOSExBUFFERxTYPE                    = &H80000004,
'    WEBxSOCKETxPINGxPONGxBUFFERxTYPE                = &H80000005,
'    WEBxSOCKETxUNSOLICITEDxPONGxBUFFERxTYPE         = &H80000006,





' status codes
Public Const WINHTTPxWEBxSOCKETxSUCCESSxCLOSExSTATUS = 1000
Public Const WINHTTPxWEBxSOCKETxENDPOINTxTERMINATEDxCLOSExSTATUS = 1001
Public Const WINHTTPxWEBxSOCKETxPROTOCOLxERRORxCLOSExSTATUS = 1002
Public Const WINHTTPxWEBxSOCKETxINVALIDxDATAxTYPExCLOSExSTATUS = 1003
Public Const WINHTTPxWEBxSOCKETxEMPTYxCLOSExSTATUS = 1005
Public Const WINHTTPxWEBxSOCKETxABORTEDxCLOSExSTATUS = 1006
Public Const WINHTTPxWEBxSOCKETxINVALIDxPAYLOADxCLOSExSTATUS = 1007
Public Const WINHTTPxWEBxSOCKETxPOLICYxVIOLATIONxCLOSExSTATUS = 1008
Public Const WINHTTPxWEBxSOCKETxMESSAGExTOOxBIGxCLOSExSTATUS = 1009
Public Const WINHTTPxWEBxSOCKETxUNSUPPORTEDxEXTENSIONSxCLOSExSTATUS = 1010
Public Const WINHTTPxWEBxSOCKETxSERVERxERRORxCLOSExSTATUS = 1011
Public Const WINHTTPxWEBxSOCKETxSECURExHANDSHAKExERRORxCLOSExSTATUS = 1015


'//
'// status manifests for WinHttp status callback
'//

Public Const WINHTTPxCALLBACKxSTATUSxRESOLVINGxNAME = &H1
Public Const WINHTTPxCALLBACKxSTATUSxNAMExRESOLVED = &H2
Public Const WINHTTPxCALLBACKxSTATUSxCONNECTINGxTOxSERVER = &H4
Public Const WINHTTPxCALLBACKxSTATUSxCONNECTEDxTOxSERVER = &H8
Public Const WINHTTPxCALLBACKxSTATUSxSENDINGxREQUEST = &H10
Public Const WINHTTPxCALLBACKxSTATUSxREQUESTxSENT = &H20
Public Const WINHTTPxCALLBACKxSTATUSxRECEIVINGxRESPONSE = &H40
Public Const WINHTTPxCALLBACKxSTATUSxRESPONSExRECEIVED = &H80
Public Const WINHTTPxCALLBACKxSTATUSxCLOSINGxCONNECTION = &H100
Public Const WINHTTPxCALLBACKxSTATUSxCONNECTIONxCLOSED = &H200
Public Const WINHTTPxCALLBACKxSTATUSxHANDLExCREATED = &H400
Public Const WINHTTPxCALLBACKxSTATUSxHANDLExCLOSING = &H800
Public Const WINHTTPxCALLBACKxSTATUSxDETECTINGxPROXY = &H1000
Public Const WINHTTPxCALLBACKxSTATUSxREDIRECT = &H4000
Public Const WINHTTPxCALLBACKxSTATUSxINTERMEDIATExRESPONSE = &H8000
Public Const WINHTTPxCALLBACKxSTATUSxSECURExFAILURE = &H10000
Public Const WINHTTPxCALLBACKxSTATUSxHEADERSxAVAILABLE = &H20000
Public Const WINHTTPxCALLBACKxSTATUSxDATAxAVAILABLE = &H40000
Public Const WINHTTPxCALLBACKxSTATUSxREADxCOMPLETE = &H80000
Public Const WINHTTPxCALLBACKxSTATUSxWRITExCOMPLETE = &H100000
Public Const WINHTTPxCALLBACKxSTATUSxREQUESTxERROR = &H200000
Public Const WINHTTPxCALLBACKxSTATUSxSENDREQUESTxCOMPLETE = &H400000


Public Const WINHTTPxCALLBACKxSTATUSxGETPROXYFORURLxCOMPLETE = &H1000000
Public Const WINHTTPxCALLBACKxSTATUSxCLOSExCOMPLETE = &H2000000
Public Const WINHTTPxCALLBACKxSTATUSxSHUTDOWNxCOMPLETE = &H4000000
Public Const WINHTTPxCALLBACKxSTATUSxSETTINGSxWRITExCOMPLETE = &H10000000
Public Const WINHTTPxCALLBACKxSTATUSxSETTINGSxREADxCOMPLETE = &H20000000

'// API Enums for WINHTTPxCALLBACKxSTATUSxREQUESTxERROR:
Public Const APIxRECEIVExRESPONSE = 1
Public Const APIxQUERYxDATAxAVAILABLE = 2
Public Const APIxREADxDATA = 3
Public Const APIxWRITExDATA = 4
Public Const APIxSENDxREQUEST = 5
Public Const APIxGETxPROXYxFORxURL = 6


Public Const WINHTTPxCALLBACKxFLAGxRESOLVExNAME = (WINHTTPxCALLBACKxSTATUSxRESOLVINGxNAME Or WINHTTPxCALLBACKxSTATUSxNAMExRESOLVED)
Public Const WINHTTPxCALLBACKxFLAGxCONNECTxTOxSERVER = (WINHTTPxCALLBACKxSTATUSxCONNECTINGxTOxSERVER Or WINHTTPxCALLBACKxSTATUSxCONNECTEDxTOxSERVER)
Public Const WINHTTPxCALLBACKxFLAGxSENDxREQUEST = (WINHTTPxCALLBACKxSTATUSxSENDINGxREQUEST Or WINHTTPxCALLBACKxSTATUSxREQUESTxSENT)
Public Const WINHTTPxCALLBACKxFLAGxRECEIVExRESPONSE = (WINHTTPxCALLBACKxSTATUSxRECEIVINGxRESPONSE Or WINHTTPxCALLBACKxSTATUSxRESPONSExRECEIVED)
Public Const WINHTTPxCALLBACKxFLAGxCLOSExCONNECTION = (WINHTTPxCALLBACKxSTATUSxCLOSINGxCONNECTION Or WINHTTPxCALLBACKxSTATUSxCONNECTIONxCLOSED)
Public Const WINHTTPxCALLBACKxFLAGxHANDLES = (WINHTTPxCALLBACKxSTATUSxHANDLExCREATED Or WINHTTPxCALLBACKxSTATUSxHANDLExCLOSING)
Public Const WINHTTPxCALLBACKxFLAGxDETECTINGxPROXY = WINHTTPxCALLBACKxSTATUSxDETECTINGxPROXY
Public Const WINHTTPxCALLBACKxFLAGxREDIRECT = WINHTTPxCALLBACKxSTATUSxREDIRECT
Public Const WINHTTPxCALLBACKxFLAGxINTERMEDIATExRESPONSE = WINHTTPxCALLBACKxSTATUSxINTERMEDIATExRESPONSE
Public Const WINHTTPxCALLBACKxFLAGxSECURExFAILURE = WINHTTPxCALLBACKxSTATUSxSECURExFAILURE
Public Const WINHTTPxCALLBACKxFLAGxSENDREQUESTxCOMPLETE = WINHTTPxCALLBACKxSTATUSxSENDREQUESTxCOMPLETE
Public Const WINHTTPxCALLBACKxFLAGxHEADERSxAVAILABLE = WINHTTPxCALLBACKxSTATUSxHEADERSxAVAILABLE
Public Const WINHTTPxCALLBACKxFLAGxDATAxAVAILABLE = WINHTTPxCALLBACKxSTATUSxDATAxAVAILABLE
Public Const WINHTTPxCALLBACKxFLAGxREADxCOMPLETE = WINHTTPxCALLBACKxSTATUSxREADxCOMPLETE
Public Const WINHTTPxCALLBACKxFLAGxWRITExCOMPLETE = WINHTTPxCALLBACKxSTATUSxWRITExCOMPLETE
Public Const WINHTTPxCALLBACKxFLAGxREQUESTxERROR = WINHTTPxCALLBACKxSTATUSxREQUESTxERROR


Public Const WINHTTPxCALLBACKxFLAGxGETPROXYFORURLxCOMPLETE = WINHTTPxCALLBACKxSTATUSxGETPROXYFORURLxCOMPLETE

Public Const WINHTTPxCALLBACKxFLAGxALLxCOMPLETIONS = (WINHTTPxCALLBACKxSTATUSxSENDREQUESTxCOMPLETE Or _
                 WINHTTPxCALLBACKxSTATUSxHEADERSxAVAILABLE Or WINHTTPxCALLBACKxSTATUSxDATAxAVAILABLE Or _
                 WINHTTPxCALLBACKxSTATUSxREADxCOMPLETE Or WINHTTPxCALLBACKxSTATUSxWRITExCOMPLETE Or _
                 WINHTTPxCALLBACKxSTATUSxREQUESTxERROR Or WINHTTPxCALLBACKxSTATUSxGETPROXYFORURLxCOMPLETE)
Public Const WINHTTPxCALLBACKxFLAGxALLxNOTIFICATIONS = &HFFFFFFFF

'//
'// if the following value is returned by WinHttpSetStatusCallback, then
'// probably an invalid (non-code) address was supplied for the callback
'//
'Public Const WINHTTPxINVALIDxSTATUSxCALLBACK      =  ((WINHTTPxSTATUSxCALLBACK)(-1L))
Public Const WINHTTPxINVALIDxSTATUSxCALLBACK As Long = -1




' ====================================================
' API functions
' ====================================================
' DO Use StrPtr to pass strings,
' DO NOT append Chr(0) to strings before passing them
' ====================================================

Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function WinHttpOpen Lib "winhttp" ( _
   ByVal pszAgentW As LongPtr, _
   ByVal dwAccessType As Long, _
   ByVal pszProxyW As LongPtr, _
   ByVal pszProxyBypassW As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpConnect Lib "winhttp" ( _
   ByVal hSession As LongPtr, _
   ByVal pswzServerName As LongPtr, _
   ByVal nServerPort As Long, _
   ByVal dwReserved As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp" ( _
   ByVal hConnect As LongPtr, _
   ByVal pwszVerb As LongPtr, _
   ByVal pwszObjectName As LongPtr, _
   ByVal pwszVersion As LongPtr, _
   ByVal pwszReferrer As LongPtr, _
   ByVal ppwszAcceptTypes As LongPtr, _
   ByVal dwFlags As Long _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpSetOption Lib "winhttp" ( _
   ByVal hInternet As LongPtr, _
   ByVal dwOption As Long, _
   ByVal lpBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpSendRequest Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpszHeaders As LongPtr, _
   ByVal dwHeadersLength As Long, _
   ByVal lpOptional As LongPtr, _
   ByVal dwOptionalLength As Long, _
   ByVal dwTotalLength As Long, _
   ByVal dwContext As Long _
   ) As Long
   
Public Declare PtrSafe Function WinHttpQueryDataAvailable Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpdwNumberOfBytesAvailable As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpReadData Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByRef pvBuffer As Any, _
   ByVal dwBufferLength As Long, _
   ByRef pdwBytesRead As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal lpReserved As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "winhttp" ( _
   ByVal hRequest As LongPtr, _
   ByVal pContext As LongPtr _
   ) As LongPtr

Public Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp" ( _
   ByVal hRequest As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketSend Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal eBufferType As Long, _
   ByVal pvBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketReceive Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef pvBuffer As Any, _
   ByVal dwBufferLength As Long, _
   ByRef pdwBytesRead As Long, _
   ByRef peBufferType As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketClose Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal usStatus As Integer, _
   ByVal pvReason As LongPtr, _
   ByVal dwReasonLength As Long _
   ) As Long

Public Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "winhttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef usStatus As Integer, _
   ByRef pvReason As Any, _
   ByVal dwReasonLength As Long, _
   ByRef pdwReasonLengthConsumed As LongPtr _
   ) As Long

Public Declare PtrSafe Function WinHttpQueryHeaders Lib "winhttp" ( _
  ByVal hRequest As LongPtr, _
  ByVal dwInfoLevel As Long, _
  ByVal pwszName As LongPtr, _
  ByRef lpBuffer As Long, _
  ByRef lpdwBufferLength As Long, _
  ByRef lpdwIndex As Long _
   ) As Long

' Pointer to an array of WEBxSOCKETxBUFFER structures that contain WebSocket buffer data.

Public Declare PtrSafe Function WebSocketGetAction Lib "websocket" ( _
  ByVal hWebSocket As LongPtr, _
  ByVal eActionQueue As Long, _
  ByVal pDataBuffers As LongPtr, _
  ByRef pulDataBufferCount As Long, _
  ByRef pAction As Long, _
  ByRef pBufferType As Long, _
  ByRef pvApplicationContext As LongPtr, _
  ByRef pvActionContext As LongPtr _
   ) As Long


' WinHttpSetStatusCallback function...

Public Declare PtrSafe Function WinHttpSetStatusCallback Lib "winhttp" ( _
  ByVal hWebSocket As LongPtr, _
  ByVal lpfnInternetCallback As LongPtr, _
  ByVal dwNotificationFlags As Long, _
  ByVal dwReserved As LongPtr _
   ) As Long


' ................................
' WinHttpStatusCallback function...

'SUB CALLBACK WINHTTP_STATUS_CALLBACK (
Public Sub WINHTTPxSTATUSxCALLBACK( _
              ByVal hInternet As LongPtr, _
              ByVal dwContext As Long, _
              ByVal dwInternetStatus As Long, _
              ByVal lpvStatusInformation As LongPtr, _
              ByVal dwStatusInformationLength As Long _
              )

' do things with this information
Debug.Print "Callback: " & dwContext & ":" & dwInternetStatus & ":" & 9&; lpvStatusInformation & ":" & 9&; dwStatusInformationLength
End Sub

' http://www.jose.it-berater.org/winhttp/winhttp_status_callback.htm
' I think this is a model for VBA code that would be called "back"?

' ERROR_WINHTTP_INCORRECT_HANDLE_TYPE = 12018
' ERROR_WINHTTP_INTERNAL_ERROR = 12004
' ERROR_NOT_ENOUGH_MEMORY = 8


'SUB CALLBACK WINHTTP_STATUS_CALLBACK (

'Private Sub WINHTTPxSTATUSxCALLBACK( _
'                      ByVal hInternet As LongPtr, _
'                      ByVal dwContext As Long, _
'                      ByVal dwInternetStatus As Long, _
'                      ByVal lpvStatusInformation As LongPtr, _
'                      ByVal dwStatusInformationLength As Long _
'                      )

' do things with this information
'Debug.Print "Callback: " & dwContext & ":" & dwInternetStatus & ":" & 9&; lpvStatusInformation & ":" & 9&; dwStatusInformationLength
'End Sub

'Sub sillySub()
'call with...
'Dim x As Long, iStatus As Long
'x = WinHttpSetStatusCallback(hWebSocket, AddressOf WINHTTPxSTATUSxCALLBACK, WINHTTPxCALLBACKxSTATUSxREADxCOMPLETE, 0)
'End Sub

'dwNotificationFlags - long
' WINHTTPxCALLBACKxFLAGxALLxCOMPLETIONS - request all things completed
' WINHTTPxCALLBACKxFLAGxALLxNOTIFICATIONS - all status things, including completion
' WINHTTPxCALLBACKxFLAGxREADxCOMPLETE   -  Activates upon completion of a data-read operation. - websockets??????
' WINHTTPxCALLBACKxFLAGxRECEIVExRESPONSE  - Activates upon beginning and completing the receipt of a resource from the HTTP server. - websockets???


' thing list ( WINHTTPxSTATUSxCALLBACK enumeration, below )
'WINHTTPxCALLBACKxSTATUSxWRITExCOMPLETE
'Data was successfully written to the server. The lpvStatusInformation parameter contains a pointer to a DWORD that indicates the number of bytes written.

'When used by WinHttpWebSocketSend, the lpvStatusInformation parameter contains a pointer to a WINHTTPxWEBxSOCKETxSTATUS structure, and the dwStatusInformationLength parameter indicates the size of lpvStatusInformation.




' ................................

' http://www.jose.it-berater.org/winhttp/winhttp_status_callback.htm
' I think this is a model for VBA code that would be called "back"?

'  SUB CALLBACK WINHTTPxSTATUSxCALLBACK ( _
'                      BYVAL hInternet AS DWORD, _
'                      BYVAL dwContext AS DWORD, _
'                      BYVAL dwInternetStatus AS DWORD, _
'                      BYVAL lpvStatusInformation AS DWORD, _
'                      BYVAL dwStatusInformationLength AS DWORD _
'                      )



'//
'// callback function for WinHttpSetStatusCallback
'//

'typedef
'VOID
'(CALLBACK * WINHTTPxSTATUSxCALLBACK)(
'    IN HINTERNET hInternet,
'    IN DWORD_PTR dwContext,
'    IN DWORD dwInternetStatus,
'    IN LPVOID lpvStatusInformation OPTIONAL,
'    IN DWORD dwStatusInformationLength
'    );

'typedef WINHTTPxSTATUSxCALLBACK * LPWINHTTPxSTATUSxCALLBACK;





















