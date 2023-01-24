Attribute VB_Name = "Module4"
'-----------------------
'   Twips to Pixels
'-----------------------
'
'Option Compare Database
Option Explicit

Private Declare Function apiGetDC Lib "USER32" Alias "GetDC" _
    (ByVal hWnd As Long) As Long
Private Declare Function apiReleaseDC Lib "USER32" Alias "ReleaseDC" _
    (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function apiGetDeviceCaps Lib "GDI32" Alias "GetDeviceCaps" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Public Const DIRECTION_VERTICAL = 1
Public Const DIRECTION_HORIZONTAL = 0

Function fTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
'   Function to convert Twips to pixels for the current screen resolution
'   Accepts:
'       lngTwips - the number of twips to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of pixels corresponding to the given twips
    On Error GoTo E_Handle
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = apiGetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
        lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = apiReleaseDC(0, lngDeviceHandle)
    fTwipsToPixels = lngTwips / 1440 * lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_Handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.number
    Resume fExit
End Function

Function fPixelsToTwips(lngPixels As Long, lngDirection As Long) As Long
'   Function to convert pixels to twips for the current screen resolution
'   Accepts:
'       lngPixels - the number of pixels to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:

'       the number of twips corresponding to the given pixels
    On Error GoTo E_Handle
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = apiGetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
    lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = apiReleaseDC(0, lngDeviceHandle)
    fPixelsToTwips = lngPixels * 1440 / lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_Handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.number
    Resume fExit
End Function
 





