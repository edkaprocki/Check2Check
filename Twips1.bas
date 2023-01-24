Attribute VB_Name = "Twips"
Option Explicit

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
  ByVal hdc As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
  ByVal nIndex As Long) As Long

Const WU_LOGPIXELSX = 88
Const WU_LOGPIXELSY = 90


Function ConvertTwipsToPixels(lngTwips As Long, _
   lngDirection As Long) As Long

   'Handle to device
   Dim lngDC As Long
   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)
   
   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch

End Function
