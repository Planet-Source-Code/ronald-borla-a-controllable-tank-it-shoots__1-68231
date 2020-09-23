Attribute VB_Name = "RotateImage"

Option Explicit
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Public Const DIB_PAL_COLORS = 1
Public Const DIB_PAL_INDICES = 2
Public Const DIB_PAL_LOGINDICES = 4
Public Const DIB_PAL_PHYSINDICES = 2
Public Const DIB_RGB_COLORS = 0
Public Const SRCCOPY = &HCC0020
Public Type BITMAPINFOHEADER
    biSize           As Long
    biWidth          As Long
    biHeight         As Long
    biPlanes         As Integer
    biBitCount       As Integer
    biCompression    As Long
    biSizeImage      As Long
    biXPelsPerMeter  As Long
    biYPelsPerMeter  As Long
    biClrUsed        As Long
    biClrImportant   As Long
End Type

Public Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Bits() As Byte             '(Colors)
End Type

Sub TranspRotate(Destdc As Long, Angle1 As Currency, Angle2 As Currency, _
                 x&, y&, W&, H&, ImgHandle1&, ImgHandle2&, Optional TranspColor&, _
                 Optional Alpha As Currency = 1, Optional pScale As Currency = 1, _
                 Optional px% = -32767, Optional py% = -32767)
  'Angle given is in rads
  'DoEvents
  Dim P1() As Byte
  Dim P2() As Byte
  Dim ProcessedBits() As Byte
  
  Dim dx1 As Currency, dy1 As Currency, tx1 As Currency, ty1 As Currency
  Dim dx2 As Currency, dy2 As Currency, tx2 As Currency, ty2 As Currency
  
  Dim ix1 As Integer, iy1 As Integer
  Dim ix2 As Integer, iy2 As Integer
  
  Dim Tmp&, cX&, CY&, XX&, YY&
  Dim TR As Byte, TB As Byte, TG As Byte
  Dim D() As Byte
  Dim BackDC As Long
  Dim BackBmp As BITMAPINFO
  Dim iBitmap As Long
  Dim TopL As Currency, TopR As Currency, BotL As Currency, BotR As Currency
  Dim TopLV As Currency, TopRV As Currency, BotLV As Currency, BotRV As Currency
  Dim pSin1 As Currency, PCos1 As Currency
  Dim pSin2 As Currency, pCos2 As Currency
  Dim PicBmp As BITMAPINFO, PicDC As Long
  
  'Get the maximum width and heigth any rotation can produce
  Tmp = Int(Sqr(W * W + H * H)) * pScale
  
  'Prepare the pixel arrays
  ReDim D(3, Tmp - 1, Tmp - 1)   'Holds The background image
  ReDim P1(3, W - 1, H - 1)     'Holds the source image
  ReDim P2(3, W - 1, H - 1)     'Holds the source image
  ReDim ProcessedBits(3, Tmp - 1, Tmp - 1) 'Holds the rotated result
  
  'Set the rotation axis default values
  If px = -32767 Then px = (W / 2)  'pivot x
  If py = -32767 Then py = (H / 2)  'pivot y
  
  '[Create a Context - Copy the Backgroung - Get Background pixels]
  With BackBmp.Header
      .biBitCount = 4 * 8
      .biPlanes = 1
      .biSize = 40
      .biWidth = Tmp
      .biHeight = -Tmp
  End With
  'Create a context
  BackDC = CreateCompatibleDC(0)
  'Create a blank picture on the BackBmp standards
  iBitmap = CreateDIBSection(BackDC, BackBmp, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
  'Copy the picture in to the context to make the context useable just like a picturebox
  SelectObject BackDC, iBitmap
  'Copy the background to the context
  BitBlt BackDC, 0, 0, Tmp, Tmp, Destdc, x - Tmp / 2, y - Tmp / 2, SRCCOPY
  'Analize the background pixels and save to the array D
  GetDIBits BackDC, iBitmap, 0, Tmp, D(0, 0, 0), BackBmp, DIB_RGB_COLORS
  
  '[Get SourceImage Pixels]
  With PicBmp.Header
      .biBitCount = 4 * 8
      .biPlanes = 1
      .biSize = 40
      .biWidth = W
      .biHeight = -H
  End With
  'Create a context
  PicDC = CreateCompatibleDC(0)
  'Copy the sourceimage in to the context to make the context useable (no need to create a new dibsection since the SourcePicture image is compatible)
  SelectObject PicDC, ImgHandle1
  'Analize the sourceimage pixels and save to the array D
  GetDIBits PicDC, ImgHandle1, 0, H, P1(0, 0, 0), PicBmp, DIB_RGB_COLORS

  'Create a context
  PicDC = CreateCompatibleDC(0)
  'Copy the sourceimage in to the context to make the context useable (no need to create a new dibsection since the SourcePicture image is compatible)
  SelectObject PicDC, ImgHandle2
  'Analize the sourceimage pixels and save to the array D
  GetDIBits PicDC, ImgHandle2, 0, H, P2(0, 0, 0), PicBmp, DIB_RGB_COLORS
  
  'Get the min values to scan
  cX = Int((Tmp - W) / 2)
  CY = Int((Tmp - H) / 2)
  
  'Convert to R,G,B format the transparent color
  TR = TranspColor And &HFF&
  TG = (TranspColor And &HFF00&) / &H100&
  TB = (TranspColor And &HFF0000) / &H10000
  
  'Precalculate the trigonometry
  PCos1 = Cos(Angle1) / pScale
  pSin1 = Sin(Angle1) / pScale
  pCos2 = Cos(Angle2) / pScale
  pSin2 = Sin(Angle2) / pScale

  'Loop through all pixels of the source image
  For XX = -cX To Tmp - cX - 1
   For YY = -CY To Tmp - CY - 1
      'Get the rotation translation (gives the SourceImage coordinate for each DestImage x,y)
      tx1 = (XX - px) * PCos1 - (YY - py) * pSin1 + px
      ty1 = (XX - px) * pSin1 + (YY - py) * PCos1 + py
      
      tx2 = (XX - px) * pCos2 - (YY - py) * pSin2 + px
      ty2 = (XX - px) * pSin2 + (YY - py) * pCos2 + py
      
      'Get nearest to the left pixel
      ix1 = Int(tx1)
      iy1 = Int(ty1)
      
      ix2 = Int(tx2)
      iy2 = Int(ty2)
      
      'Get the digits after the decimal point
      dx1 = Abs(tx1 - ix1)
      dy1 = Abs(ty1 - iy1)
      
      dx2 = Abs(tx2 - ix2)
      dy2 = Abs(ty2 - iy2)
      
      'Color the destination with the background color
      ProcessedBits(0, XX + cX, YY + CY) = D(0, XX + cX, YY + CY)
      ProcessedBits(1, XX + cX, YY + CY) = D(1, XX + cX, YY + CY)
      ProcessedBits(2, XX + cX, YY + CY) = D(2, XX + cX, YY + CY)

      If tx1 >= 0 And ix1 + 1 < W Then
       If ty1 >= 0 And iy1 + 1 < H Then
           'These variables hold Alpha value if the source pixel is non-transparent
           'If it's transparent they hold zero
           TopLV = -CBool(P1(0, ix1, iy1) <> TR Or P1(1, ix1, iy1) <> TG Or P1(2, ix1, iy1) <> TB) * Alpha
           TopRV = -CBool(P1(0, ix1 + 1, iy1) <> TR Or P1(1, ix1 + 1, iy1) <> TG Or P1(2, ix1 + 1, iy1) <> TB) * Alpha
           BotLV = -CBool(P1(0, ix1, iy1 + 1) <> TR Or P1(1, ix1, iy1 + 1) <> TG Or P1(2, ix1, iy1 + 1) <> TB) * Alpha
           BotRV = -CBool(P1(0, ix1 + 1, iy1 + 1) <> TR Or P1(1, ix1 + 1, iy1 + 1) <> TG Or P1(2, ix1 + 1, iy1 + 1) <> TB) * Alpha
           
           'The SourcePixel color maybe a combination of upto four pixels as tx1 and ty1 are not integers
           'The intersepted (by the current calculated source pixel) area each pixel involved (see .doc for more info)
           TopL = (1 - dx1) * (1 - dy1)
           TopR = dx1 * (1 - dy1)
           BotL = (1 - dx1) * dy1
           BotR = dx1 * dy1
        
           'Simplified explanation of the routine combination:
           'Alphablending (alpha being a real value from 0 to 1): DestColor = SourceImageColor * Alpha + BackImageColor * (1-Alpha)
           'Antialiasing: DestColor = SourceTopLeftPixel * TopLeftAreaIntersectedBySourcePixel +SourceTopRightPixel * TopRightAreaIntersectedBySourcePixel + bottomleft... + bottomrigth...
           
           'The AntiAliased Alpha assigment of colors
           ProcessedBits(0, XX + cX, YY + CY) = (P1(0, ix1, iy1) * TopLV + D(0, XX + cX, YY + CY) * (1 - TopLV)) * TopL + (P1(0, ix1 + 1, iy1) * TopRV + D(0, XX + cX, YY + CY) * (1 - TopRV)) * TopR + (P1(0, ix1, iy1 + 1) * BotLV + D(0, XX + cX, YY + CY) * (1 - BotLV)) * BotL + (P1(0, ix1 + 1, iy1 + 1) * BotRV + D(0, XX + cX, YY + CY) * (1 - BotRV)) * BotR
           ProcessedBits(1, XX + cX, YY + CY) = (P1(1, ix1, iy1) * TopLV + D(1, XX + cX, YY + CY) * (1 - TopLV)) * TopL + (P1(1, ix1 + 1, iy1) * TopRV + D(1, XX + cX, YY + CY) * (1 - TopRV)) * TopR + (P1(1, ix1, iy1 + 1) * BotLV + D(1, XX + cX, YY + CY) * (1 - BotLV)) * BotL + (P1(1, ix1 + 1, iy1 + 1) * BotRV + D(1, XX + cX, YY + CY) * (1 - BotRV)) * BotR
           ProcessedBits(2, XX + cX, YY + CY) = (P1(2, ix1, iy1) * TopLV + D(2, XX + cX, YY + CY) * (1 - TopLV)) * TopL + (P1(2, ix1 + 1, iy1) * TopRV + D(2, XX + cX, YY + CY) * (1 - TopRV)) * TopR + (P1(2, ix1, iy1 + 1) * BotLV + D(2, XX + cX, YY + CY) * (1 - BotLV)) * BotL + (P1(2, ix1 + 1, iy1 + 1) * BotRV + D(2, XX + cX, YY + CY) * (1 - BotRV)) * BotR
       End If
      End If
      
      D(0, XX + cX, YY + CY) = ProcessedBits(0, XX + cX, YY + CY)
      D(1, XX + cX, YY + CY) = ProcessedBits(1, XX + cX, YY + CY)
      D(2, XX + cX, YY + CY) = ProcessedBits(2, XX + cX, YY + CY)
      
      If tx2 >= 0 And ix2 + 1 < W Then
       If ty2 >= 0 And iy2 + 1 < H Then
           'These variables hold Alpha value if the source pixel is non-transparent
           'If it's transparent they hold zero
           TopLV = -CBool(P2(0, ix2, iy2) <> TR Or P2(1, ix2, iy2) <> TG Or P2(2, ix2, iy2) <> TB) * Alpha
           TopRV = -CBool(P2(0, ix2 + 1, iy2) <> TR Or P2(1, ix2 + 1, iy2) <> TG Or P2(2, ix2 + 1, iy2) <> TB) * Alpha
           BotLV = -CBool(P2(0, ix2, iy2 + 1) <> TR Or P2(1, ix2, iy2 + 1) <> TG Or P2(2, ix2, iy2 + 1) <> TB) * Alpha
           BotRV = -CBool(P2(0, ix2 + 1, iy2 + 1) <> TR Or P2(1, ix2 + 1, iy2 + 1) <> TG Or P2(2, ix2 + 1, iy2 + 1) <> TB) * Alpha
           
           'The SourcePixel color maybe a combination of upto four pixels as tx2 and ty2 are not integers
           'The intersepted (by the current calculated source pixel) area each pixel involved (see .doc for more info)
           TopL = (1 - dx2) * (1 - dy2)
           TopR = dx2 * (1 - dy2)
           BotL = (1 - dx2) * dy2
           BotR = dx2 * dy2
        
           'Simplified explanation of the routine combination:
           'Alphablending (alpha being a real value from 0 to 1): DestColor = SourceImageColor * Alpha + BackImageColor * (1-Alpha)
           'Antialiasing: DestColor = SourceTopLeftPixel * TopLeftAreaIntersectedBySourcePixel +SourceTopRightPixel * TopRightAreaIntersectedBySourcePixel + bottomleft... + bottomrigth...
           
           'The AntiAliased Alpha assigment of colors
           ProcessedBits(0, XX + cX, YY + CY) = (P2(0, ix2, iy2) * TopLV + D(0, XX + cX, YY + CY) * (1 - TopLV)) * TopL + (P2(0, ix2 + 1, iy2) * TopRV + D(0, XX + cX, YY + CY) * (1 - TopRV)) * TopR + (P2(0, ix2, iy2 + 1) * BotLV + D(0, XX + cX, YY + CY) * (1 - BotLV)) * BotL + (P2(0, ix2 + 1, iy2 + 1) * BotRV + D(0, XX + cX, YY + CY) * (1 - BotRV)) * BotR
           ProcessedBits(1, XX + cX, YY + CY) = (P2(1, ix2, iy2) * TopLV + D(1, XX + cX, YY + CY) * (1 - TopLV)) * TopL + (P2(1, ix2 + 1, iy2) * TopRV + D(1, XX + cX, YY + CY) * (1 - TopRV)) * TopR + (P2(1, ix2, iy2 + 1) * BotLV + D(1, XX + cX, YY + CY) * (1 - BotLV)) * BotL + (P2(1, ix2 + 1, iy2 + 1) * BotRV + D(1, XX + cX, YY + CY) * (1 - BotRV)) * BotR
           ProcessedBits(2, XX + cX, YY + CY) = (P2(2, ix2, iy2) * TopLV + D(2, XX + cX, YY + CY) * (1 - TopLV)) * TopL + (P2(2, ix2 + 1, iy2) * TopRV + D(2, XX + cX, YY + CY) * (1 - TopRV)) * TopR + (P2(2, ix2, iy2 + 1) * BotLV + D(2, XX + cX, YY + CY) * (1 - BotLV)) * BotL + (P2(2, ix2 + 1, iy2 + 1) * BotRV + D(2, XX + cX, YY + CY) * (1 - BotRV)) * BotR
       End If
      End If
   Next
  Next
  
  'Draw the pixel array
  StretchDIBits Destdc, x - Tmp / 2, y - Tmp / 2, Tmp, Tmp, 0, 0, Tmp, Tmp, ProcessedBits(0, 0, 0), BackBmp, DIB_RGB_COLORS, SRCCOPY
  'Clear the variables
  Erase D
  Erase ProcessedBits
  Erase P1
  Erase P2
End Sub

