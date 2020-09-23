Attribute VB_Name = "modTrans"
Option Explicit
Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum TransType
  LWA_OPAQUE = 0
  LWA_COLORKEY = 1
  LWA_ALPHA = 2
End Enum

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Private Const zFormOrPictBoxStr = "Must pass in the name of either a Form or a PictureBox."

Private Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Function MakeTrans(zForm As Form, Optional ByVal TransColor As Long = &HFF00FF) As Boolean
  On Local Error Resume Next
  Dim Msg As Long
  Msg = GetWindowLong(zForm.hwnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong zForm.hwnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.hwnd, TransColor, 0, LWA_COLORKEY
  If Err Then
    MakeTrans = False
  Else
    MakeTrans = True
  End If
End Function
