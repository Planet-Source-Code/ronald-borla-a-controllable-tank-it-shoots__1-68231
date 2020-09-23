VERSION 5.00
Begin VB.Form frmTank 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcHead 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      Picture         =   "frmTank.frx":0000
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox pcBody 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      Picture         =   "frmTank.frx":14AE
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer tmrMove 
      Interval        =   1
      Left            =   1695
      Top             =   1890
   End
End
Attribute VB_Name = "frmTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Rad As Currency = 1.74532925199433E-02

Dim hAngle As Currency, bAngle As Currency
Dim Speed As Integer
Dim KeyboardMap As String
Dim Shooting As Boolean

Private Sub RotateNow(ByVal hDir As Integer, ByVal bDir As Integer, ByVal hSize As Integer, ByVal bSize As Integer)
DoEvents
Me.Cls
bAngle = bAngle + (bDir * bSize)
hAngle = hAngle + (hDir * hSize)
TranspRotate Me.hdc, bAngle * Rad, hAngle * Rad, pcHead.Left + pcHead.Width / 2, pcHead.Top + pcHead.Height / 2, _
             pcHead.Width, pcBody.Height, pcBody.Image.handle, pcHead.Image.handle, vbMagenta
Me.Refresh
End Sub

Private Sub Form_Click()
Unload frmBullet
Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
hAngle = 180
bAngle = 180
Me.ScaleWidth = pcHead.Width
MakeTrans Me, vbWhite
Dim intX As Integer
KeyboardMap = ""
For intX = 1 To 255
    KeyboardMap = KeyboardMap + "-"
Next intX
Speed = 50
RotateNow 0, 0, 0, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hDir As Integer, bDir As Integer, hSize As Integer, bSize As Integer
Mid$(KeyboardMap, KeyCode, 1) = Chr$(KeyCode)
If Mid$(KeyboardMap, Asc("A"), 1) = "A" Then
    If Mid$(KeyboardMap, Asc("D"), 1) = "D" Then
        AccelerateNow Speed
    Else
        AccelerateNow Speed
        bDir = -1: bSize = 5
    End If
End If
If Mid$(KeyboardMap, Asc("D"), 1) = "D" Then
    If Mid$(KeyboardMap, Asc("A"), 1) = "A" Then
        AccelerateNow Speed
    Else
        AccelerateNow Speed
        bDir = 1: bSize = 5
    End If
End If
If Mid$(KeyboardMap, Asc("J"), 1) = "J" Then
    If Mid$(KeyboardMap, Asc("L"), 1) = "L" Then
        ShootNow
    Else
        hDir = 1: hSize = 5
    End If
End If
If Mid$(KeyboardMap, Asc("L"), 1) = "L" Then
    If Mid$(KeyboardMap, Asc("J"), 1) = "J" Then
        ShootNow
    Else
        hDir = -1: hSize = 5
    End If
End If
If Mid$(KeyboardMap, Asc("K"), 1) = "K" Then
    ShootNow
    Mid$(KeyboardMap, KeyCode, 1) = "-"
End If
If Mid$(KeyboardMap, Asc(Chr(KeyCode)), 1) <> "-" Then
    RotateNow hDir, bDir, hSize, bSize
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Mid$(KeyboardMap, KeyCode, 1) = "-"
End Sub

Private Sub ShootNow()
If Shooting Then Exit Sub
Shooting = True
frmBullet.Show
frmBullet.Left = ((Me.Left + Me.Width / 2) + ((Me.Width / 2) * -CCos(hAngle))) - frmBullet.Width / 2
frmBullet.Top = ((Me.Top + Me.Height / 2) + ((Me.Height / 2) * CSin(hAngle))) - frmBullet.Height / 2
Launch 50
Shooting = False
End Sub

Private Sub AccelerateNow(Speed As Integer)
Dim xS As Integer, yS As Integer
xS = CInt(CCos(bAngle) * Speed)
yS = CInt(CSin(bAngle) * Speed)
Me.Move Me.Left + -xS, Me.Top + yS
End Sub

Private Sub Launch(Speed As Integer)
Dim xS As Integer, yS As Integer
xS = CInt(CCos(hAngle) * Speed)
yS = CInt(CSin(hAngle) * Speed)
With frmBullet
    While (.Top + .Height >= 0) And (.Top <= Screen.Height) And (.Left + .Width >= 0) And (.Left <= Screen.Width)
        .Move .Left + -xS, .Top + yS
        Pause
    Wend
    .Hide
End With
End Sub

Private Sub Pause()
Dim i As Long
For i = 1 To 100
    DoEvents
Next i
End Sub
