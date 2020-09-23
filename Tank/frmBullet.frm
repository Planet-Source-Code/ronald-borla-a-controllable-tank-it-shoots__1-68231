VERSION 5.00
Begin VB.Form frmBullet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   LinkTopic       =   "Form1"
   ScaleHeight     =   195
   ScaleWidth      =   195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   15
      Shape           =   3  'Circle
      Top             =   15
      Width           =   195
   End
End
Attribute VB_Name = "frmBullet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeTrans Me, vbWhite
End Sub
