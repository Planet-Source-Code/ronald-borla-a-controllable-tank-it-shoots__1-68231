Attribute VB_Name = "modMath"
Option Explicit

Private Const PI As Double = 3.14159265358979

Public Function CSin(ByVal Degrees As Double) As Double
CSin = Sin(Degrees * (PI / 180))
End Function

Public Function CCos(ByVal Degrees As Double) As Double
CCos = Cos(Degrees * (PI / 180))
End Function

Public Function CTan(ByVal Degrees As Double) As Double
CTan = Tan(Degrees * (PI / 180))
End Function
