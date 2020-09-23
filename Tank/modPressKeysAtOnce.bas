Attribute VB_Name = "modPressKeysAtOnce"
'**************************************
' Name: Capture all the keys that are pr
'     essed.
' Description:Allows users to press mult
'     iple keys on the keyboard. This is very
'     useful for when making games.
'EG Forward, Left or Right, and Fire - Shields also, all at the SAME time!
' By: Scott Gunn
'
'
' Inputs:Just keyboard presses.
'
' Returns:A 255 Character String saying
'     what keys are pressed.
'
'Assumes:None
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.5851/lngWId.1/qx/
'     vb/scripts/ShowCode.htm
'for details.
'**************************************

Option Explicit

Dim KeyboardMap As String


Private Sub Form_Load()
    Dim intX As Integer
    'Set KeyboardMap to make say all keys un
    '     pressed.
    KeyboardMap = ""


    For intX = 1 To 255
        KeyboardMap = KeyboardMap + "-" 'Set to a char you will not use
    Next intX
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Set the character at the position of th
    '     e character of the pressed key
    Mid$(KeyboardMap, KeyCode, 1) = Chr$(KeyCode)
    Label1.Caption = KeyboardMap
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Set the character at the position of th
    '     e pressed key value to " "
    Mid$(KeyboardMap, KeyCode, 1) = "-" 'Set to a char you will not use
    Label1.Caption = KeyboardMap
End Sub
'Examples of detection routines
'if Mid$(KeyboardMap,Asc("A"),1) = "A" t
'     hen the A key is pressed
'if Mid$(KeyboardMap,Asc("1"),1) = "1" t
'     hen the 1 key is pressed

        

