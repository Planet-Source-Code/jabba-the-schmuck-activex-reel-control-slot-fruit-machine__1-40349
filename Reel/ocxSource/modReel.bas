Attribute VB_Name = "modReel"
'Ensure all variables are declared.
Option Explicit

'Can be used as a timer, much better.
Declare Function GetTickCount Lib "kernel32" () As Long

'Optimize loops by only calling DoEvents if this routine returns zero.
Declare Function GetInputState Lib "user32" () As Long

'Used for the graphics/drawing side of things.
Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

'Flags for the BitBlt API.
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

Public Sub Incr(ByRef what As Long, Optional ByVal value As Integer = 1)

    'Increment the variable by a given amount.  (Default 1)
    what = what + value

End Sub

