Attribute VB_Name = "ModCircleProg"

'the api for bitblt \/
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

Const ShadeChange = 50 'the shade change from 'light' to 'dark'

Function GetBlue(CVal) As Long
'the blue value of a color
GetBlue = Int(CVal / 65536)
End Function

Function GetGreen(CVal) As Long
'get the green value of a color
GetGreen = Int((CVal - ((65536) * GetBlue(CVal))) / 255)
End Function

Function GetRed(CVal) As Long
'get the red value of a color
GetRed = CVal - (65536 * GetBlue(CVal) + 256 * GetGreen(CVal))
End Function

'a function to make sure the rgb value numbers dont
'go out of bound
Function EndNum(val As Long)
If val < 0 Then val = 0
If val > 255 Then val = 255
End Function

'a function to make a darker color of another color
Function GetDarkerColor(CVal As Long) As Long
Dim r As Long, g As Long, b As Long 'declare an r, g, and b var

r = GetRed(CVal) 'get the values
g = GetGreen(CVal)
b = GetBlue(CVal)

r = r - ShadeChange 'darken the values
g = g - ShadeChange
b = b - ShadeChange

EndNum r 'shift the numbers into bounds
EndNum g
EndNum b

GetDarkerColor = RGB(r, g, b) 'get the final color
End Function

