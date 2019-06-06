VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private a As Long
Private r As Long
Private g As Long
Private b As Long



Public Sub RGB(red As Integer, green As Integer, blue As Integer)
r = red
g = green
b = blue´
End Sub
Private Sub Class_Initialize()
 a = 255
 r = 0
 g = 0
 b = 0
End Sub
Public Sub ARGB(alpha As Integer, red As Integer, green As Integer, blue As Integer)
a = alpha
r = red
g = green
b = blue´
End Sub

Public Sub ConvertToRGB(c As Long)
r = c Mod 256
g = c \ 256 Mod 256
b = c \ 65536 Mod 256

r = r * -1
g = g * -1
b = b * -1

End Sub

     Function GetRGB(RGBval As Long, Num As Integer) As Integer
         ' Check if Num, RGBval are valid.
         If Num > 0 And Num < 4 And RGBval > -1 And RGBval < 16777216 Then
           GetRGB = RGBval \ 256 ^ (Num - 1) And 255
         Else
           ' Return True (-1) if Num or RGBval are invalid.
           GetRGB = True
         End If
      End Function

Public Function get_a() As Integer
get_a = a
End Function

Public Function get_r() As Integer
get_r = r
End Function

Public Function get_g() As Integer
get_g = g
End Function

Public Function get_b() As Integer
get_b = b
End Function

Public Sub set_a(alpha As Integer)
a = alpha
End Sub

Public Sub set_r(red As Integer)
r = red
End Sub

Public Sub set_g(green As Integer)
green = green
End Sub

Public Sub set_b(blue As Integer)
b = blue
End Sub
