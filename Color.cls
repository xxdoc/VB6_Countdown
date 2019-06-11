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


Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal Palet As Long, col As Long) As Long
Private a As Long
Private r As Long
Private g As Long
Private b As Long



Public Sub RGB(red As Integer, green As Integer, blue As Integer)
    r = red
    g = green
    b = blue´
End Sub

Public Sub RGBStr(col As String)
    On Error GoTo ErrorHandler
    Dim splittedcol As String
    splittedcol = Split(col, " ")
    r = splittedcol(0)
    g = splittedcol(1)
    b = splittedcol(2)
ErrorHandler:
    MsgBox "String is not RGB fromat 'R G B (example : 255 0 0 0)' or wrong datatype. Reinit Class to R=0 G=0 B=0", vbCritical
    Class_Initialize

End Sub

Public Sub ARGBStr(col As String)
    On Error GoTo ErrorHandler
    Dim splittedcol As String
    splittedcol = Split(col, " ")
    a = splittedcol(0)
    r = splittedcol(1)
    g = splittedcol(2)
    b = splittedcol(3)

ErrorHandler:
    MsgBox "String is not ARGB fromat 'A R G B (example : 255 0 0 0)' or wrong datatype. Reinit Class to A=255 R=0 G=0 B=0", vbCritical
    Class_Initialize


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


Function GetSystemColor(ByVal Color As OLE_COLOR) As Long
    Dim real As Long
    real = TranslateColor(Color, 0, GetSystemColor)
    ' Handle errors here
    If real <> 0 Then
        MsgBox "Cannot convert to RGB", vbCritical
    End If
End Function


Public Function OleColorToRgb(ByVal Color As OLE_COLOR) As String
    Dim converted As String
    Dim rgb_col() As String

    converted = HexToRGB(GetSystemColor(Color))
    rgb_col = Split(converted, " ")

    r = CStr(rgb_col(0))
    g = CStr(rgb_col(1))
    b = CStr(rgb_col(2))

    OleColorToRgb = CStr(r) + " " + CStr(g) + " " + CStr(b)



End Function

Public Function HexToRGB(c As Long) As String
    r = c Mod 256
    g = c \ 256 Mod 256
    b = c \ 65536 Mod 256

    HexToRGB = CStr(r) + " " + CStr(g) + " " + CStr(b)

End Function

Public Function toString() As String
toString = a + " " + r + " " + b + " " + g
End Function

Public Function toStringRGB() As String
toStringRGB = r + " " + b + " " + g
End Function

Public Function toStringARGB() As String
    toStringARGB = toString
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
