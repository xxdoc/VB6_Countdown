Attribute VB_Name = "modChooseFontDilettante"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(31) As Integer
End Type

Private Type tagCHOOSEFONT
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hdc As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String   '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String        '  return the style field here
    nFontType As Integer       '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type

Private Type FONTDESC
    cbSizeofstruct As Long
    lpstrName As Long
    cySize As Currency
    sWeight As Integer
    sCharset As Integer
    fItalic As Long
    fUnderline As Long
    fStrikethrough As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreateFontIndirect Lib "olepro32" (pFontDesc As FONTDESC, riid As Guid, ppvObj As IFont) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontW" (pChoosefont As tagCHOOSEFONT) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub IIDFromString Lib "ole32" (ByVal lpsz As Long, lpiid As Any)
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSY        As Long = 90        '  Logical pixels/inch in Y
Private Const SW_HIDE           As Long = 0
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOMOVE        As Long = &H2
Private Const WM_INITDIALOG     As Long = &H110
Private Const WM_DESTROY        As Long = &H2
Private Const CF_ENABLEHOOK     As Long = 8
Private Const CF_EFFECTS        As Long = &H100
Private Const LF_FACESIZE       As Long = 32
Private Const cmb4              As Long = &H473
Private Const grp1              As Long = &H430
Private Const stc6              As Long = &H445
Private Const IID_IFont         As String = "{BEF6E002-A874-101A-8BBA-00AA00300CAB}"

Dim cf  As tagCHOOSEFONT

Public Function SelectFont(col As Long) As IFont
    Dim lf  As LOGFONT
    Dim iid As Guid
    Dim fd  As FONTDESC
    Dim hdc As Long
       
    cf.lStructSize = Len(cf)
    cf.hwndOwner = frmMain.hwnd
    cf.lpLogFont = VarPtr(lf)
    cf.lpfnHook = GetAddr(AddressOf CFHookProc)
    cf.flags = CF_ENABLEHOOK Or CF_EFFECTS
    
    If ChooseFont(cf) Then
        IIDFromString StrPtr(IID_IFont), iid
        
        hdc = GetDC(0)
        
        fd.cbSizeofstruct = Len(fd)
        fd.fItalic = lf.lfItalic
        fd.cySize = MulDiv(-lf.lfHeight, 72, GetDeviceCaps(hdc, LOGPIXELSY))
        fd.fStrikethrough = lf.lfStrikeOut
        fd.fUnderline = lf.lfUnderline
        fd.lpstrName = VarPtr(lf.lfFaceName(0))
        fd.sCharset = lf.lfCharSet
        fd.sWeight = lf.lfWeight
        
        ReleaseDC 0, hdc
        
        OleCreateFontIndirect fd, iid, SelectFont
        
        col = frmMain.lblColor.BackColor
        
    End If
    
End Function

Public Function GetAddr(ByVal Addr As Long) As Long
    GetAddr = Addr
End Function

Public Function CFHookProc(ByVal hwnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uiMsg
    Case WM_INITDIALOG
        Dim hCbo As Long:    Dim hMy     As Long
        Dim rc   As RECT:    Dim hGrp As Long
        Dim cboHght As Long: Dim pic As PictureBox
        
        Set pic = frmMain.picContainer
        hMy = pic.hwnd
        pic.Visible = True
        hCbo = GetDlgItem(hwnd, cmb4)
        SetParent hMy, hwnd
        GetClientRect hCbo, rc: MapWindowPoints hCbo, hwnd, rc, 2
        cboHght = rc.Bottom - rc.Top
        SetWindowPos hMy, 0, rc.Left, rc.Top, rc.Right - rc.Left, pic.ScaleHeight, SWP_NOZORDER
        hGrp = GetDlgItem(hwnd, grp1)
        GetClientRect hGrp, rc: MapWindowPoints hGrp, hwnd, rc, 2
        SetWindowPos hGrp, 0, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top + pic.ScaleHeight - cboHght, SWP_NOZORDER
        ShowWindow hCbo, SW_HIDE
        ShowWindow GetDlgItem(hwnd, stc6), SW_HIDE
        GetWindowRect hwnd, rc
        SetWindowPos hwnd, 0, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top + pic.ScaleHeight - cboHght, SWP_NOZORDER
    Case WM_DESTROY
        SetParent frmMain.picContainer.hwnd, frmMain.hwnd
        frmMain.picContainer.Visible = False
    End Select
    
End Function

