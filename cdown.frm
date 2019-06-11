VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Countdown 
   Caption         =   "Countdown"
   ClientHeight    =   3915
   ClientLeft      =   2400
   ClientTop       =   2715
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   8865
   Begin VB.CommandButton btnPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   7320
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   6840
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox tboxdays 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "0"
      Top             =   3240
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   75628545
      CurrentDate     =   43622
   End
   Begin VB.Timer Timer1 
      Left            =   11400
      Top             =   240
   End
   Begin VB.Label Tage 
      Caption         =   "Tage"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   54.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
   Begin VB.Menu Settings 
      Caption         =   "Settings"
      Begin VB.Menu color_menu 
         Caption         =   "Farbe"
         Begin VB.Menu lbl_menu 
            Caption         =   "Label"
            Begin VB.Menu foreground_color_menu 
               Caption         =   "Vordergrund"
               Shortcut        =   ^{F3}
            End
            Begin VB.Menu background_color_menu 
               Caption         =   "Hintergrund"
               Shortcut        =   ^{F4}
            End
            Begin VB.Menu sep0 
               Caption         =   "-"
            End
            Begin VB.Menu Reset_lblcol 
               Caption         =   "Reset"
            End
         End
         Begin VB.Menu application_menu 
            Caption         =   "Anwendung"
            Begin VB.Menu foreground_color_applicaton_menu 
               Caption         =   "Vordergrund"
            End
            Begin VB.Menu Background_Color_Applicaton_menu 
               Caption         =   "Hintergrund"
            End
            Begin VB.Menu label_color_to_app_color_menue 
               Caption         =   "= Label Farbe"
            End
            Begin VB.Menu sep1 
               Caption         =   "-"
               Index           =   0
            End
            Begin VB.Menu Reset_Appcol 
               Caption         =   "Reset"
            End
         End
      End
      Begin VB.Menu font_category 
         Caption         =   "Font"
         Begin VB.Menu font_menu 
            Caption         =   "Font"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu sep2 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu Reset_Font 
            Caption         =   "Reset"
         End
      End
      Begin VB.Menu show_days_menu 
         Caption         =   "Tage anzeigen"
         Checked         =   -1  'True
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu Reset_all_settings 
         Caption         =   "Reset All"
         Shortcut        =   ^{F5}
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Countdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time As Date
Dim timewas As Date
Dim h_buf As Long
Dim s_buf As Long
Dim m_buf As Long
Dim d_buf As Long


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER As Long = &H2000&    'or 8192 in decimal form
Private Const RESET_NO As Integer = 0
Private Const RESET_ALL As Integer = 1
Private Const RESET_LBL_FONT As Integer = 2
Private Const RESET_LBL_COLOR As Integer = 3
Private Const RESET_APP_COLOR As Integer = 4


Private show_days As Boolean



Private Sub About_Click()
    Load frmAbout
    frmAbout.Left = Me.Left
    frmAbout.Top = Me.Top
    frmAbout.Show vbModal, Me


End Sub

Private Sub Background_Color_Applicaton_menu_Click()
    On Error Resume Next
    With dlgColor
        .CancelError = True

        ' Anfnagsfarbe
        .Color = Me.BackColor

        ' Wichtig: Flag cdlCCRGBInit muss hierzu gesetzt werden!
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor

        If Err.Number = 0 Then
            ' Neue ausgewählte Farbe
            Me.BackColor = dlgColor.Color
            Me.Tage.BackColor = dlgColor.Color
            Me.btnStart.BackColor = dlgColor.Color
            Me.btnPause.BackColor = dlgColor.Color


        End If
    End With
    Dim col As Color
    Set col = New Color
    Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\application", "backcolor", col.OleColorToRgb(dlgColor.Color), Registry.HKEY_CURRENT_USER


    On Error GoTo 0
End Sub

Private Sub set_color(ByRef col As Long)
    On Error Resume Next
    With dlgColor
        .CancelError = True

        ' Anfnagsfarbe
        .Color = Me.BackColor

        ' Wichtig: Flag cdlCCRGBInit muss hierzu gesetzt werden!
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor

        If Err.Number = 0 Then
            ' Neue ausgewählte Farbe
            Me.BackColor = dlgColor.Color
            Me.Tage.BackColor = dlgColor.Color
            Me.btnStart.BackColor = dlgColor.Color
            Me.btnPause.BackColor = dlgColor.Color
            Me.tboxdays.BackColor = dlgColor.Color

        End If
    End With
    On Error GoTo 0
End Sub

Private Sub background_color_menu_Click()
    On Error Resume Next
    With dlgColor
        .CancelError = True

        ' Anfnagsfarbe
        .Color = Label1.BackColor

        ' Wichtig: Flag cdlCCRGBInit muss hierzu gesetzt werden!
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor

        If Err.Number = 0 Then
            ' Neue ausgewählte Farbe
            Label1.BackColor = dlgColor.Color
        End If
    End With

    Dim col As Color
    Set col = New Color
    Registry.SetKeyValue "SOFTWARE\" + App.ProductName _
                         + "\cdownlbl\lblbg", "color", col.OleColorToRgb(dlgColor.Color), Registry.HKEY_CURRENT_USER
    On Error GoTo 0
End Sub

Private Sub btnPause_Click()
    If Timer1.Enabled = False Then
        btnPause.Caption = "Pause"
        Dim sres As Long
        Timer1.Enabled = True
        init_time d_buf, h_buf, m_buf, s_buf

    Else

        btnPause.Caption = "Resume"
        Timer1.Enabled = False

        s_buf = DateDiff("s", Now, time)
        m_buf = s_buf \ 60
        s_buf = s_buf - m_buf * 60
        h_buf = m_buf \ 60
        m_buf = m_buf - h_buf * 60
        d_buf = h_buf \ 24
        h_buf = h_buf - d_buf * 24



    End If

End Sub

Sub init_time(day As Long, hour As Long, minute As Long, second As Long)
    time = Now
    time = DateAdd("h", hour, time)
    time = DateAdd("n", minute, time)
    time = DateAdd("s", second, time)
    time = DateAdd("d", day, time)
End Sub

Private Sub btnStart_Click()

    If btnStart.Caption = "Start" Then
        btnStart.Caption = "End"
        Timer1.Enabled = True
        init_time CLng(tboxdays), hour(DTPicker1.Value), minute(DTPicker1.Value), second(DTPicker1.Value)


    Else
        btnStart.Caption = "Start"
        Timer1.Enabled = False
        time = Now
        If show_days = True Then
            Label1.Caption = "0 Tage 00:00:00"
        Else
            Label1.Caption = "00:00:00"
        End If

    End If



End Sub



Private Sub font_menu_Click()

    On Error Resume Next
    With dlgFont
        .CancelError = True
        .FontName = Label1.FontName
        .FontBold = Label1.FontBold
        .FontItalic = Label1.FontItalic
        .FontSize = Label1.FontSize
        .FontStrikethru = Label1.FontStrikethru
        .FontUnderline = Label1.FontUnderline
        .Color = Label1.ForeColor

        ' Die Flags-Eigenschaft muss auf cdlCFScreenFonts,
        ' cdlCFPrinterFonts oder cdlCFBoth gesetzt werden,
        ' bevor das Dialogfeld Schriftart angezeigt wird,
        ' sonst tritt der Fehler "Keine Schriftarten vorhanden" auf.
        .Flags = cdlCFEffects Or cdlCFBoth

        ' Dialogfeld Schriftart anzeigen
        .ShowFont


        If Err = 0 Then
            ' Text markieren und Benutzereingaben übernehmen
            With Label1
                .FontName = dlgFont.FontName
                .FontBold = dlgFont.FontBold
                .FontItalic = dlgFont.FontItalic
                .FontSize = dlgFont.FontSize
                .FontStrikethru = dlgFont.FontStrikethru
                .FontUnderline = dlgFont.FontUnderline
                .ForeColor = dlgFont.Color
            End With

            Dim col As Color
            Set col = New Color

            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "name", dlgFont.FontName, Registry.HKEY_CURRENT_USER
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "bold", Abs(dlgFont.FontBold), Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "italic", Abs(dlgFont.FontItalic), Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "size", dlgFont.FontSize, Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "strikethru", Abs(dlgFont.FontStrikethru), Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "underline", Abs(dlgFont.FontUnderline), Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "color", col.OleColorToRgb(dlgFont.Color), Registry.HKEY_CURRENT_USER

        End If
    End With
End Sub

Private Sub foreground_color_applicaton_menu_Click()
    On Error Resume Next
    With dlgColor
        .CancelError = True

        ' Anfnagsfarbe
        .Color = Me.ForeColor


        ' Wichtig: Flag cdlCCRGBInit muss hierzu gesetzt werden!
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor

        If Err.Number = 0 Then
            ' Neue ausgewählte Farbe
            Me.ForeColor = dlgColor.Color
            Tage.ForeColor = Me.ForeColor
            Me.tboxdays.ForeColor = Me.ForeColor


        End If
    End With
    
    Dim col As Color
    Set col = New Color

    Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\application", "forecolor", col.OleColorToRgb(dlgColor.Color), Registry.HKEY_CURRENT_USER

    On Error GoTo 0
End Sub

Private Sub foreground_color_menu_Click()
    On Error Resume Next
    With dlgColor
        .CancelError = True

        ' Anfnagsfarbe
        .Color = Label1.ForeColor

        ' Wichtig: Flag cdlCCRGBInit muss hierzu gesetzt werden!
        .Flags = cdlCCFullOpen Or cdlCCRGBInit
        .ShowColor

        If Err.Number = 0 Then
            ' Neue ausgewählte Farbe
            Label1.ForeColor = dlgColor.Color
        End If
    End With
    On Error GoTo 0
End Sub

Private Sub set_zero_time()
    If show_days = True Then
        Me.Label1.Caption = "0 Tage 00:00:00"
    Else
        Me.Label1.Caption = "00:00:00"
    End If

End Sub


Private Sub Form_Load()


    Dim col As Color
    Set col = New Color








    'checks registry key and values if key not created it creates the key
    Dim res As Long



    init_registry
    init_settings




    set_zero_time
    Dim style As Long
    tboxdays.Text = ""
    style = GetWindowLong(tboxdays.hWnd, GWL_STYLE)
    SetWindowLong tboxdays.hWnd, GWL_STYLE, style Or ES_NUMBER
    Timer1.Interval = 1000
    Timer1.Enabled = False
    With DTPicker1
        .Format = dtpCustom
        'use this format as is - upper/lower case is important (HH = 24 hours; hh = 12 hours)
        .CustomFormat = "HH:mm:ss"
        .UpDown = True    '<<< set this so calendar will not show
        .Value = 0
    End With

    With UpDown1
        .BuddyControl = tboxdays
        .BuddyProperty = "Text"
    End With
    tboxdays.Text = 0

End Sub

Private Sub label_color_to_app_color_menue_Click()
    Me.BackColor = Label1.BackColor
    Me.ForeColor = Label1.ForeColor
    Me.Tage.ForeColor = Label1.ForeColor
    Me.Tage.BackColor = Label1.BackColor
    
    Dim col As Color
    Set col = New Color

    Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\application", "forecolor", col.OleColorToRgb(dlgColor.Color), Registry.HKEY_CURRENT_USER
    Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\application", "backcolor", col.OleColorToRgb(dlgColor.Color), Registry.HKEY_CURRENT_USER


End Sub




Private Sub Reset_all_settings_Click()
init_registry RESET_ALL
init_settings
End Sub

'Private Const RESET_NO As Integer = 0
'Private Const RESET_ALL As Integer = 1
'Private Const RESET_LBL_FONT As Integer = 2
'Private Const RESET_LBL_COLOR As Integer = 3
'Private Const RESET_APP_COLOR As Integer = 4

Private Sub Reset_Appcol_Click()
init_registry RESET_APP_COLOR
init_settings
End Sub

Private Sub Reset_Font_Click()
init_registry RESET_LBL_FONT
init_settings
End Sub

Private Sub Reset_lblcol_Click()
init_registry RESET_LBL_COLOR
init_settings
End Sub

Private Sub show_days_menu_Click()
    If show_days_menu.Checked = True Then
        show_days = False
        show_days_menu.Checked = False

    Else
        show_days = True
        show_days_menu.Checked = True


    End If
    set_zero_time

End Sub

Private Sub init_settings()
    On Error GoTo ErrorHandler
    Dim col As Color
    Set col = New Color
    'checks registry key and values if key not created it creates the key
    Dim res As Long
    Dim rgbarr() As String
    Dim valtosplit As String

    Label1.FontName = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "name", Registry.HKEY_CURRENT_USER)
    Label1.FontBold = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "bold", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
    Label1.FontItalic = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "italic", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
    Label1.FontSize = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "size", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
    Label1.FontStrikethru = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "strikethru", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
    Label1.FontUnderline = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "underline", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)


    col.RGBStr Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "color", Registry.HKEY_CURRENT_USER)

    With Label1
        .FontName = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "name", Registry.HKEY_CURRENT_USER)
        .FontBold = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "bold", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
        .FontItalic = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "italic", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
        .FontSize = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "size", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
        .FontStrikethru = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "strikethru", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
        .FontUnderline = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "underline", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
        .ForeColor = RGB(col.get_r, col.get_g, col.get_b)
    End With



    col.RGBStr Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl\lblbg", "color", Registry.HKEY_CURRENT_USER)

    Label1.BackColor = RGB(col.get_r, col.get_g, col.get_b)




    col.RGBStr Registry.QueryValue("SOFTWARE\" + App.ProductName + "\application", "backcolor", Registry.HKEY_CURRENT_USER)

    Me.BackColor = RGB(col.get_r, col.get_g, col.get_b)
    Me.btnPause.BackColor = Me.BackColor
    Me.btnStart.BackColor = Me.BackColor
    Me.Tage.BackColor = Me.BackColor
    col.RGBStr Registry.QueryValue("SOFTWARE\" + App.ProductName + "\application", "forecolor", Registry.HKEY_CURRENT_USER)
    Me.ForeColor = RGB(col.get_r, col.get_g, col.get_b)
    Me.Tage.ForeColor = Me.ForeColor
    
    show_days = Registry.QueryValue("SOFTWARE\" + App.ProductName + "\cdownlbl", "show_days", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD)
    
    Exit Sub

ErrorHandler:
    MsgBox "Cannot Read all settings. Reset app to Default Values!!!"
    init_registry RESET_ALL



End Sub


Private Sub init_registry(Optional reset As Integer = RESET_NO)
    Dim col As Color
    Set col = New Color
    'checks registry key and values if key not created it creates the key
    Dim res As Long


    'Private Const RESET_NO As Integer = 0
    'Private Const RESET_ALL As Integer = 1
    'Private Const RESET_LBL_FONT As Integer = 2
    'Private Const RESET_LBL_COLOR As Integer = 3
    'Private Const RESET_APP_COLOR As Integer = 4

    If Registry.RegOpenKeyEx(Registry.HKEY_CURRENT_USER, "SOFTWARE\" + App.ProductName, 0, Registry.KEY_ALL_ACCESS, res) _
       <> Registry.ERROR_NONE Or reset = RESET_ALL Or reset = RESET_LBL_FONT Or reset = RESET_LBL_COLOR Or reset = RESET_APP_COLOR Then

        Registry.CreateNewKey "SOFTWARE\" + App.ProductName, Registry.HKEY_CURRENT_USER

        If Registry.RegOpenKeyEx(Registry.HKEY_CURRENT_USER, "SOFTWARE\cdownlbl", 0, Registry.KEY_ALL_ACCESS, res) <> Registry.ERROR_NONE Or reset = RESET_ALL Or reset = RESET_LBL_FONT Or reset = RESET_LBL_COLOR Then
            Registry.CreateNewKey "SOFTWARE\" + App.ProductName + "\cdownlbl", Registry.HKEY_CURRENT_USER
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl", "show_days", "1", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
            If Registry.RegOpenKeyEx(Registry.HKEY_CURRENT_USER, "SOFTWARE\cdownlbl\lblfont", 0, Registry.KEY_ALL_ACCESS, res) <> Registry.ERROR_NONE Or reset = RESET_ALL Or reset = RESET_LBL_FONT Then
                Registry.CreateNewKey "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", Registry.HKEY_CURRENT_USER
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "name", "Arial", Registry.HKEY_CURRENT_USER
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "bold", "0", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "italic", "0", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "size", "50", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "strikethru", "0", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "underline", "0", Registry.HKEY_CURRENT_USER, Registry.REG_DWORD
    
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblfont", "color", Registry.QueryValue("Control Panel\Colors", "WindowText", Registry.HKEY_CURRENT_USER), Registry.HKEY_CURRENT_USER
            End If

            If Registry.RegOpenKeyEx(Registry.HKEY_CURRENT_USER, "SOFTWARE\cdownlbl\lblbg", 0, Registry.KEY_ALL_ACCESS, res) <> Registry.ERROR_NONE Or reset = RESET_ALL Or reset = RESET_LBL_COLOR Then
                Registry.CreateNewKey "SOFTWARE\" + App.ProductName + "\cdownlbl\lblbg", Registry.HKEY_CURRENT_USER
                Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\cdownlbl\lblbg", "color", Registry.QueryValue("Control Panel\Colors", "Menu", Registry.HKEY_CURRENT_USER), Registry.HKEY_CURRENT_USER
            End If



        End If

        If Registry.RegOpenKeyEx(Registry.HKEY_CURRENT_USER, "SOFTWARE\application", 0, Registry.KEY_ALL_ACCESS, res) <> Registry.ERROR_NONE Or reset = RESET_ALL Or reset = RESET_APP_COLOR Then
            Registry.CreateNewKey "SOFTWARE\" + App.ProductName + "\application", Registry.HKEY_CURRENT_USER
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\application", "backcolor", Registry.QueryValue("Control Panel\Colors", "Menu", Registry.HKEY_CURRENT_USER), Registry.HKEY_CURRENT_USER
            Registry.SetKeyValue "SOFTWARE\" + App.ProductName + "\application", "forecolor", Registry.QueryValue("Control Panel\Colors", "WindowText", Registry.HKEY_CURRENT_USER), Registry.HKEY_CURRENT_USER

        End If

    End If

End Sub

Private Sub Timer1_Timer()
    Dim time_now As Date
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long
    Dim days As Long

    time_now = Now
    If time_now >= time Then
        Timer1.Enabled = False
        set_zero_time
        Me.btnStart.Caption = "Start"
        myMsgBox "Time is up!!!", vbInformation + vbOKOnly, "Time is up", Me.hWnd




    Else

        seconds = DateDiff("s", time_now, time)
        minutes = seconds \ 60
        seconds = seconds - minutes * 60
        hours = minutes \ 60
        minutes = minutes - hours * 60

        If show_days = True Then
            days = hours \ 24
            hours = hours - days * 24
            Label1.Caption = CStr(days) + " Tage " + _
                             Format$(hours, "00") & ":" & _
                             Format$(minutes, "00") & ":" & _
                             Format$(seconds, "00")

        Else

            Label1.Caption = _
            Format$(hours, "00") & ":" & _
                             Format$(minutes, "00") & ":" & _
                             Format$(seconds, "00")
        End If






    End If






End Sub



