VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Form1 
   Caption         =   "Formular1"
   ClientHeight    =   3915
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   204079105
      CurrentDate     =   43622
   End
   Begin VB.CommandButton btnPause 
      Caption         =   "Pause"
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   11400
      Top             =   240
   End
   Begin VB.Label Tage 
      Caption         =   "Tage"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
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
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hours As Long
Dim minutes As Long
Dim seconds As Long
Dim days As Long
Dim time As Date

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER As Long = &H2000& 'or 8192 in decimal form

Private Sub btnStart_Click()

If btnStart.Caption = "Start" Then
btnStart.Caption = "End"
Timer1.Enabled = True
time = Now
time = DateAdd("h", Hour(DTPicker1.Value), time)
time = DateAdd("m", Minute(DTPicker1.Value), time)
time = DateAdd("s", Second(DTPicker1.Value), time)

Else
btnStart.Caption = "Start"
Timer1.Enabled = False
time = Now
Label1.Caption = Format(time, "d Tage HH:mm:ss")
End If



End Sub

Private Sub Form_Load()
Dim style As Long
    Text1.Text = ""
    style = GetWindowLong(Text1.hwnd, GWL_STYLE)
    SetWindowLong Text1.hwnd, GWL_STYLE, style Or ES_NUMBER
Timer1.Interval = 1000
Timer1.Enabled = False
With DTPicker1
        .Format = dtpCustom
        'use this format as is - upper/lower case is important (HH = 24 hours; hh = 12 hours)
        .CustomFormat = "HH:mm:ss"
        .UpDown = True '<<< set this so calendar will not show
        .Value = 0
End With

With UpDown1
.BuddyControl = Text1
.BuddyProperty = "Text"
End With
Text1.Text = 0

End Sub

Private Sub Timer1_Timer()
        Dim time_now As Date
        
        
        time_now = Now
        
        
        
        If time_now >= time Then
        Label1.Caption = "0 Tage 00:00:00"
        Timer1.Enabled = False
        
        Else
    
        seconds = DateDiff("s", time_now, time)
        minutes = seconds \ 60
        seconds = seconds - minutes * 60
        hours = minutes \ 60
        minutes = minutes - hours * 60
       
    
        Label1.Caption = _
            Format$(days, "0 Tage") & " " & _
            Format$(hours, "00") & ":" & _
            Format$(minutes, "00") & ":" & _
            Format$(seconds, "00")
    
        End If
    
    
    
    
    Label1.Caption = Format(time, "d Tage HH:mm:ss")
End Sub
