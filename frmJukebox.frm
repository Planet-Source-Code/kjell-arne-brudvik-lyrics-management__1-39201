VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmJukebox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lyrics Management - Jukebox"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJukebox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton StartButton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Start"
      Height          =   336
      Left            =   1080
      TabIndex        =   52
      Top             =   7680
      Width           =   984
   End
   Begin VB.CommandButton StopButton 
      BackColor       =   &H0000C000&
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   336
      Left            =   120
      TabIndex        =   51
      Top             =   7680
      Width           =   984
   End
   Begin VB.PictureBox Picture10 
      Height          =   255
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   48
      Top             =   2640
      Width           =   1215
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Hz"
         Height          =   255
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Left            =   6000
      Picture         =   "frmJukebox.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Send to tray"
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox chkOnTop 
      Height          =   375
      Left            =   1080
      Picture         =   "frmJukebox.frx":091F
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Always OnTop"
      Top             =   3360
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox chkList 
      Height          =   375
      Left            =   600
      Picture         =   "frmJukebox.frx":0993
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Song text"
      Top             =   3360
      Value           =   1  'Checked
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   6480
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkMute 
      Height          =   375
      Left            =   120
      Picture         =   "frmJukebox.frx":09FD
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Mute sound"
      Top             =   3360
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.Timer tmrVolume 
      Interval        =   1
      Left            =   6480
      Top             =   2280
   End
   Begin VB.PictureBox Picture9 
      Height          =   255
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   34
      Top             =   1800
      Width           =   1215
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Balance"
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture8 
      Height          =   255
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   32
      Top             =   840
      Width           =   1215
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command8 
      Height          =   375
      Left            =   6000
      Picture         =   "frmJukebox.frx":0A5F
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   5520
      Picture         =   "frmJukebox.frx":0AB4
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   3600
      Picture         =   "frmJukebox.frx":0B11
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   6195
      TabIndex        =   15
      Top             =   4080
      Width           =   6255
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   6255
         TabIndex        =   17
         Top             =   2400
         Width           =   6255
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   375
            Left            =   840
            TabIndex        =   38
            Top             =   120
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   661
            _Version        =   393216
            Min             =   1
            Max             =   1000
            SelStart        =   100
            TickFrequency   =   1000
            Value           =   100
         End
         Begin VB.CommandButton Command10 
            Height          =   375
            Left            =   5760
            Picture         =   "frmJukebox.frx":0B6F
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command9 
            Height          =   375
            Left            =   4320
            Picture         =   "frmJukebox.frx":0BCC
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command5 
            Height          =   375
            Left            =   5280
            Picture         =   "frmJukebox.frx":0C2A
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command4 
            Height          =   375
            Left            =   4800
            Picture         =   "frmJukebox.frx":0C83
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Speed:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   180
            Width           =   735
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00404040&
            X1              =   0
            X2              =   6240
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Label lblSong 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   60
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6195
      TabIndex        =   14
      Top             =   3840
      Width           =   6255
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Song text"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   6255
      End
   End
   Begin MSComctlLib.Slider sldBalance 
      Height          =   630
      Left            =   5055
      TabIndex        =   13
      Top             =   2040
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1111
      _Version        =   393216
      Min             =   -5000
      Max             =   5000
      TickStyle       =   2
      TickFrequency   =   1000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   4875
      TabIndex        =   8
      Top             =   1080
      Width           =   4935
      Begin VB.PictureBox Scope 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   690
         Left            =   0
         ScaleHeight     =   57.621
         ScaleMode       =   0  'User
         ScaleWidth      =   380.781
         TabIndex        =   44
         Top             =   840
         Width           =   4875
         Begin VB.PictureBox HyPiano 
            Height          =   372
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   915
            TabIndex        =   47
            Top             =   120
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.PictureBox ScopeBuff 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C000C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000002&
            Height          =   336
            Left            =   3240
            ScaleHeight     =   22
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   78
            TabIndex        =   46
            Top             =   240
            Visible         =   0   'False
            Width           =   1176
         End
         Begin VB.ComboBox DevicesBox 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   0
            Visible         =   0   'False
            Width           =   3108
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   4935
         TabIndex        =   11
         Top             =   1560
         Width           =   4935
         Begin MSComctlLib.Slider sldProgress 
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   450
            _Version        =   393216
            TickFrequency   =   100
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00404040&
            X1              =   0
            X2              =   4920
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Label lblFreq 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fq:"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblBalance 
         BackStyle       =   0  'Transparent
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   135
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4875
      TabIndex        =   7
      Top             =   840
      Width           =   4935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jukebox"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6495
      TabIndex        =   4
      Top             =   0
      Width           =   6495
      Begin VB.Line Line6 
         X1              =   0
         X2              =   9840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   9840
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmJukebox.frx":0CD9
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Lyrics Management"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Jukebox - play your songs here"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   5040
      Picture         =   "frmJukebox.frx":19A3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   4560
      Picture         =   "frmJukebox.frx":19F4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   4080
      Picture         =   "frmJukebox.frx":1A4D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   375
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   6480
      Top             =   1320
   End
   Begin VB.Timer tmrSong 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   1800
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   630
      Left            =   5055
      TabIndex        =   27
      Top             =   1080
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1111
      _Version        =   393216
      Max             =   2500
      TickStyle       =   2
      TickFrequency   =   250
   End
   Begin VB.Label lblHz 
      Alignment       =   2  'Center
      Caption         =   "0 hz"
      Height          =   255
      Left            =   5160
      TabIndex        =   50
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6360
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   6360
      Y1              =   3240
      Y2              =   3240
   End
   Begin MediaPlayerCtl.MediaPlayer Player 
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuCurrentSong 
         Caption         =   "None"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "&Show Jukebox..."
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play..."
         Enabled         =   0   'False
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuPause 
         Caption         =   "P&ause..."
         Enabled         =   0   'False
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop..."
         Enabled         =   0   'False
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open file..."
         Enabled         =   0   'False
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuMute 
         Caption         =   "&Mute sound..."
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Jukebox"
         Shortcut        =   +{F12}
      End
   End
End
Attribute VB_Name = "frmJukebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DevHandle As Long
Private Visualizing As Boolean
Private Divisor As Long
Private ScopeHeight As Long
                           
Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer
    ProductID As Integer
    DriverVersion As Long
    ProductName(1 To 32) As Byte
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&
Private Const WAVE_FORMAT_1M08 = &H1&
Private Const WAVE_FORMAT_1S08 = &H2&
Private Const WAVE_FORMAT_1M16 = &H4&
Private Const WAVE_FORMAT_1S16 = &H8&
Private Const WAVE_FORMAT_2M08 = &H10&
Private Const WAVE_FORMAT_2S08 = &H20&
Private Const WAVE_FORMAT_2M16 = &H40&
Private Const WAVE_FORMAT_2S16 = &H80&
Private Const WAVE_FORMAT_4M08 = &H100&
Private Const WAVE_FORMAT_4S08 = &H200&
Private Const WAVE_FORMAT_4M16 = &H400&
Private Const WAVE_FORMAT_4S16 = &H800&
Private Const WAVE_FORMAT_PCM = 1
Private Const WHDR_DONE = &H1&
Private Const WHDR_PREPARED = &H2&
Private Const WHDR_BEGINLOOP = &H4&
Private Const WHDR_ENDLOOP = &H8&
Private Const WHDR_INQUEUE = &H10&
Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long
Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Dim maxvol As Long, Hz As Long, oscila As Long
Dim HzColor As Long, xMax As Integer, HzTip As Long

Sub InitDevices()

    '// Fill the DevicesBox box with all the compatible audio input devices
    '// Bail if there are none.
    
    '// Declares
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        If Caps.Formats And WAVE_FORMAT_4M16 Then '16-bit mono devices
            Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Error!"
        End
    End If
    DevicesBox.ListIndex = 0
    
End Sub

Private Sub DoStop()

    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    DevicesBox.Enabled = True
    
End Sub

Private Sub Visualize()

    '// Declares
    Static X As Long
    Static Wave As WaveHdr
    Static InData(0 To NumSamples - 1) As Integer
    Static OutData(0 To NumSamples - 1) As Single
    
        Do
            Wave.lpData = VarPtr(InData(0))
            Wave.dwBufferLength = NumSamples
            Wave.dwFlags = 0
            Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
            Do
    
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
            If DevHandle = 0 Then Exit Do
            
            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call FFTAudio(InData, OutData)

            ScopeBuff.Picture = LoadPicture()
            ScopeBuff.BackColor = vbBlack
            oscila = vbGreen
            
            'If mpiano.Checked = True
            
            Dim c As Double, LowMidHig
            
            For X = 1 To 511
                ScopeBuff.DrawWidth = 1
                ScopeBuff.PSet (X, ScopeHeight / 3 - (InData(X) / 500)), oscila ' oscilloscope
               
                If Abs(OutData(X)) > maxvol Then
                    maxvol = Abs(OutData(X))
                    Hz = Int(44100 * X) / 1024
                    lblHz.Caption = Hz & " Hz"
                    HzColor = vbRed '+ Hz
                    LowMidHig = ScopeHeight

                    If X < 11 Then Hz = Hz - 22
                    If X > 11 Then Hz = Hz / 10: HzColor = StopButton.BackColor + X: LowMidHig = (ScopeHeight / 3) * 2
                    If X > 119 Then Hz = X - 119: HzColor = vbYellow: LowMidHig = (ScopeHeight / 3): OutData(X) = OutData(X) * 2
                    xMax = X
                End If
            Next
            
            X = xMax
            c = 0.5 * (1 - Cos(X * 2 * 3.1416 / 512))
            OutData(X) = c * OutData(X)

            HzColor = (Int(44100 * X) / 1024) * 10000
            ScopeBuff.DrawWidth = 5
            If X > 11 Then ScopeBuff.DrawWidth = 3
            If X > 119 Then ScopeBuff.DrawWidth = 4
            ScopeBuff.Line (Hz, LowMidHig)-(Hz, LowMidHig - ((Abs(OutData(X)) / 10))), HzColor
            ScopeBuff.DrawWidth = 1
            maxvol = 0
            Scope.Picture = ScopeBuff.Image
            ScopeBuff.Cls
            DoEvents
        Loop While DevHandle <> 0
        
End Sub

Private Sub chkList_Click()

    '// Show/hide song list.
    If chkList.Value = 1 Then
        Me.Height = 7545
    Else
        Me.Height = 3690
    End If

End Sub

Private Sub chkMute_Click()

    '// Either mute or unmute
    If chkMute.Value = 1 Then
        Player.Mute = False
    Else
        Player.Mute = True
    End If

End Sub

Private Sub chkOnTop_Click()

    '// Always on top
    If chkOnTop.Value = 1 Then
        Call AlwaysOnTop(frmJukebox, True)
    Else
        Call AlwaysOnTop(frmJukebox, False)
    End If

End Sub

Private Sub Command1_Click()

    On Error Resume Next

    '// Play song
    Player.Play
    sldProgress.Max = Player.Duration
    tmrSong.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
    Command8.Enabled = False
    
    If GetSetting("LyricsManagement", "Settings", "Oscilloscope", "0") = "0" Then
        Exit Sub
    End If
    
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = 44100
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    maxvol = waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    
    Call waveInStart(DevHandle)
    DevicesBox.Enabled = False
    Call Visualize

End Sub
Private Function SecondsToTime(lSeconds As Double) As String
    Dim sTime As String
    Dim iSeconds As Integer
    Dim iMinutes As Integer
    
    iSeconds = Abs(Fix(lSeconds)) Mod 60
    iMinutes = Fix(Abs(Fix(lSeconds)) / 60)
    
    sTime = iMinutes & ":" & IIf(iSeconds < 10, "0", "") & iSeconds
    
    SecondsToTime = sTime
End Function

Private Sub Command10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Move forward
    lblSong.Top = lblSong.Top - 100

End Sub

Private Sub Command11_Click()

    '// Add the program to the tray.
    Call AddToTray(Me.Icon, Me.Caption, Me)

End Sub

Private Sub Command2_Click()

    On Error Resume Next

    '// Pause song
    Player.Pause
    tmrSong.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = True
    Command8.Enabled = False

End Sub

Private Sub Command3_Click()

    On Error Resume Next

    '// Stop song
    Player.Stop
    Player.FastReverse
    Player.CurrentPosition = 0
    sldProgress.Value = 0
    Player.AutoRewind = True
    tmrSong.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
    Command8.Enabled = True

    If GetSetting("LyricsManagement", "Settings", "Oscilloscope", "0") = "0" Then
        Exit Sub
    End If

    Call DoStop

End Sub

Private Sub Command4_Click()

    '// Start timer
    tmrSong.Enabled = True

End Sub

Private Sub Command5_Click()

    '// Stop timer
    tmrSong.Enabled = False

End Sub

Private Sub Command8_Click()

    '// Open a file.
    cDialog.DialogTitle = "Add a song (MP3 file)"
    cDialog.Filter = "MP3 files (*.MP3)|*.mp3|All files (*.*)|*.*"

    If GetSetting("LyricsManagement", "Settings", "LastDir", "0") = "0" Then
        cDialog.InitDir = App.path
    Else
        If GetSetting("LyricsManagement", "Settings", "LastDirectory", "") = "" Then
            cDialog.InitDir = App.path
        Else
            If Dir(GetSetting("LyricsManagement", "Settings", "LastDirectory", ""), vbDirectory) = "" Then
                cDialog.InitDir = App.path
            Else
                cDialog.InitDir = GetSetting("LyricsManagement", "Settings", "LastDirectory", "")
            End If
        End If
    End If
    
    cDialog.ShowOpen
    
    '// Cancel if no file selected
    If cDialog.FileName = "" Then Exit Sub
    
    '// Load the file.
    Dim ObjMP3 As clsMP3
    Set ObjMP3 = New clsMP3
    
    '// Read the MP3 data
    ObjMP3.ReadMP3 cDialog.FileName
    
    '// Set the values
    lblSong.Caption = ""
    
    If ObjMP3.Artist = "" And ObjMP3.Songname = "" Then
        lblTitle.Caption = GetFileName(cDialog.FileName)
    Else
        lblTitle.Caption = ObjMP3.Artist & " - " & ObjMP3.Songname
    End If
    
    lblFreq.Caption = ObjMP3.Frequency

    Player.FileName = cDialog.FileName
    sldProgress.Max = Player.Duration
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
    Command8.Enabled = False
    
    lblSong.Top = 2280
    '// Check if we got that lyric, if so display it.
    If Dir(App.path & "\data\" & ObjMP3.Artist & "¤" & ObjMP3.Songname & ".sng") = "" Then
    Else
    
        '// Load song
        Dim Var1
        Open App.path & "\data\" & ObjMP3.Artist & "¤" & ObjMP3.Songname & ".sng" For Input As #1
            FileLength = LOF(1)
            Var1 = Input(FileLength, #1)
            lblSong.Caption = Var1
        Close #1
        tmrSong.Enabled = True
    
    End If

End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Move back
    lblSong.Top = lblSong.Top + 100

End Sub

Private Sub Form_Load()

    '// Declares
    Dim sStatus As String
    
    Select Case Player.PlayState
        Case mpClosed
            sStatus = "No file selected - "
        Case mpPaused
            sStatus = "Paused - "
        Case mpPlaying
            sStatus = "Playing - "
            'HCSSlider1.Value = MediaPlayer1.CurrentPosition
        Case mpStopped
            sStatus = "Stopped - "
    End Select
    
    If Player.PlayState <> mpClosed And Player.PlayState <> mpStopped Then
        lblTime.Caption = sStatus & SecondsToTime(Player.CurrentPosition) & "/" & SecondsToTime(Player.Duration)
    Else
        lblTime.Caption = sStatus & "0:00/0:00"
    End If

    '// See if we shall place it on top.
    If GetSetting("LyricsManagement", "Settings", "JukeboxOnTop", "0") = "1" Then
        Call AlwaysOnTop(frmJukebox, True)
        chkOnTop.Value = 1
    Else
        Call AlwaysOnTop(frmJukebox, False)
        chkOnTop.Value = 0
    End If
    
    Player.AutoStart = False
    Player.AutoRewind = True
    
    '// Set oscilloscope
    Call InitDevices
    Call DoReverse

    ScopeBuff.Width = Scope.ScaleWidth
    ScopeBuff.Height = Scope.ScaleHeight
    ScopeBuff.BackColor = Scope.BackColor
    ScopeHeight = Scope.Height

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Get tray icon action.
    If RespondToTray(X) = 1 Then
        Call ShowFormAgain(Me)
    ElseIf RespondToTray(X) = 2 Then
        PopupMenu mnuFile, , , , mnuShow
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '// Unload
    Player.Stop
    
    If DevHandle <> 0 Then
        Call DoStop
        Cancel = 1
        If Visualizing = True Then
            Unload Me
        End If
    End If
    
    Unload Me
    Call ShowFormAgain(frmMain)
    
End Sub


Private Sub lblSong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Stop timer
    tmrSong.Enabled = False

End Sub

Private Sub lblSong_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Move label
    If Button = 1 Then
        lblSong.Move 120, Y
    End If

End Sub

Private Sub lblSong_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Start timer
    tmrSong.Enabled = True

End Sub

Private Sub mnuClose_Click()

    '// Stop jukebox
    Player.Stop
    Unload Me

End Sub

Private Sub mnuMute_Click()

    '// Check if muted
    If chkMute.Value = 0 Then
        chkMute.Value = 1
        mnuMute.Checked = False
        Player.Mute = False
    Else
        chkMute.Value = 0
        mnuMute.Checked = True
        Player.Mute = True
    End If

End Sub

Private Sub mnuOpen_Click()

    '// Open
    Command8_Click

End Sub

Private Sub mnuPause_Click()

    '// Pause
    Command2_Click

End Sub

Private Sub mnuPlay_Click()

    '// Play
    Command1_Click

End Sub

Private Sub mnuShow_Click()

    '// Remove from tray
    Call ShowFormAgain(Me)

End Sub

Private Sub mnuStop_Click()

    '// Stop
    Command3_Click

End Sub

Private Sub Scope_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Y >= ScaleHeight / 3 And Y < ScaleHeight / 3 * 2 Then Scope.ToolTipText = "516.84 to 5,125.33Hz": Exit Sub
  If Y >= ScaleHeight / 3 * 2 Then Scope.ToolTipText = "43.07 to 473.77 Hz": Exit Sub
  Scope.ToolTipText = "5,168 to 22K Hz"

End Sub

Private Sub sldBalance_Click()

    On Error Resume Next

    '// Set balance
    If sldBalance.Value > -500 And sldBalance.Value < 500 Then
        lblBalance.Caption = "Center"
    ElseIf sldBalance.Value < -500 Then
        lblBalance.Caption = "Left"
    ElseIf sldBalance.Value > 500 Then
        lblBalance.Caption = "Right"
    End If
    
    Player.Balance = sldBalance.Value

End Sub

Private Sub sldProgress_Scroll()

    '// Set position
    Player.CurrentPosition = sldProgress.Value

End Sub

Private Sub sldSpeed_Click()

    '// Set speed
    tmrSong.Interval = sldSpeed.Value

End Sub

Private Sub sldVolume_Scroll()

    On Error Resume Next

    '// Declares
    Dim tmpVolume
    Dim tmpSetVolume
    Dim tmpMin As Integer
    Dim tmpMax As Integer

    lblVolume.ForeColor = RGB(0 + sldVolume.Value / 10, 0, 0)
    tmpSetVolume = sldVolume.Value - 2500
    Player.Volume = tmpSetVolume
    
    tmpMin = sldVolume.Min
    tmpMax = sldVolume.Value
    
    lblVolume.Caption = CDbl(tmpMax) \ CDbl(25) & " %"

End Sub

Private Sub tmrSong_Timer()

    '// Show song
    lblSong.Top = lblSong.Top - 10

End Sub

Private Sub tmrTime_Timer()

    '// Declares
    Dim sStatus As String
    
    Select Case Player.PlayState
        Case mpClosed
            sStatus = "No file selected - "
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            mnuPlay.Enabled = False
            mnuPause.Enabled = False
            mnuStop.Enabled = False
            mnuOpen.Enabled = True
        Case mpPaused
            mnuPlay.Enabled = True
            mnuPause.Enabled = False
            mnuStop.Enabled = True
            mnuOpen.Enabled = False
            sStatus = "Paused - "
        Case mpPlaying
            mnuPlay.Enabled = False
            mnuPause.Enabled = True
            mnuStop.Enabled = True
            mnuOpen.Enabled = False
            sStatus = "Playing - "
            sldProgress.Value = Player.CurrentPosition
        Case mpStopped
            Command1.Enabled = True
            Command2.Enabled = False
            Command3.Enabled = False
            Command8.Enabled = True
            mnuPlay.Enabled = True
            mnuPause.Enabled = False
            mnuStop.Enabled = False
            mnuOpen.Enabled = True
            sStatus = "Stopped - "
    End Select
    
    If Player.PlayState <> mpClosed And Player.PlayState <> mpStopped Then
        lblTime.Caption = sStatus & SecondsToTime(Player.CurrentPosition) & "/" & SecondsToTime(Player.Duration)
    Else
        lblTime.Caption = sStatus & "0:00/0:00"
    End If
    mnuCurrentSong.Caption = lblTitle.Caption
    
    '// Strip down title
    If Len(lblTitle.Caption) > 40 Then
        lblTitle.Caption = Mid(lblTitle.Caption, 1, 40) & "..."
    End If
    
End Sub

Private Sub tmrVolume_Timer()

    On Error Resume Next

    '// Declares
    Dim tmpVolume
    Dim tmpSetVolume
    Dim tmpMin As Integer
    Dim tmpMax As Integer

    lblVolume.ForeColor = RGB(0 + sldVolume.Value / 10, 0, 0)
    tmpSetVolume = sldVolume.Value - 2500
    Player.Volume = tmpSetVolume
    
    tmpMin = sldVolume.Min
    tmpMax = sldVolume.Value
    
    lblVolume.Caption = CDbl(tmpMax) \ CDbl(25) & " %"
    tmrVolume.Enabled = False

End Sub
