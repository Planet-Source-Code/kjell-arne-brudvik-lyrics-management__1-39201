VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lyrics Management - Options"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkScope 
      Caption         =   "Use Oscilloscope (slows down some PCs)"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CheckBox chkJukeboxOnTop 
      Caption         =   "Keep Jukebox Always OnTop"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CheckBox chkLastDir 
      Caption         =   "Remember last dir in Add dialog"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox chkClearSearch 
      Caption         =   "Clear search box after search"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Line Line6 
         X1              =   0
         X2              =   8760
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   8760
         Y1              =   710
         Y2              =   710
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":0CCA
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
         TabIndex        =   2
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Options - configure Lyrics Management as you wish"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5640
      Y1              =   3010
      Y2              =   3010
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5640
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkClearSearch_Click()

    '// Save setting
    Call SaveSetting("LyricsManagement", "Settings", "SearchClear", chkClearSearch.Value)

End Sub

Private Sub chkJukeboxOnTop_Click()

    '// Save setting
    Call SaveSetting("LyricsManagement", "Settings", "JukeboxOnTop", chkJukeboxOnTop.Value)

End Sub

Private Sub chkLastDir_Click()

    '// Save setting
    Call SaveSetting("LyricsManagement", "Settings", "LastDir", chkLastDir.Value)

End Sub

Private Sub chkScope_Click()

    '// Save setting
    Call SaveSetting("LyricsManagement", "Settings", "Oscilloscope", chkScope.Value)

End Sub

Private Sub Command1_Click()

    '// Close screen
    Unload Me

End Sub

Private Sub Form_Load()

    '// Get settings
    chkClearSearch.Value = GetSetting("LyricsManagement", "Settings", "SearchClear", "0")
    chkLastDir.Value = GetSetting("LyricsManagement", "Settings", "LastDir", "0")
    chkJukeboxOnTop.Value = GetSetting("LyricsManagement", "Settings", "JukeboxOnTop", "0")
    chkScope.Value = GetSetting("LyricsManagement", "Settings", "Oscilloscope", "0")

End Sub
