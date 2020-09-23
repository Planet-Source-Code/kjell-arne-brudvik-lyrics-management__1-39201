VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1770
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Timer tmrStart 
         Interval        =   2000
         Left            =   4920
         Top             =   120
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   5535
         TabIndex        =   1
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Email: brudvik@online.no"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) 2002 - Kjell Arne Brudvik"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lyrics Management"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   120
         Picture         =   "frmSplash.frx":0CCA
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1200
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Loading..."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5175
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    '// Set version
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & " - Revision: " & App.Revision

End Sub

Private Sub tmrStart_Timer()

    '// Move on
    Unload Me
    frmMain.Show

End Sub
