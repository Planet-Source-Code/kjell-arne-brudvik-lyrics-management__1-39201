VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About the program.."
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Register"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   720
      X2              =   4560
      Y1              =   2650
      Y2              =   2650
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   720
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label4 
      Caption         =   "brudvik@online.no"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright (c) 2002 - Kjell Arne Brudvik"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":0CCA
      Height          =   1095
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   720
      X2              =   4560
      Y1              =   730
      Y2              =   730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   720
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblVersion 
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0D87
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Unload
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Show register
    Unload Me
    frmRegister.Show

End Sub

Private Sub Form_Load()

    '// Set version
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & " - Revision: " & App.Revision

End Sub
