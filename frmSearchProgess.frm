VERSION 5.00
Begin VB.Form frmSearchProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Searching..."
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchProgess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimate 
      Interval        =   1000
      Left            =   120
      Top             =   960
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearchProgess.frx":08CA
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblStatus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Searching.. please wait.."
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
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearchProgess.frx":1194
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmSearchProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrAnimate_Timer()

    '// Show/hide image
    If Image2.Visible = False Then
        Image2.Visible = True
    Else
        Image2.Visible = False
    End If

End Sub
