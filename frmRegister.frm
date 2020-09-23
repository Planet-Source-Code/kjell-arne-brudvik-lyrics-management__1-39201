VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lyrics Management - Register"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Unregister"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   24
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Register"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   4680
      Width           =   4095
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Index           =   1
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   3
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Index           =   2
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   4
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Index           =   3
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   5
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Index           =   4
      Left            =   5040
      MaxLength       =   5
      TabIndex        =   6
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   720
      X2              =   5760
      Y1              =   5535
      Y2              =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   720
      X2              =   5760
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   23
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   22
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   21
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   20
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial:"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Enter your registration:"
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
      TabIndex        =   17
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   720
      X2              =   5760
      Y1              =   4215
      Y2              =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   720
      X2              =   5760
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label9 
      Caption         =   "The price for this product is; $20 - this will allow for lifetime upgrades."
      Height          =   495
      Left            =   720
      TabIndex        =   16
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label Label8 
      Caption         =   $"frmRegister.frx":0CCA
      Height          =   615
      Left            =   720
      TabIndex        =   15
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label Label7 
      Caption         =   "- No nag which pops up from time to time."
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label Label6 
      Caption         =   "- Unlimited songs that can be stored."
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "- Unlimited search possibillities for online search."
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Register, and get access to this:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   $"frmRegister.frx":0D74
      Height          =   855
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   840
      Width           =   5055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   720
      X2              =   5760
      Y1              =   730
      Y2              =   730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   720
      X2              =   5760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Registration procedure"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   9
      Top             =   360
      Width           =   5055
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
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmRegister.frx":0E48
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Declares
    Dim tmpKey As String

    '// Check if key is valid
    If ValidateKey(txtKey(0).Text, txtKey(1).Text, txtKey(2).Text, txtKey(3).Text, txtKey(4).Text) = True Then
        MsgBox "Thank you for registering this software, your help is much appreaciated!", vbInformation, "Thanks for registering!"
        tmpKey = txtKey(0).Text & "-" & txtKey(1).Text & "-" & txtKey(2).Text & "-" & txtKey(3).Text & "-" & txtKey(4).Text
        Call SaveSetting("LyricsManagement", "Settings", "RegName", txtName.Text)
        Call SaveSetting("LyricsManagement", "Settings", "RegKey", tmpKey)
        Call SaveSetting("LyricsManagement", "Settings", "Registered", "1")
        frmMain.Caption = "Lyrics Management"
        frmMain.tmrUnregistered.Enabled = False
        Unload Me
    Else
        MsgBox "Invalid Registration Key - please try again!", vbCritical, "Registration is invalid!"
        Call SaveSetting("LyricsManagement", "Settings", "Registered", "0")
    End If

End Sub

Private Sub Command2_Click()

    '// Unload me
    Unload Me

End Sub

Private Sub Command3_Click()

    '// Ask if user really wants to unregister his program.
    Dim Answer
    Answer = MsgBox("Are you REALLY sure you want to unregister this copy?", vbCritical + vbYesNo, "Unregister?")
    
    If Answer = vbYes Then
        Call SaveSetting("LyricsManagement", "Settings", "Registered", "0")
        Command3.Enabled = False
        
        Command1.Enabled = True
        txtName.Text = "UNREGISTERED"
        txtKey(0).Text = ""
        txtKey(1).Text = ""
        txtKey(2).Text = ""
        txtKey(3).Text = ""
        txtKey(4).Text = ""
        txtName.Enabled = True
        txtKey(0).Enabled = True
        txtKey(1).Enabled = True
        txtKey(2).Enabled = True
        txtKey(3).Enabled = True
        txtKey(4).Enabled = True
        frmMain.tmrUnregistered.Enabled = True
        frmMain.Caption = frmMain.Caption & " - UNREGISTERED"
        
    End If

End Sub

Private Sub Form_Load()

    '// Declares
    Dim tmpKey As String
    Dim I As Integer
    Dim tmpString() As String

    '// Get registration info.
    If GetSetting("LyricsManagement", "Settings", "Registered", "0") = "0" Then
        Command1.Enabled = True
        txtName.Text = "UNREGISTERED"
        txtKey(0).Text = ""
        txtKey(1).Text = ""
        txtKey(2).Text = ""
        txtKey(3).Text = ""
        txtKey(4).Text = ""
    Else
        Command1.Enabled = False
        txtName.Enabled = False
        txtKey(0).Enabled = False
        txtKey(1).Enabled = False
        txtKey(2).Enabled = False
        txtKey(3).Enabled = False
        txtKey(4).Enabled = False
        Command3.Enabled = True
        txtName.Text = GetSetting("LyricsManagement", "Settings", "RegName", "")
        tmpKey = GetSetting("LyricsManagement", "Settings", "RegKey", "")
        
        tmpString = Split(tmpKey, "-")
        
        '// Get the key separated.
        For I = 0 To UBound(tmpString)
            txtKey(I) = tmpString(I)
        Next I
    End If

End Sub

Private Sub txtKey_Change(Index As Integer)

    '// Highlight
    txtKey(Index).SelStart = 0
    txtKey(Index).SelLength = Len(txtKey(Index).Text)

End Sub
