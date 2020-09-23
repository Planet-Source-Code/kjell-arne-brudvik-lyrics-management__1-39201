VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAddSong 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lyrics Management - Add/edit song"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddSong.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "NOTE:"
      Height          =   615
      Left            =   720
      TabIndex        =   24
      Top             =   6720
      Width           =   6735
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "You need to hit refresh (F5) in order to make your added songs appear."
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   6495
      End
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Song text:"
      Height          =   2535
      Left            =   720
      TabIndex        =   20
      Top             =   4080
      Width           =   6735
      Begin VB.CommandButton Command4 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5280
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label Label5 
         Caption         =   "NOTE: Use the 'Search' function to the right, if you have internet, and want to get the text from there."
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Song essentials:"
      Height          =   3135
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   6735
      Begin VB.ListBox lstGenre 
         Height          =   255
         ItemData        =   "frmAddSong.frx":0CCA
         Left            =   6240
         List            =   "frmAddSong.frx":0E8A
         TabIndex        =   32
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtGenreSearch 
         Height          =   285
         Left            =   4800
         TabIndex        =   31
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkBackup 
         Caption         =   "Take a backup of the MP3 file to program data folder"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   2760
         Width           =   5055
      End
      Begin VB.ComboBox cboGenre 
         Height          =   315
         ItemData        =   "frmAddSong.frx":145A
         Left            =   1560
         List            =   "frmAddSong.frx":161A
         TabIndex        =   19
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtLength 
         Height          =   285
         Left            =   4200
         TabIndex        =   18
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtAlbum 
         Height          =   285
         Left            =   4200
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   1320
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Genre search:"
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblFiletitle 
         Caption         =   "-"
         Height          =   255
         Left            =   6000
         TabIndex        =   27
         Top             =   2400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Album:"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   $"frmAddSong.frx":1BEA
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7485
      TabIndex        =   0
      Top             =   0
      Width           =   7485
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
         Picture         =   "frmAddSong.frx":1CD0
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
         Caption         =   "Add/edit your songs here"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmAddSong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2

Private Sub Command1_Click()

    On Error Resume Next

    Dim tmpDirectory As String

    '// Open a file.
    cDialog.DialogTitle = "Add a song (MP3 file)"
    cDialog.Filter = "MP3 files (*.MP3)|*.mp3|All files (*.*)|*.*"
    
    If GetSetting("LyricsManagement", "Settings", "LastDir", "0") = "0" Then
        cDialog.InitDir = App.Path
    Else
        If GetSetting("LyricsManagement", "Settings", "LastDirectory", "") = "" Then
            cDialog.InitDir = App.Path
        Else
            If Dir(GetSetting("LyricsManagement", "Settings", "LastDirectory", ""), vbDirectory) = "" Then
                cDialog.InitDir = App.Path
            Else
                cDialog.InitDir = GetSetting("LyricsManagement", "Settings", "LastDirectory", "")
            End If
        End If
    End If
    
    cDialog.ShowOpen
    
    txtFilename.Text = cDialog.FileName
    
    '// Get directory
    tmpDirectory = GetFilePath(cDialog.FileName)
    Call SaveSetting("LyricsManagement", "Settings", "LastDirectory", tmpDirectory)
    
    '// Cancel if no file.
    If txtFilename.Text = "" Then Exit Sub
    
    '// Create an instance of the MP3 Class
    Dim ObjMP3 As clsMP3
    Set ObjMP3 = New clsMP3
    
    '// Read the MP3 data
    ObjMP3.ReadMP3 cDialog.FileName
    
    '// Set the values
    txtTitle.Text = ObjMP3.Songname
    txtArtist.Text = ObjMP3.Artist
    txtAlbum.Text = ObjMP3.Album
    cboYear.Text = ObjMP3.Year
    txtLength.Text = ObjMP3.Duration
    cboGenre.Text = ObjMP3.Genre
    lblFiletitle.Caption = cDialog.FileTitle
    
    chkBackup.Enabled = True

End Sub

Private Sub Command2_Click()

    On Error Resume Next

    '// Declares
    Dim tmpFilename As String
    Dim tmpFreeFile
    Dim tmpTitle As String
    Dim tmpArtist As String
    Dim tmpAlbum As String
    Dim tmpYear As String
    Dim tmpLength As String
    Dim tmpGenre As String
    Dim tmpText As String
    Dim tmpTextfile As String
    Dim tmpFile As String
    
    '// Check if all needed data is added.
    If txtTitle.Text = "" Then
        MsgBox "You need to enter a title!", vbCritical, "Error!"
        Exit Sub
    End If
    
    If txtArtist.Text = "" Then
        MsgBox "You need to enter an Artist!", vbCritical, "Error!"
        Exit Sub
    End If
    
    If GetSetting("LyricsManagement", "Settings", "Registered", "0") = "0" Then
        If tmpAddCount = 5 Then
            MsgBox "You have reached your maximum add count for this session '5', register the program to have unlimited add.", vbCritical, "Add limit reached!"
            Exit Sub
        Else
            tmpAddCount = tmpAddCount + 1
        End If
    End If
    
    tmpFreeFile = FreeFile
    tmpFilename = txtArtist.Text & "¤" & txtTitle.Text & ".lyr"
    tmpTitle = txtTitle.Text
    tmpArtist = txtArtist.Text
    tmpAlbum = txtAlbum.Text
    tmpYear = cboYear.Text
    tmpLength = txtLength.Text
    tmpGenre = cboGenre.Text
    tmpText = txtText.Text
    tmpFile = txtFilename.Text
    tmpTextfile = txtArtist.Text & "¤" & txtTitle.Text & ".sng"

    '// Check if the entry already exist.
    If Dir(App.Path & "\data\" & tmpFilename) = "" Then
        
WriteInfo:
        
        '// Check if we should backup the file.
        If chkBackup.Value = 1 Then
            If Dir(txtFilename.Text) = "" Then
            Else
                
                '// Copy file.
                Dim tmpSourceFile As String
                Dim tmpDestinationFile As String
                tmpSourceFile = txtFilename.Text
                tmpDestinationFile = App.Path & "\data\mp3\" & lblFiletitle.Caption
                FileCopy tmpSourceFile, tmpDestinationFile
                tmpFile = tmpDestinationFile
                
            End If
        End If
        
        '// Save the song information.
        Open App.Path & "\data\" & tmpFilename For Random As tmpFreeFile
            Put tmpFreeFile, 1, tmpFile
            Put tmpFreeFile, 2, tmpTitle
            Put tmpFreeFile, 3, tmpArtist
            Put tmpFreeFile, 4, tmpAlbum
            Put tmpFreeFile, 5, tmpYear
            Put tmpFreeFile, 6, tmpLength
            Put tmpFreeFile, 7, tmpGenre
            Put tmpFreeFile, 8, tmpTextfile
        Close tmpFreeFile
        
        '// Save text
        If Dir(App.Path & "\data\" & tmpTextfile) = "" Then
        Else
            Kill App.Path & "\data\" & tmpTextfile
        End If
        
        Open App.Path & "\data\" & tmpTextfile For Append As tmpFreeFile
            Print #tmpFreeFile, tmpText
        Close tmpFreeFile
        
        '// Update the listing.
        frmMain.Show
        Unload Me
        Exit Sub
    Else
        Dim Answer
        Answer = MsgBox("That song already exist in the database, update?", vbCritical + vbYesNo, "Song exists!")
        
        If Answer = vbYes Then
            GoTo WriteInfo
        Else
        End If
        
        Exit Sub
    End If

End Sub

Private Sub Command3_Click()

    '// Cancel the operation.
    frmMain.Show
    Unload Me

End Sub

Private Sub Command4_Click()

    '// Show search
    frmSearch.txtSearch.Text = txtTitle.Text
    frmSearch.tmrSearch.Enabled = True
    frmSearch.Command4.Enabled = True
    frmSearch.Command3.Enabled = False
    frmSearch.Show

End Sub

Private Sub Form_Load()

    '// Load year data
    Dim I As Long
    
    For I = 1930 To 2010 Step 1
        cboYear.AddItem I
    Next I
    cboYear.Text = "2002"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '// Cancel the operation.
    frmMain.Show
    Unload Me

End Sub

Private Sub txtGenreSearch_Change()

    '// Auto find items.
    lstGenre.ListIndex = SendMessage(lstGenre.hWnd, LB_FINDSTRING, -1, ByVal (txtGenreSearch.Text))
    cboGenre.ListIndex = lstGenre.ListIndex

End Sub
