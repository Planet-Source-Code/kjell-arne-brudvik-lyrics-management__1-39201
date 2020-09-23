VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lyrics Management"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Height          =   375
      Left            =   9120
      Picture         =   "frmMain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Send to tray"
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status:"
      Height          =   1215
      Left            =   0
      TabIndex        =   31
      Top             =   5400
      Width           =   3495
      Begin VB.Label lblEntries 
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
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "NOTE: Right click a song in order to be able to edit/delete it."
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Entries:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search:"
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   6720
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   285
         Left            =   2760
         TabIndex        =   30
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label12 
         Caption         =   "Search for entries:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Timer tmrUnregistered 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2400
      Top             =   1320
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   1800
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9630
      TabIndex        =   20
      Top             =   0
      Width           =   9630
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "The ideal way to store lyrics on your home computer"
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   360
         Width           =   5895
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
         TabIndex        =   21
         Top             =   120
         Width           =   5655
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":0D1F
         Top             =   120
         Width           =   480
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   9840
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   9840
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Song information:"
      Height          =   2055
      Left            =   3720
      TabIndex        =   9
      Top             =   5640
      Width           =   5895
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   5160
         Picture         =   "frmMain.frx":19E9
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblFileshow 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label lblGenre 
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblLength 
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label lblYear 
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lblAlbum 
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label lblFilename 
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label lblArtist 
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label lblTitle 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Album:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   840
      Width           =   5895
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Song text"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   5895
      End
   End
   Begin VB.TextBox txtSong 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   840
      Width           =   3495
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Lyrics listing"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.FileListBox fileLyrics 
      Height          =   285
      Left            =   -240
      Pattern         =   "*.lyr"
      TabIndex        =   2
      Top             =   8160
      Width           =   495
   End
   Begin MSComctlLib.TreeView tvLyrics 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7435
      _Version        =   393217
      Indentation     =   353
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   -360
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C67
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   8760
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9600
      Y1              =   7815
      Y2              =   7815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   9600
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Label lblSongname 
      Caption         =   "-"
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
      TabIndex        =   8
      Top             =   5280
      Width           =   5895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3615
      X2              =   3615
      Y1              =   840
      Y2              =   7680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   3600
      X2              =   3600
      Y1              =   840
      Y2              =   7680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuShow 
         Caption         =   "&Show Lyrics Manager..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu mnuSongs 
      Caption         =   "&Songs"
      Begin VB.Menu mnuAddSong 
         Caption         =   "&Add Song..."
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuEditSong 
         Caption         =   "&Edit Song..."
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuDeleteSong 
         Caption         =   "&Delete Song..."
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddDirectory 
         Caption         =   "&Add directory..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuList7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextSearch 
         Caption         =   "&Search for text..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh list..."
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuJukebox 
         Caption         =   "&Jukebox..."
         Shortcut        =   ^{F5}
      End
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "&Online"
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search for lyrics..."
         Shortcut        =   ^{F6}
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Configuration"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Shortcut        =   ^{F7}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register..."
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About the program.."
         Shortcut        =   ^{F9}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Check if there are any filenames to load
    If lblFilename.Caption = "" Then
        MsgBox "There is no MP3 file associated with this song!", vbCritical, "Error!"
        Exit Sub
    End If

    '// Load jukebox
    'frmJukebox.Player.FileName = lblFilename.Caption
    'frmJukebox.Player.BaseURL = lblFilename.Caption
    frmJukebox.Player.FileName = lblFilename.Caption
    frmJukebox.lblTitle.Caption = lblArtist.Caption & " - " & lblTitle.Caption
    
    '// Load the file.
    Dim ObjMP3 As clsMP3
    Set ObjMP3 = New clsMP3
    
    '// Read the MP3 data
    ObjMP3.ReadMP3 lblFilename.Caption
    
    '// Set the values
    frmJukebox.lblFreq.Caption = ObjMP3.Frequency
    
    frmJukebox.lblSong.Caption = txtSong.Text
    frmJukebox.lblSong.Top = 1800
    frmJukebox.sldVolume.Value = 1250
    frmJukebox.Command1.Enabled = False
    frmJukebox.Command2.Enabled = True
    frmJukebox.Command3.Enabled = True
    frmJukebox.Command8.Enabled = False
    frmJukebox.Show
    frmJukebox.sldProgress.Max = frmJukebox.Player.Duration

End Sub

Private Sub Command11_Click()

    '// Add the program to the tray.
    Call AddToTray(Me.Icon, Me.Caption, Me)

End Sub

Private Sub Command2_Click()

    '// Check if there is anything to search for.
    If txtSearch.Text = "" Then Exit Sub

    '// Declares
    Dim I As Long
    Dim tmpFound As Long
    
    tmpFound = 0
    Frame2.Caption = "Search:"
    For I = 1 To tvLyrics.Nodes.Count Step 1
        Frame2.Caption = "Search: " & I & "/" & tvLyrics.Nodes.Count
        If InStr(1, tvLyrics.Nodes(I).Text, txtSearch.Text, vbTextCompare) = "0" Then
            tvLyrics.Nodes.Item(I).Expanded = False
        Else
            tvLyrics.Nodes(I).Selected = True
            tmpFound = tmpFound + 1
        End If
    Next I
    Frame2.Caption = "Search: Found - " & tmpFound & " entries."
    
    '// Check if we want to clear the search field.
    If GetSetting("LyricsManagement", "Settings", "SearchClear", "0") = "1" Then
        txtSearch.Text = ""
    Else
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next

    '// Declares
    Dim I As Long
    Dim tmpTitle As String
    Dim tmpGroup As String
    Dim tmpFilename As String
    Dim tmpString As String
    Dim tmpLetter As String
    Dim tmpSplit As Long
    
    tmpSearchCount = 0
    tmpAddCount = 0
    txtSong.Locked = True
    
    '// Make sure data directory exist.
    If Dir(App.path & "\data", vbDirectory) = "" Then MkDir App.path & "\data"
    If Dir(App.path & "\data\mp3", vbDirectory) = "" Then MkDir App.path & "\data\mp3"
    fileLyrics.Pattern = "*.lyr"
    fileLyrics.path = App.path & "\data"
    fileLyrics.Refresh

    '// Open the progressbar
    frmProgress.pBar.Min = 0
    frmProgress.pBar.Max = fileLyrics.ListCount - 1
    frmProgress.Show
    DoEvents

    '// Load up the startup data.
    tvLyrics.Nodes.Clear
    tvLyrics.Nodes.Add , , "Lyrics", "Lyrics", 1

    For I = 0 To fileLyrics.ListCount - 1 Step 1
        fileLyrics.ListIndex = I
        tmpFilename = Mid(fileLyrics.FileName, 1, Len(fileLyrics.FileName) - 4)
        tmpSplit = InStr(1, tmpFilename, "¤", vbTextCompare)
        tmpTitle = Mid(tmpFilename, tmpSplit + 1, Len(tmpFilename))
        tmpGroup = Mid(tmpFilename, 1, tmpSplit - 1)
        
        frmProgress.lblAction.Caption = fileLyrics.FileName
        DoEvents
        
        '// Capitalize Title + Group first letter.
        tmpLetter = Mid(tmpTitle, 1, 1)
        tmpLetter = UCase(tmpLetter)
        tmpString = tmpLetter & Mid(tmpTitle, 2, Len(tmpTitle))
        tmpTitle = tmpString
        
        tmpLetter = Mid(tmpGroup, 1, 1)
        tmpLetter = UCase(tmpLetter)
        tmpString = tmpLetter & Mid(tmpGroup, 2, Len(tmpGroup))
        tmpGroup = tmpString
        
        tvLyrics.Nodes.Add "Lyrics", tvwChild, tmpGroup, tmpGroup, 3
        tvLyrics.Nodes.Add tmpGroup, tvwChild, tmpFilename, tmpTitle, 2
        frmProgress.pBar.Value = I
        lblEntries.Caption = I
    Next I

    '// Check if registered.
    If GetSetting("LyricsManagement", "Settings", "Registered", "0") = "0" Then
        frmMain.tmrUnregistered.Enabled = True
        Me.Caption = Me.Caption & " - UNREGISTERED"
    Else
        frmMain.tmrUnregistered.Enabled = False
    End If

    tvLyrics.Nodes.Item(1).Expanded = True
    Unload frmProgress

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

    '// Exit program
    End

End Sub

Private Sub lblFilename_Change()

    '// Set tooltip
    lblFilename.ToolTipText = lblFilename.Caption
    lblFileshow.ToolTipText = lblFilename.Caption
    
    '// Strip down
    If Len(lblFilename.Caption) > 33 Then
        lblFileshow.Caption = Mid(lblFilename.Caption, 1, 33) & "..."
    Else
        lblFileshow.Caption = lblFilename.Caption
    End If

End Sub

Private Sub mnuAbout_Click()

    '// Show about
    frmAbout.Show 1

End Sub

Private Sub mnuAddDirectory_Click()

    '// Show add directory dialog
    Me.Hide
    frmAddDirectory.Show

End Sub

Private Sub mnuAddSong_Click()

    '// Show add song dialog.
    frmAddSong.Show
    Me.Hide

End Sub

Private Sub mnuDeleteSong_Click()

    On Error Resume Next

    '// Delete song.
    Dim Answer
    Dim tmpFile As String
    
    Answer = MsgBox("Do you want to remove this song:" & vbCrLf & vbCrLf & lblArtist.Caption & " - " & lblTitle.Caption, vbCritical + vbYesNo, "Remove song?")
    
    If Answer = vbYes Then
        tmpFile = App.path & "\data\" & tvLyrics.SelectedItem.Key & ".lyr"
        Kill tmpFile
        tmpFile = App.path & "\data\" & tvLyrics.SelectedItem.Key & ".sng"
        Kill tmpFile
        tmrRefresh.Enabled = True
    End If

End Sub

Private Sub mnuEditSong_Click()

    '// Show info
    frmAddSong.txtAlbum.Text = lblAlbum.Caption
    frmAddSong.txtArtist.Text = lblArtist.Caption
    frmAddSong.txtFilename.Text = lblFilename.Caption
    frmAddSong.txtLength.Text = lblLength.Caption
    frmAddSong.txtText.Text = txtSong.Text
    frmAddSong.txtTitle.Text = lblTitle.Caption
    frmAddSong.cboGenre.Text = lblGenre.Caption
    frmAddSong.cboYear.Text = lblYear.Caption
    frmAddSong.Show

End Sub

Private Sub mnuExit_Click()

    '// Exit the program
    End

End Sub

Private Sub mnuJukebox_Click()

    '// Show jukebox
    Call AddToTray(Me.Icon, Me.Caption, Me)
    frmJukebox.sldVolume.Value = 1250
    frmJukebox.Show

End Sub

Private Sub mnuOptions_Click()

    '// Show options
    frmOptions.Show

End Sub

Private Sub mnuRefresh_Click()

    '// Refresh list.
    tmrRefresh.Enabled = True

End Sub

Private Sub mnuRegister_Click()

    '// Show register
    frmRegister.Show

End Sub

Private Sub mnuSearch_Click()

    '// Show search dialog.
    frmSearch.Show

End Sub

Private Sub mnuShow_Click()

    '// Remove from tray
    Call ShowFormAgain(Me)

End Sub

Private Sub mnuTextSearch_Click()

    '// Show info
    frmAddSong.txtAlbum.Text = lblAlbum.Caption
    frmAddSong.txtArtist.Text = lblArtist.Caption
    frmAddSong.txtFilename.Text = lblFilename.Caption
    frmAddSong.txtLength.Text = lblLength.Caption
    frmAddSong.txtText.Text = txtSong.Text
    frmAddSong.txtTitle.Text = lblTitle.Caption
    frmAddSong.cboGenre.Text = lblGenre.Caption
    frmAddSong.cboYear.Text = lblYear.Caption
    Load frmAddSong

    '// Show search
    frmSearch.txtSearch.Text = tvLyrics.SelectedItem.Text
    frmSearch.tmrSearch.Enabled = True
    frmSearch.Command4.Enabled = True
    frmSearch.Command3.Enabled = False
    frmSearch.Show

End Sub

Private Sub tmrRefresh_Timer()

    tmrRefresh.Enabled = False
    On Error Resume Next

    '// Declares
    Dim I As Long
    Dim tmpTitle As String
    Dim tmpGroup As String
    Dim tmpFilename As String
    Dim tmpString As String
    Dim tmpLetter As String
    Dim tmpSplit As Long
    
    '// Make sure data directory exist.
    If Dir(App.path & "\data", vbDirectory) = "" Then MkDir App.path & "\data"
    fileLyrics.Pattern = "*.lyr"
    fileLyrics.path = App.path & "\data"
    fileLyrics.Refresh

    '// Open the progressbar
    frmProgress.pBar.Min = 0
    frmProgress.pBar.Max = fileLyrics.ListCount - 1
    frmProgress.Show
    DoEvents

    '// Load up the startup data.
    tvLyrics.Nodes.Clear
    tvLyrics.Nodes.Add , , "Lyrics", "Lyrics", 1

    For I = 0 To fileLyrics.ListCount - 1 Step 1
        fileLyrics.ListIndex = I
        tmpFilename = Mid(fileLyrics.FileName, 1, Len(fileLyrics.FileName) - 4)
        tmpSplit = InStr(1, tmpFilename, "¤", vbTextCompare)
        tmpTitle = Mid(tmpFilename, tmpSplit + 1, Len(tmpFilename))
        tmpGroup = Mid(tmpFilename, 1, tmpSplit - 1)
        
        '// Capitalize Title + Group first letter.
        tmpLetter = Mid(tmpTitle, 1, 1)
        tmpLetter = UCase(tmpLetter)
        tmpString = tmpLetter & Mid(tmpTitle, 2, Len(tmpTitle))
        tmpTitle = tmpString
        
        tmpLetter = Mid(tmpGroup, 1, 1)
        tmpLetter = UCase(tmpLetter)
        tmpString = tmpLetter & Mid(tmpGroup, 2, Len(tmpGroup))
        tmpGroup = tmpString
        
        tvLyrics.Nodes.Add "Lyrics", tvwChild, tmpGroup, tmpGroup, 3
        tvLyrics.Nodes.Add tmpGroup, tvwChild, tmpFilename, tmpTitle, 2
        frmProgress.pBar.Value = I
        lblEntries.Caption = I
    Next I

    tvLyrics.Nodes.Item(1).Expanded = True
    Unload frmProgress

End Sub

Private Sub tmrUnregistered_Timer()

    '// Show nag.
    MsgBox "This program is unregistered, please register it, thanks!", vbCritical, "Unregistered!"

End Sub

Private Sub tvLyrics_Click()

    On Error Resume Next

    '// Check if the filename exist.
    Dim tmpFile As String
    Dim tmpFilename As String
    Dim tmpTitle As String
    Dim tmpGroup As String
    Dim tmpString As String
    Dim tmpLetter As String
    Dim tmpSplit As Long
    Dim tmpArtist As String
    Dim tmpAlbum As String
    Dim tmpYear As String
    Dim tmpLength As String
    Dim tmpGenre As String
    Dim tmpText As String
    Dim tmpTextfile As String

    txtSong.Text = ""
    tmpFile = tvLyrics.SelectedItem.Key
    
    If Dir(App.path & "\data\" & tmpFile & ".lyr") = "" Then
        mnuEditSong.Enabled = False
        mnuDeleteSong.Enabled = False
        mnuTextSearch.Enabled = False
        lblFilename.Caption = ""
        Exit Sub
    Else
        mnuTextSearch.Enabled = True
        mnuEditSong.Enabled = True
        mnuDeleteSong.Enabled = True
    End If

    '// Load the song information
    tmpFilename = Mid(tmpFile, 1, Len(tmpFile))
    tmpSplit = InStr(1, tmpFilename, "¤", vbTextCompare)
    tmpTitle = Mid(tmpFilename, tmpSplit + 1, Len(tmpFilename))
    tmpGroup = Mid(tmpFilename, 1, tmpSplit - 1)

    '// Capitalize Title + Group first letter.
    tmpLetter = Mid(tmpTitle, 1, 1)
    tmpLetter = UCase(tmpLetter)
    tmpString = tmpLetter & Mid(tmpTitle, 2, Len(tmpTitle))
    tmpTitle = tmpString
        
    tmpLetter = Mid(tmpGroup, 1, 1)
    tmpLetter = UCase(tmpLetter)
    tmpString = tmpLetter & Mid(tmpGroup, 2, Len(tmpGroup))
    tmpGroup = tmpString

    lblSongname.Caption = tmpGroup & " - " & tmpTitle

    Dim tmpFreeFile
    tmpFreeFile = FreeFile
    
    '// Get the song information.
    Open App.path & "\data\" & tmpFile & ".lyr" For Random As tmpFreeFile
        Get tmpFreeFile, 1, tmpFile
        Get tmpFreeFile, 2, tmpTitle
        Get tmpFreeFile, 3, tmpArtist
        Get tmpFreeFile, 4, tmpAlbum
        Get tmpFreeFile, 5, tmpYear
        Get tmpFreeFile, 6, tmpLength
        Get tmpFreeFile, 7, tmpGenre
        Get tmpFreeFile, 8, tmpTextfile
    Close tmpFreeFile
    
    '// Display the information.
    lblTitle.Caption = tmpTitle
    lblFilename.Caption = tmpFile
    lblArtist.Caption = tmpArtist
    lblAlbum.Caption = tmpAlbum
    lblYear.Caption = tmpYear
    lblLength.Caption = tmpLength
    lblGenre.Caption = tmpGenre
    
    '// Load song
    Dim Var1
    Open App.path & "\data\" & tmpTextfile For Input As #1
        FileLength = LOF(1)
        Var1 = Input(FileLength, #1)
        txtSong.Text = Var1
    Close #1

End Sub

Private Sub tvLyrics_DblClick()

    '// Check if there are any filenames to load
    If lblFilename.Caption = "" Then
        Exit Sub
    End If

    '// Load jukebox
    'frmJukebox.Player.FileName = lblFilename.Caption
    'frmJukebox.Player.BaseURL = lblFilename.Caption
    frmJukebox.Player.FileName = lblFilename.Caption
    frmJukebox.lblTitle.Caption = lblArtist.Caption & " - " & lblTitle.Caption
    
    '// Load the file.
    Dim ObjMP3 As clsMP3
    Set ObjMP3 = New clsMP3
    
    '// Read the MP3 data
    ObjMP3.ReadMP3 lblFilename.Caption
    
    '// Set the values
    frmJukebox.lblFreq.Caption = ObjMP3.Frequency
    
    frmJukebox.lblSong.Caption = txtSong.Text
    frmJukebox.lblSong.Top = 1800
    frmJukebox.sldVolume.Value = 1250
    frmJukebox.Command1.Enabled = False
    frmJukebox.Command2.Enabled = True
    frmJukebox.Command3.Enabled = True
    frmJukebox.Command8.Enabled = False
    frmJukebox.Show
    frmJukebox.sldProgress.Max = frmJukebox.Player.Duration

End Sub

Private Sub tvLyrics_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Check if right click.
    If Button = 2 Then
        PopupMenu mnuSongs, , , , mnuAddSong
    End If

End Sub
