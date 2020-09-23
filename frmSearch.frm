VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lyrics Management - Online search"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Search:"
      Height          =   1335
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   7215
      Begin VB.CommandButton Command5 
         Caption         =   "&Go"
         Height          =   285
         Left            =   6240
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtSearch2 
         Height          =   285
         Left            =   960
         TabIndex        =   30
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label11 
         Caption         =   $"frmSearch.frx":08CA
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Timer tmrSearch 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   1320
   End
   Begin VB.Timer tmrText 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1320
   End
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Song text:"
      Height          =   3015
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   7215
      Begin VB.CommandButton Command4 
         Caption         =   "&Add >"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6120
         TabIndex        =   25
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save"
         Height          =   375
         Left            =   6120
         TabIndex        =   24
         Top             =   2520
         Width           =   975
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
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label lblTitle 
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label lblArtist 
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.TreeView tvResult 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2143
      _Version        =   393217
      Indentation     =   353
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList"
      Appearance      =   1
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7440
      TabIndex        =   12
      Top             =   7815
      Width           =   7500
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Result:"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   45
         Width           =   735
      End
      Begin VB.Label lblResult 
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   45
         Width           =   6615
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7500
      TabIndex        =   9
      Top             =   0
      Width           =   7500
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lyrics Management - Online search"
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
         TabIndex        =   10
         Top             =   120
         Width           =   6015
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmSearch.frx":0952
         Top             =   120
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8040
         Y1              =   710
         Y2              =   710
      End
   End
   Begin VB.ListBox lst_Result 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   2280
      TabIndex        =   8
      Top             =   8760
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Refresh"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   8760
      Width           =   7215
   End
   Begin VB.ListBox lstResult_Artist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1440
      TabIndex        =   5
      Top             =   8760
      Width           =   3255
   End
   Begin VB.Timer tmrCheckData 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   1320
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   120
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtIncoming 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   7920
      Width           =   7815
   End
   Begin VB.ComboBox cboSection 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSearch.frx":121C
      Left            =   6240
      List            =   "frmSearch.frx":1229
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmSearch.frx":1242
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":2BF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Note:"
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
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Right click the song you want, and click 'Get text' to get the song text."
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   3000
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "Search for:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Search result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Menu mnuSong 
      Caption         =   "&Song"
      Visible         =   0   'False
      Begin VB.Menu mnuGetSong 
         Caption         =   "&Get text..."
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Declares
    Dim tmpSearch As String
    Dim tmpEntry As String
    Dim tmpSection As String
    
    '// Check if there's anything to search for.
    If txtSearch.Text = "" Then Exit Sub
    
    If GetSetting("LyricsManagement", "Settings", "Registered", "0") = "0" Then
        If tmpSearchCount = 5 Then
            MsgBox "You have reached your maximum search for this session '5', register the program to have unlimited search.", vbCritical, "Search limit reached!"
            Exit Sub
        Else
            tmpSearchCount = tmpSearchCount + 1
        End If
    End If
    
    frmSearchProgress.lblStatus.Caption = "Searching.."
    frmSearchProgress.Show
    DoEvents
    tmrStatus.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    tmpEntry = txtSearch.Text
    tmpSection = cboSection.Text

    '// Create the search criteria, make sure all spaces
    '// are made to %20's.
    tmpEntry = Replace(tmpEntry, " ", "%20")
    tvResult.Nodes.Clear
    frmSearch.lst_Result.Clear
    frmSearch.lstResult_Artist.Clear

    tvResult.Nodes.Clear
    tvResult.Nodes.Add , , "Search", "Search result", 1

    tmrCheckData.Enabled = True
    txtIncoming.Text = Inet.OpenURL("http://www.letssingit.com/cgi-exe/am.cgi?a=search&p=1&s=" & tmpEntry & "&l=" & tmpSection)

    '// Check if we want to clear the search field.
    If GetSetting("LyricsManagement", "Settings", "SearchClear", "0") = "1" Then
        txtSearch.Text = ""
    Else
    End If

End Sub

Private Sub Command2_Click()

    On Error Resume Next

    '// Sort the result.
    Dim I As Long
    Dim tmpArtist As String
    Dim tmpArtistURL As String
    Dim tmpArtistInfo As String
    Dim tmpTitle As String
    Dim tmpTitleURL As String
    Dim tmpTitleInfo As String
    Dim tmpSection As Long
    Dim tmpStart As Long
    Dim tmpEnd As Long

    tmpSection = 1
    frmSearchProgress.lblStatus.Caption = "Getting list.."
    DoEvents
    For I = 0 To lstResult_Artist.ListCount - 1 Step 1
        lstResult_Artist.ListIndex = I
        
        If tmpSection = 1 Then
            tmpArtist = lstResult_Artist.Text
            tmpSection = 2
        ElseIf tmpSection = 2 Then
            tmpTitle = lstResult_Artist.Text
            tmpSection = 1
            lst_Result.AddItem tmpArtist & " ¤ " & tmpTitle
        End If
    Next I
    
    For I = 0 To lst_Result.ListCount - 1 Step 1
        lst_Result.ListIndex = I
        
        '// Add the artist
        
        tmpArtistInfo = lst_Result
        tmpEnd = InStr(1, tmpArtistInfo, "¤", vbTextCompare)
        tmpArtistInfo = Mid(tmpArtistInfo, 1, tmpEnd - 1)
        tmpEnd = InStr(1, tmpArtistInfo, "§", vbTextCompare)
        tmpArtist = Mid(tmpArtistInfo, 1, tmpEnd - 1)
        tmpArtistURL = Mid(tmpArtistInfo, tmpEnd + 2, Len(tmpArtistInfo) - 1)
       
        tvResult.Nodes.Add "Search", tvwChild, tmpArtistURL, tmpArtist, 3
        
        '// Add the song
        tmpTitleInfo = lst_Result
        tmpStart = InStr(1, tmpTitleInfo, "¤", vbTextCompare)
        tmpTitleInfo = Mid(tmpTitleInfo, tmpStart, Len(tmpTitleInfo))
        tmpEnd = InStr(1, tmpTitleInfo, "§", vbTextCompare)
        tmpTitle = Mid(tmpTitleInfo, 3, tmpEnd - 1)
        tmpTitleURL = Mid(tmpTitleInfo, tmpEnd + 2, Len(tmpTitleInfo))
        tmpTitle = RTrim(tmpTitle)
        tmpTitle = Mid(tmpTitle, 1, Len(tmpTitle) - 2)
       
        tvResult.Nodes.Add tmpArtistURL, tvwChild, tmpTitleURL, tmpTitle, 2
        
    Next I
    tvResult.Nodes.Item(1).Expanded = True
    Unload frmSearchProgress

End Sub

Private Sub Command3_Click()

    On Error Resume Next

    '// Check if necessary info are collected.
    If lblArtist.Caption = "" Then
        MsgBox "You have not selected an artist yet, do that first!", vbCritical, "Error!"
        Exit Sub
    End If
    
    If lblTitle.Caption = "" Then
        MsgBox "You have not selected any song yet, do that first!", vbCritical, "Error!"
        Exit Sub
    End If
    
    Dim Answer
    If txtSong.Text = "" Then
        Answer = MsgBox("Do you want to add this song, even though you have not downloaded the song text?", vbCritical + vbYesNo, "Add song?")
        If Answer = vbNo Then
            Exit Sub
        End If
    End If
    
    '// Add song
    frmAddSong.txtArtist.Text = lblArtist.Caption
    frmAddSong.txtTitle.Text = lblTitle.Caption
    frmAddSong.txtText.Text = txtSong.Text
    frmAddSong.Show
    Unload Me

End Sub

Private Sub Command4_Click()

    On Error Resume Next

    '// Send text
    frmAddSong.txtText.Text = txtSong.Text
    Command4.Enabled = False
    frmAddSong.Show
    Unload Me

End Sub

Private Sub Command5_Click()

    On Error Resume Next

    '// Check if there is anything to search for.
    If txtSearch2.Text = "" Then Exit Sub

    '// Declares
    Dim I As Long
    Dim tmpFound As Long
    
    Command2_Click
    
    tmpFound = 0
    Frame2.Caption = "Search:"
    For I = 1 To tvResult.Nodes.Count Step 1
        Frame2.Caption = "Search: " & I & "/" & tvResult.Nodes.Count
        If InStr(1, tvResult.Nodes(I).Text, txtSearch2.Text, vbTextCompare) = "0" Then
            tvResult.Nodes.Item(I).Expanded = False
        Else
            tvResult.Nodes(I).Selected = True
            tmpFound = tmpFound + 1
        End If
    Next I
    Frame2.Caption = "Search: Found - " & tmpFound & " entries."
    
    '// Check if we want to clear the search field.
    If GetSetting("LyricsManagement", "Settings", "SearchClear", "0") = "1" Then
        txtSearch2.Text = ""
    Else
    End If

End Sub

Private Sub Form_Load()

    '// Set default values.
    cboSection.Text = "song"

    '// Load up the startup data.
    tvResult.Nodes.Clear
    tvResult.Nodes.Add , , "Search", "Search result", 1

End Sub

Private Sub lstResult_Click()

    '// Debug
    Text1.Text = lstResult.Text

End Sub

Private Sub lst_Result_Click()

    '// Debug
    Text1.Text = lst_Result.Text

End Sub

Private Sub lstResult_Artist_Click()

    '// Debug
    Text1.Text = lstResult_Artist.Text

End Sub

Private Sub mnuGetSong_Click()

    '// Get information
    tmrText.Enabled = True
    Text1.Text = Inet.OpenURL(tvResult.SelectedItem.Key)

End Sub

Private Sub tmrCheckData_Timer()

    On Error Resume Next

    '// Declares
    Dim tmpStart As Long
    Dim tmpEnd As Long
    Dim tmpString As String
    Dim tmpData As String
    Dim I As Long
    Dim tmpUrl As String

    '// Check if we are still getting data, if so wait.
    If Inet.StillExecuting = True Then
        frmSearchProgress.lblStatus.Caption = "Searching.."
    Else
        frmSearchProgress.lblStatus.Caption = "Processing info.."
        DoEvents
        tmrStatus.Enabled = False
        lblStatus.Caption = "Search complete!"
        tmrCheckData.Enabled = False
        tmpData = txtIncoming.Text
        
        tmpStart = InStr(1, tmpData, "class=loc", vbTextCompare)
        tmpEnd = InStr(tmpStart, tmpData, "</td>", vbTextCompare)
        tmpString = Mid(tmpData, tmpStart + 10, tmpEnd - 20)
        tmpEnd = InStr(1, tmpString, "</TD>", vbTextCompare)
        tmpString = Mid(tmpString, 1, tmpEnd - 1)
        lblResult.Caption = tmpString

        Dim tmpArtist As String
        Dim tmpTitle As String

        For I = 0 To Len(tmpData) Step 1
            tmpString = Mid(tmpData, I, 9)
            
                If tmpString = "class=res" Then
                    tmpEnd = InStr(I + 1, tmpData, "</A>", vbTextCompare)
                    tmpString = Mid(tmpData, I + 10, tmpEnd - 50)
                    tmpString = Trim(tmpString)
                    tmpEnd = InStr(1, tmpString, "</A>", vbTextCompare)
                    tmpString = Mid(tmpString, 1, tmpEnd)
                    tmpStart = InStr(1, tmpString, ">", vbTextCompare)
                    tmpString = Mid(tmpString, tmpStart + 1, Len(tmpString))
                    tmpString = Trim(tmpString)
                    tmpString = Mid(tmpString, 1, Len(tmpString) - 1)
                    
                    tmpStart = InStr(I, tmpData, "<A href=", vbTextCompare)
                    tmpEnd = InStr(I, tmpData, ">", vbTextCompare)
                    tmpUrl = Mid(tmpData, tmpStart, tmpEnd)
                    tmpEnd = InStr(1, tmpUrl, ">", vbTextCompare)
                    tmpUrl = Mid(tmpUrl, 10, tmpEnd)
                    tmpUrl = Trim(tmpUrl)
                    tmpEnd = InStr(1, tmpUrl, ">", vbTextCompare)
                    tmpUrl = "http://www.letssingit.com" & Mid(tmpUrl, 1, tmpEnd - 2)
                    
                    lstResult_Artist.AddItem tmpString & " § " & tmpUrl
                End If

        Next I
        Command2.Enabled = True
        Command1.Enabled = True

        '// Display
        Command2_Click

    End If

End Sub

Private Sub tmrSearch_Timer()

    '// Declares
    Dim tmpSearch As String
    Dim tmpEntry As String
    Dim tmpSection As String
    
    tmrStatus.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    tmpEntry = txtSearch.Text
    tmpSection = cboSection.Text

    '// Create the search criteria, make sure all spaces
    '// are made to %20's.
    tmpEntry = Replace(tmpEntry, " ", "%20")
    tvResult.Nodes.Clear
    frmSearch.lst_Result.Clear
    frmSearch.lstResult_Artist.Clear

    tvResult.Nodes.Clear
    tvResult.Nodes.Add , , "Search", "Search result", 1

    tmrCheckData.Enabled = True
    txtIncoming.Text = Inet.OpenURL("http://www.letssingit.com/cgi-exe/am.cgi?a=search&p=1&s=" & tmpEntry & "&l=" & tmpSection)
    tmrSearch.Enabled = False

End Sub

Private Sub tmrStatus_Timer()

    '// Show movement on status bar.
    Dim tmpStart As Long
    tmpStart = 1
    
    If tmpStart = 1 Then
        lblStatus.Caption = "Searching."
        lblStatus.Refresh
        DoEvents
        tmpStart = 2
    ElseIf tmpStart = 2 Then
        lblStatus.Caption = "Searching.."
        lblStatus.Refresh
        DoEvents
        tmpStart = 3
    ElseIf tmpStart = 3 Then
        lblStatus.Caption = "Searching..."
        lblStatus.Refresh
        DoEvents
        tmpStart = 1
    End If

End Sub

Private Sub tmrText_Timer()

    On Error Resume Next

    '// Declares
    Dim tmpStart As Long
    Dim tmpEnd As Long
    Dim tmpString As String

    '// Check if we are still getting data, if so wait.
    If Inet.StillExecuting = True Then
    Else
        tmrText.Enabled = False
        
        tmpStart = InStr(1, Text1.Text, "<PRE", vbTextCompare)
        tmpEnd = InStr(1, Text1.Text, "</PRE", vbTextCompare)
        tmpString = Mid(Text1.Text, tmpStart + 29, tmpEnd)
        tmpEnd = InStr(1, tmpString, "</PRE", vbTextCompare)
        tmpString = Mid(tmpString, 1, tmpEnd - 1)
        txtSong.Text = tmpString
        
        txtSong.Text = Replace(txtSong.Text, Chr(10), vbCrLf)
        
    End If

End Sub

Private Sub tvResult_Click()

    '// Debug
    Text1.Text = tvResult.SelectedItem.Key
    
    '// Set Artist/Song.
    If InStr(1, tvResult.SelectedItem.Key, "artist", vbTextCompare) = 0 Then
        lblTitle.Caption = tvResult.SelectedItem.Text
    Else
        lblArtist.Caption = tvResult.SelectedItem.Text
        lblTitle.Caption = ""
    End If

End Sub

Private Sub tvResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// See if right button click.
    If Button = 2 Then
        PopupMenu mnuSong
    End If

End Sub

Private Sub txtSearch_Click()

    '// Select the full text
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.Text)

End Sub

Private Sub txtSearch2_Click()

    '// Select the full text
    txtSearch2.SelStart = 0
    txtSearch2.SelLength = Len(txtSearch2.Text)

End Sub
