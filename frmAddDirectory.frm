VERSION 5.00
Begin VB.Form frmAddDirectory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lyrics Management - Add directory"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddDirectory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2040
      ScaleHeight     =   2415
      ScaleWidth      =   4215
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command6 
         Caption         =   "&OK"
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   1920
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   10
         ScaleHeight     =   255
         ScaleWidth      =   4185
         TabIndex        =   25
         Top             =   10
         Width           =   4190
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Status report"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   4215
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   4200
            Y1              =   240
            Y2              =   240
         End
      End
      Begin VB.Label lblSkipped 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblAdded 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblScanned 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   31
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAddDirectory.frx":08CA
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Files skipped:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Files added:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Files scanned:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line Line8 
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   2400
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   4200
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line4 
         X1              =   4200
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2400
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   23
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add files..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status:"
      Height          =   3015
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   6735
      Begin VB.CommandButton Command2 
         Caption         =   "&Start search"
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   6495
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblFound 
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
         Left            =   1200
         TabIndex        =   18
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblFiles 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label lblDirectory 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   5415
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   6600
         Y1              =   730
         Y2              =   730
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   6600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Files:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options:"
      Height          =   855
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   6735
      Begin VB.CheckBox chkReadID3 
         Caption         =   "Read ID3 tag of files"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkRecSearch 
         Caption         =   "Recursively search"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select directory to add:"
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   6120
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7605
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Add a whole directory"
         Height          =   255
         Left            =   720
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   120
         Width           =   5655
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmAddDirectory.frx":0951
         Top             =   120
         Width           =   480
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   8760
         Y1              =   710
         Y2              =   710
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   8760
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "frmAddDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you many not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const vbDot = 46
Private Const MAXDWORD = &HFFFFFFFF
'Private Const MAX_PATH = 260
'Private Const INVALID_HANDLE_VALUE = -1
'Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type

'Private Type WIN32_FIND_DATA
'   dwFileAttributes As Long
'   ftCreationTime As FILETIME
'   ftLastAccessTime As FILETIME
'   ftLastWriteTime As FILETIME
'   nFileSizeHigh As Long
'   nFileSizeLow As Long
'   dwReserved0 As Long
'   dwReserved1 As Long
'   cFileName As String * MAX_PATH
'   cAlternate As String * 14
'End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

'Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub GetFileInformation(FP As FILE_PARAMS)

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
  'FP.sFileRoot contains the path to search.
  'FP.sFileNameExt contains the full path and filespec.
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   
  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      Do
         
        'Even though this routine uses file specs,
        '*.* is still valid and will cause the search
        'to return folders as well as files, so a
        'check against folders is still required.
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
                 FILE_ATTRIBUTE_DIRECTORY Then

           'this is where you add code to store
           'or display the returned file listing.
           '
           'if you want the file name only, save 'sTmp'.
           'if you want the full path, save 'sRoot & sTmp'

           'remove trailing nulls
            FP.Count = FP.Count + 1
            sTmp = TrimNull(WFD.cFileName)
            List1.AddItem sRoot & sTmp

         End If
         
      Loop While FindNextFile(hFile, WFD)
      
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If

End Sub


Private Sub SearchForFiles(FP As FILE_PARAMS)

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   
  'obtain handle to the first match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then
   
     'This is where the method obtains the file
     'list and data for the folder passed.
      Call GetFileInformation(FP)

      Do
      
        'if the returned item is a folder...
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
           '..and the Recurse flag was specified
            If FP.bRecurse Then
            
              'and if the folder is not the default
              'self and parent folders (a . or ..)
               If Asc(WFD.cFileName) <> vbDot Then
               
                 '..then the item is a real folder, which
                 'may contain other sub folders, so assign
                 'the new folder name to FP.sFileRoot and
                 'recursively call this function again with
                 'the amended information.

                 'remove trailing nulls
                  FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
                  Call SearchForFiles(FP)
                  
               End If
               
            End If
            
         End If
         
     'continue looping until FindNextFile returns
     '0 (no more matches)
      Loop While FindNextFile(hFile, WFD)
      
     'close the find handle
      hFile = FindClose(hFile)
   
   End If
   
End Sub


Private Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If Right$(sPath, 1) <> "\" Then
         QualifyPath = sPath & "\"
   Else: QualifyPath = sPath
   End If
      
End Function


Private Function TrimNull(startstr As String) As String

  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function

Public Function rgbCountFilesAll(sSourcePath As String, sFileType As String) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim SA As SECURITY_ATTRIBUTES
   Dim hFile As Long
   Dim bNext As Long
   Dim fCount As Long
   Dim currFile As String
      
  'Start searching for files in the Source directory by
  'obtaining a file handle to the first file matching the
  'filespec passed
   hFile = FindFirstFile(sSourcePath & sFileType, WFD)
   
   If (hFile = INVALID_HANDLE_VALUE) Then
     
      'no match, so bail out
       rgbCountFilesAll = 0
       Exit Function
       
   End If
       
  'must have at least one, so ...
   If hFile Then
      
      Do
         
        'increment the counter and find the next
        'file matching the filespec
         fCount = fCount + 1
         bNext = FindNextFile(hFile, WFD)
               
      Loop Until bNext = 0
      
   End If
      
  'Close the search handle
   Call FindClose(hFile)
      
  'return the number of files found
   rgbCountFilesAll = fCount
   
End Function

Private Sub Command1_Click()

    '// Declares
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim path As String
    Dim pos As Integer

    txtDirectory.Text = ""

    '// Fill the BROWSEINFO structure with the
    '// needed data. To accomodate comments, the
    '// With/End With sytax has not been used, though
    '// it should be your'final' version.

    '// hwnd of the window that receives messages
    '// from the call. Can be your application
    '// or the handle from GetDesktopWindow().
    bi.hOwner = Me.hWnd

    '// Pointer to the item identifier list specifying
    '// the location of the "root" folder to browse from.
    '// If NULL, the desktop folder is used.
    bi.pidlRoot = 0&

    '// message to be displayed in the Browse dialog
    bi.lpszTitle = "Add what directory?"

    '// the type of folder to return.
    bi.ulFlags = BIF_RETURNONLYFSDIRS

    '// show the browse for folders dialog
    pidl = SHBrowseForFolder(bi)
 
    '// the dialog has closed, so parse & display the
    '// user's returned folder selection contained in pidl
    path = Space$(MAX_PATH)

    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
        pos = InStr(path, Chr$(0))
        txtDirectory.Text = Left(path, pos - 1)
        lblDirectory.Caption = txtDirectory.Text
        
        If Right(lblDirectory.Caption, 1) <> "\" Then lblDirectory.Caption = lblDirectory.Caption & "\"
        
        Dim sSourcePath As String
        Dim sDestination As String
        Dim sFileType As String
        Dim numFiles As Long
      
        '// set the appropriate initializing values
        sSourcePath = lblDirectory.Caption
        sFileType = "*.mp3"
      
        '// get the count
        numFiles = rgbCountFilesAll(sSourcePath, sFileType)
        lblFiles.Caption = numFiles

    End If

    Call CoTaskMemFree(pidl)

End Sub

Private Sub Command2_Click()

    '// Declares
    Dim FP As FILE_PARAMS
    Dim tstart As Single
    Dim tend As Single
   
    '// setting the list visibility to false
    '// increases the load time
    List1.Clear
    List1.Visible = False
   
    '// Check if directory spesified.
    If lblDirectory.Caption = "" Then Exit Sub
   
    '// set up search params
    With FP
        .sFileRoot = lblDirectory.Caption
        .sFileNameExt = "*.mp3"
        .bRecurse = chkRecSearch.Value = 1
    End With
   
    '// get start time, get files, and get finish time
    tstart = GetTickCount()
    Call SearchForFiles(FP)
    tend = GetTickCount()
   
    List1.Visible = True
   
    '// show the results
    lblFound.Caption = Format$(FP.Count, "###,###,###,##0") & " found (" & FP.sFileNameExt & ")"
    lblTime.Caption = FormatNumber((tend - tstart) / 1000, 2) & " seconds"
    
    If List1.ListCount > 0 Then Command3.Enabled = True

End Sub

Private Sub Command3_Click()

    On Error Resume Next

    '// Declares
    Dim I As Long
    Dim ObjMP3 As clsMP3
    Set ObjMP3 = New clsMP3
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
    Dim tmpSkipped As Long
    Dim tmpAdded As Long
    
    tmpSkipped = 0
    tmpAdded = 0
    
    Me.Hide
    frmProgress.lblAction.Caption = "Scanning files..."
    frmProgress.pBar.Min = 0
    frmProgress.pBar.Max = List1.ListCount - 1
    frmProgress.Show
    For I = 0 To List1.ListCount - 1 Step 1
        List1.ListIndex = I
          
        '// Read the MP3 data
        ObjMP3.ReadMP3 List1.Text
        frmProgress.lblAction.Caption = GetFileName(List1.Text)
        frmProgress.pBar.Value = I
               
        '// Check if we got any data, if not, skip this file.
        If ObjMP3.Artist = "" Or ObjMP3.Songname = "" Then
            tmpSkipped = tmpSkipped + 1
        Else
            tmpAdded = tmpAdded + 1
            tmpFreeFile = FreeFile
            tmpFilename = ObjMP3.Artist & "¤" & ObjMP3.Songname & ".lyr"
            tmpTitle = ObjMP3.Songname
            tmpArtist = ObjMP3.Artist
            tmpAlbum = ObjMP3.Album
            tmpYear = ObjMP3.Year
            tmpLength = ObjMP3.Duration
            tmpGenre = ObjMP3.Genre
            tmpText = ""
            tmpFile = List1.Text
            tmpTextfile = ObjMP3.Artist & "¤" & ObjMP3.Songname & ".sng"
            
            '// Save the song information.
            Open App.path & "\data\" & tmpFilename For Random As tmpFreeFile
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
            If Dir(App.path & "\data\" & tmpTextfile) = "" Then
            Else
                Kill App.path & "\data\" & tmpTextfile
            End If
        
            Open App.path & "\data\" & tmpTextfile For Append As tmpFreeFile
                Print #tmpFreeFile, tmpText
            Close tmpFreeFile
        
        End If

        DoEvents
       
    Next I
    Unload frmProgress
    Me.Show
    Command5.Visible = True
    lblAdded.Caption = tmpAdded
    lblSkipped.Caption = tmpSkipped
    lblScanned.Caption = lblFound.Caption
    Picture1.Visible = True

End Sub

Private Sub Command4_Click()

    '// Unload me
    Unload Me

End Sub

Private Sub Command5_Click()

    '// Unload me
    Unload Me

End Sub

Private Sub Command6_Click()

    '// Hide
    Picture1.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    '// Show main form
    frmMain.Show
    Unload Me

End Sub
