VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Puzzles"
   ClientHeight    =   4752
   ClientLeft      =   132
   ClientTop       =   360
   ClientWidth     =   6612
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   25
      Top             =   4500
      Width           =   6612
      _ExtentX        =   11663
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3704
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3704
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3704
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PicClip.PictureClip PicClip 
      Left            =   4680
      Top             =   3120
      _ExtentX        =   2773
      _ExtentY        =   2138
      _Version        =   393216
      Rows            =   5
      Cols            =   5
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   24
      Left            =   3480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   24
      Top             =   3480
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   23
      Left            =   2640
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   23
      Top             =   3480
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   22
      Left            =   1800
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   22
      Top             =   3480
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   21
      Left            =   960
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   21
      Top             =   3480
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   20
      Left            =   120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   20
      Top             =   3480
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   19
      Left            =   3480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   19
      Top             =   2640
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   18
      Left            =   2640
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   18
      Top             =   2640
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   17
      Left            =   1800
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   17
      Top             =   2640
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   16
      Left            =   960
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   16
      Top             =   2640
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   15
      Left            =   120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   15
      Top             =   2640
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   14
      Left            =   3480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   14
      Top             =   1800
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   13
      Left            =   2640
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   13
      Top             =   1800
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   12
      Left            =   1800
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   12
      Top             =   1800
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   11
      Left            =   960
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   11
      Top             =   1800
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   10
      Left            =   120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   10
      Top             =   1800
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   9
      Left            =   3480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   9
      Top             =   960
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   8
      Left            =   2640
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   8
      Top             =   960
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   7
      Left            =   1800
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   7
      Top             =   960
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   6
      Left            =   960
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   6
      Top             =   960
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   5
      Left            =   120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   5
      Top             =   960
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   4
      Left            =   3480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   4
      Top             =   120
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   3
      Left            =   2640
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   3
      Top             =   120
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   2
      Left            =   1800
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   1
      Left            =   960
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin VB.PictureBox picPiece 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   852
      Index           =   0
      Left            =   120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   0
      Top             =   120
      Width           =   852
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuGameBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Puzzles..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Option Explicit

    Private numSwaps As Long    '// Number of swaps

Private Sub Form_Load()
    '// Load the help file
    App.HelpFile = App.Path & "\puzzles.hlp"
End Sub

Private Sub mnuGameExit_Click()
    '// End the program
    End
End Sub

Public Sub mnuGameNew_Click()
    '// If options are not defined, then show Options form
    If Not mvarOptionsDefined Then
        frmOptions.Show
    '// The options are defined
    Else
        PicClip.ROWS = mvarRows     '// Set number of rows of the PicClip
        PicClip.Cols = mvarColumns  '// Set number of columns of the PicClip
        PicClip.Picture = LoadPicture(mvarPicture)      '// Set picture of the PicClip
        
            LoadArray               '// Initialize array
            ScrambleArray           '// Scramble puzzle pieces
            DisplayArray            '// Display array/pieces
            sbStatus.Panels.Item(3).Text = "Correct " & inPlace & " out of 25"
    End If
    
    '// Initialize the values used for the new Game
    mvarOptionsDefined = False
    numSwaps = 0
    
    '// Initialize values in the Status bar
    sbStatus.Panels.Item(1).Text = ""
    sbStatus.Panels.Item(2).Text = "Number of swaps: 0"
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show     '// Show about form
End Sub

Private Function PieceWidth() As Long
    '// Determine width of a single puzzle - piece
    PieceWidth = PicClip.Width / mvarColumns
End Function

Private Function PieceHeight() As Long
    '// Determine height of a single puzzle - piece
    PieceHeight = PicClip.Height / mvarRows
End Function

Private Function LoadArray()
    Dim i   As Integer      '// Iterator for pieces
    
    While i < UBound(mvarPictures)
        '// Set picture, and determine its final position
        Set mvarPictures(i).thePicture = PicClip.GraphicCell(i)
        mvarPictures(i).position = i + 1
        i = i + 1           '// Next piece
    Wend
End Function

Private Function ScrambleArray()
'// This function scrambles the puzzle pieces by
'// randomly swapping values in the array

    Dim aPiece As Piece     '// Temporary value; for swapping
    Dim i       As Integer  '// iterator for "for" loop
    Dim j       As Integer  '// iterator for "for" loop
    
    '// Use random number generator to swap values
    For i = 0 To 1000 - Int((Rnd * 500))
        
        For j = 0 To 25 - (Int(Rnd * 25)) - 1
            Dim rndVal  As Integer      '// Rendomly generated value
            rndVal = Int(Rnd * 25)      '// initialize value
            
            '// perform swapping
            aPiece = mvarPictures(j)
            mvarPictures(j) = mvarPictures(rndVal)
            mvarPictures(rndVal) = aPiece
        
        Next j
    Next i
    
End Function

Private Function DisplayArray()
'// This function displays the pictures/pieces of the
'// puzzle.  The position of the pictures is determined
'// by the mvarColumns and mvarRows (fixed = 5) and
'// pieces' width and height.

    Dim i               As Integer  '// iterator
    Dim j               As Integer  '// iterator
    Dim mvarClipWidth   As Long     '// Width of a piece
    Dim mvarClipHeight  As Long     '// height of a piece
    Dim x, y                        '// x & y positions of a piece

    '// Find values of the widht and height for a single piece
    mvarClipWidth = PieceWidth
    mvarClipHeight = PieceHeight
    
    i = 1
    y = 5   '// Start at 5 pixels from the top of frmMain
    x = 5   '// Start at 5 pixels from the left of frmMain
    
    '// Keep displaying pictures until all are displayed
    While i <= mvarColumns * mvarRows
        
        '// Set postions x and y and then assign picture
        picPiece.Item(i - 1).Top = y
        picPiece.Item(i - 1).Left = x
        picPiece.Item(i - 1).Picture = mvarPictures(i - 1).thePicture
        
        '// increment position x
        x = x + mvarClipWidth
        
        '// if there are 5 columns displayed
        If i Mod mvarColumns = 0 Then
            y = y + mvarClipHeight  '// Increment y
            x = 5                   '// start at first column
        End If
        
        i = i + 1   '// next piece
        
    Wend
    
End Function

Private Sub mnuHelpHelp_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub picPiece_Click(Index As Integer)
    '// perform swapping of the pieces if the first piece
    '// has already been selected; otherwise set clicked
    '// to first selected piece
    If sbStatus.Panels.Item(1).Text = "" Then
        sbStatus.Panels.Item(1).Text = Index
    Else
        swap sbStatus.Panels.Item(1).Text, Index
        numSwaps = numSwaps + 1
        sbStatus.Panels.Item(1).Text = ""
        sbStatus.Panels.Item(2).Text = "Number of swaps: " & numSwaps
        sbStatus.Panels.Item(3).Text = "Correct " & inPlace & " out of  25"
    End If
End Sub

Private Function swap(ByVal cell1 As Integer, ByVal cell2 As Integer)
'// This function performs the swapping of the two pictures/pieces
'// and determines whether all the pieces are in correct postion.
    On Error GoTo ErrHandler
    Dim temp  As Piece
    '// Swap
    temp = mvarPictures(cell1)
    mvarPictures(cell1) = mvarPictures(cell2)
    mvarPictures(cell2) = temp
    
    DisplayArray
    
    If inPlace = 25 Then
        Dim answer As Integer
        answer = MsgBox("Congratulations! You win!" & vbCrLf _
        & "Would you like to play again?", vbYesNo)
            If answer = vbYes Then
                mnuGameNew_Click
            Else
                mnuGameExit_Click
            End If
    End If
    
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Err.Clear
        mnuGameNew_Click
        Exit Function
    End If
End Function

Private Function inPlace() As Integer
'// This function determines how many pieces of the
'// puzzle are in in their place.
'// the function returns the number of the pieces in
'// correct position
    Dim i           As Integer
    Dim correct     As Integer
    
    While i < UBound(mvarPictures)
        If (mvarPictures(i).position = i + 1) Then correct = correct + 1
        i = i + 1
    Wend
    
    inPlace = correct
    
End Function
