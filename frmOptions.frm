VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Puzzle Options"
   ClientHeight    =   3576
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   4668
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3576
   ScaleWidth      =   4668
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   372
      Left            =   3240
      TabIndex        =   10
      Top             =   3120
      Width           =   1212
   End
   Begin VB.Frame Frame2 
      Caption         =   "Puzzle Information"
      Height          =   1572
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4452
      Begin VB.TextBox txtPieces 
         Height          =   288
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   852
      End
      Begin VB.TextBox txtRows 
         Height          =   288
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   852
      End
      Begin VB.TextBox txtColumns 
         Height          =   288
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label3 
         Caption         =   "Total number of puzzle pieces:"
         Height          =   252
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   2292
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Rows:"
         Height          =   252
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   2412
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Columns:"
         Height          =   252
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2652
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select Picture"
      Height          =   372
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   1332
   End
   Begin VB.Frame Frame1 
      Caption         =   "Puzzle Picture"
      Height          =   1212
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4452
      Begin VB.TextBox txtPicture 
         Height          =   288
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4212
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1800
      Top             =   1440
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    '// Class Constants
    Private Const COLUMNS = 10
    Private Const ROWS = 10
    
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    '// Set the public variables if the picture has
    '// been selected
    If txtPicture.Text <> "" Then
        mvarPicture = txtPicture.Text
        mvarRows = txtRows.Text
        mvarColumns = txtColumns.Text
        mvarPieces = txtPieces.Text
        mvarOptionsDefined = True
    Else
        MsgBox "Specify Puzzle Picture!"
        txtPicture.SetFocus
        Exit Sub
    End If
    
    fmain.mnuGameNew_Click
    cmdCancel_Click
    
End Sub

Private Sub cmdSelect_Click()
'// use the file dialog box to select a picture for
'// puzzle

    Dim sFile As String     '// Path to puzzle picture

    With dlgCommonDialog
        .DialogTitle = "Selecting PUZZLE picture..."
        .CancelError = False
        .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
        .Filter = "Window Bitmaps (*.bmp)|*.bmp"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    txtPicture.Text = sFile '// set txtPicture text to selected path
End Sub

Private Sub Form_Load()
    txtPicture.Text = mvarPicture
    txtColumns.Text = 5
    txtRows.Text = 5
    txtPieces.Text = txtColumns.Text * txtRows.Text
End Sub

