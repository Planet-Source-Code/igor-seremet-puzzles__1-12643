Attribute VB_Name = "Start"
Option Explicit

    '// Public type ( a single piece of the puzzle )
    Public Type Piece
        thePicture  As Variant  '// The picture Object
        position    As Integer  '// The position of the picture
    End Type
    
    '// Public Attributes
    Public fmain                As frmMain  '// Main form
    Public mvarColumns          As Integer  '// Number of columns of the puzzle - fixed
    Public mvarRows             As Integer  '// Number of rows of the puzzle - fixed
    Public mvarPieces           As Integer  '// Number of puzzle pieces
    Public mvarPicture          As String   '// Path to the picture
    Public mvarOptionsDefined   As Boolean  '// Are options set?
    Public mvarPictures(25)     As Piece    '// Array of pieces
    
Public Sub Main()
    Set fmain = New frmMain     '// Create an instance of frmMain
    fmain.Show                  '// Show form
End Sub
