VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Puzzles"
   ClientHeight    =   3384
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   4368
   FillStyle       =   0  'Solid
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3384
   ScaleWidth      =   4368
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   3600
      Top             =   2280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   3132
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3972
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    '// Class Attributes
    Private msg As String   '// Holds text to be displayed
    
Private Sub Form_Load()
    '// Initialize the msg variable
    msg = "PROGRAM INFORMATION:" & Chr(10) _
    & "Name:    " & "Puzzles" & Chr(10) _
    & "Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(10) & Chr(10) _
    & "PROGRAMMER INFORMATION:" & Chr(10) _
    & "Name:    Igor Seremet" & Chr(10) _
    & "E-mail:  iseremet@yahoo.com" & Chr(10) & Chr(10) & Chr(10) _
    & "WARNING" & Chr(10) _
    & "The author of this program will not be held responsible for any " _
    & "problems that may arise due to use of Puzzles.exe or any supporting attributes."
End Sub


Private Sub Timer1_Timer()
    Static x    As Integer      '// Position in text
    x = x + 1                   '// Move to next character
    Label1.Caption = Left(msg, x)   '// Display new text as caption of label
    
    '// Disable the timer if all text chars have been displayed
    If x = Len(msg) Then Timer1.Enabled = False
End Sub
