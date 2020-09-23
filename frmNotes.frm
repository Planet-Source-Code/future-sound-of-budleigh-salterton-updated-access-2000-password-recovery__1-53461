VERSION 5.00
Begin VB.Form frmNotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notes"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblNotes 
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   4830
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim s As String
s = "If every other letter is returned or plainly wrong characters, try this:" & vbCrLf
s = s & "Set your system time to te date the database was created and make a copy."
s = s & "Run this program on the copy.  If that does not work, then set the date to"
s = s & "the last modified date of the original and try again."
lblNotes.Caption = s
End Sub
