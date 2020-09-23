VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crack Access 2000 Password"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSelect 
      Height          =   1230
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4560
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   675
         Width           =   3435
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Top             =   675
         Width           =   780
      End
      Begin VB.Label lblInstruction 
         Caption         =   "Click on the Access database you wish to recover the password for"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   6
         Top             =   225
         Width           =   4200
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2970
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraPassword 
      Height          =   2670
      Left            =   0
      TabIndex        =   0
      Top             =   1215
      Width           =   4560
      Begin VB.CommandButton cmdChangeTime 
         Caption         =   "Search by Changing the Creation Date"
         Height          =   285
         Left            =   270
         TabIndex        =   7
         ToolTipText     =   "Use this if the database's creation date has changed."
         Top             =   2295
         Width           =   3930
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   405
         Width           =   3885
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   $"frmMain.frx":030A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   180
         TabIndex        =   8
         Top             =   765
         Width           =   4320
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblResults 
         AutoSize        =   -1  'True
         Caption         =   "The password is:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   135
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmdBrowse_Click()
On Error GoTo TheError
    cdlg.Filter = "Access 2000 database(*.mdb)|*.mdb"

    cdlg.ShowOpen
    txtFileName.Text = cdlg.FileName
    txtPassword.Text = GuessAccess2000Password(cdlg.FileName)
    If txtPassword.Text = "" Then
        Call MsgBox("The database does not have a password.", vbInformation, App.EXEName)
    End If
    Exit Sub
TheError:
    Resume Next
End Sub

Private Sub cmdChangeTime_Click()
    If txtFileName.Text <> "" Then
        frmChangeTime.mFileStr = txtFileName.Text
        frmChangeTime.Show 1
    End If
End Sub


