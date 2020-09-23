VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChangeTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Scan"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   Icon            =   "frmChangeDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbarTime 
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   4860
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtDays 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2700
      TabIndex        =   3
      Text            =   "50"
      Top             =   135
      Width           =   555
   End
   Begin MSFlexGridLib.MSFlexGrid grdTime 
      Height          =   4110
      Left            =   0
      TabIndex        =   1
      Top             =   765
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   7250
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   330
      Left            =   3285
      TabIndex        =   0
      Top             =   135
      Width           =   960
   End
   Begin VB.Label lblResults 
      Height          =   330
      Left            =   90
      TabIndex        =   5
      Top             =   495
      Width           =   4110
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTime 
      Caption         =   "Enter the number of days to check"
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   180
      Width           =   2580
   End
End
Attribute VB_Name = "frmChangeTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mDate As Date
Public mFileStr As String


Private Sub cmdGo_Click()
Dim i As Integer
Dim dayCount As Integer
Dim password As String

Dim hFile As Long, rval As Long
Dim buff As OFSTRUCT
Dim ctime As FILETIME, mtime As FILETIME, latime As FILETIME
Dim stime As SYSTEMTIME
Dim filen As String
Dim c As String
Dim OK As Boolean
Dim check As Integer
    dayCount = Trim(txtDays.Text)
    If dayCount > 0 Then
        
        Screen.MousePointer = vbHourglass
        grdTime.Visible = False
        grdTime.Clear
        initGrdTime
        pbarTime.Visible = True
        pbarTime.Max = dayCount
        For i = 1 To dayCount
            OK = True
            pbarTime.Value = i
            mDate = Date - i
            filen = CStr(mFileStr)
            
            hFile = OpenFile(filen, buff, OF_WRITE)
            
            If hFile Then

            'get original file times
                rval = GetFileTime(hFile, ctime, latime, mtime)
                'convert system to file time
                rval = FileTimeToLocalFileTime(ctime, ctime)
                rval = FileTimeToSystemTime(ctime, stime)

                'Change filetimes
                stime.wYear = Year(mDate)
                stime.wMonth = Month(mDate)
                stime.wDay = Day(mDate)
                stime.wHour = Hour(Time)
                stime.wMinute = Minute(Time)
                stime.wSecond = Second(Time)
             '   MsgBox Day(mDate)
                rval = SystemTimeToFileTime(stime, ctime)
                rval = LocalFileTimeToFileTime(ctime, ctime)
                rval = SetFileTime(hFile, ctime, latime, mtime)
                rval = CloseHandle(hFile)
                password = GuessAccess2000Password(mFileStr)
                If i Mod 10 = 0 Then DoEvents
            End If
            For check = 1 To Len(password)
                c = Mid(password, check, 1)
                If Asc(c) < 32 Or Asc(c) > 127 Then
                    OK = False
                    Exit For
                End If
                
           Next
           If OK Then Call addToGrdTime(mDate, password)
        Next
        pbarTime.Visible = False
        grdTime.Visible = True
        Screen.MousePointer = vbDefault
        lblResults.Caption = "Results for " & mFileStr
    End If
End Sub
Private Sub addToGrdTime(d As Date, s As String)
    With grdTime
        .Rows = .Rows + 1
        .Col = 0
        .Row = .Rows - 1
        .Text = d
        .Col = 1
        .Text = s
    End With
End Sub
Private Sub Form_Load()
    mDate = Date
    initGrdTime
    pbarTime.Visible = False
    pbarTime.Left = grdTime.Left
    pbarTime.Width = grdTime.Width
    Me.Caption = "Date Scan - " & mFileStr
End Sub
Private Sub initGrdTime()
    With grdTime
        .Col = 0
        .Row = 0
        .ColWidth(0) = 1200
        .Text = "Creation Date"
        .Col = 1
        .Text = "Password"
        .ColWidth(1) = 3000
        .Rows = 1
    End With
End Sub
