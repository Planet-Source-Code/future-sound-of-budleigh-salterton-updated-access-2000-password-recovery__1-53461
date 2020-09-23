Attribute VB_Name = "basCrack"
Option Explicit
'Thanks to the person who figured this code out.
Dim NullDate As Date
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
   ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
   ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const OFS_MAXPATHNAME = 128
Private Const OF_READ = &H0

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As _
   BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As _
   SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As _
   FILETIME) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, _
   ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Function GetFileDate(File As String) As Date
Dim fhi As BY_HANDLE_FILE_INFORMATION
Dim ctime As FILETIME, atime As FILETIME, wtime As FILETIME
Dim ftime As SYSTEMTIME
Dim buff As OFSTRUCT
Dim rval As Long, hFile As Long
    hFile = OpenFile(File, buff, OF_READ)
    If hFile = -1 Then
        GetFileDate = NullDate
    Else
        GetFileInformationByHandle hFile, fhi
        ctime = fhi.ftCreationTime
        'Convert File Time Zone to Local
        rval = FileTimeToLocalFileTime(ctime, ctime)
        'Convert File Time Format to System Time Format
        rval = FileTimeToSystemTime(ctime, ftime)
        GetFileDate = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
    End If
    CloseHandle hFile
End Function

Public Function GuessAccess2000Password(ProtectedFile As String) As String
' The trick of this function is that Access 2000 codifies your password by
' making the XOR of it and a mask. I don't know the algorithm to get this mask
' but I do know that it only depends on the creation date of the database.
' So the proccess is: create a dummy database with the same date as the protected
' database, and then make the XOR of the encrypted bytes of the protected database
' and the password-less database we create in this process.

Dim n As Long, s1 As String * 1, s2 As String * 1
Dim Password As String
Dim x1 As Byte, x2 As Byte
Dim TempFile As String
Dim DateFile As Date, PreviousDate As Date
Dim Handle1 As Long, Handle2 As Long

    ' Get the creation date of the protected database
    DateFile = GetFileDate(ProtectedFile)
    If DateFile = NullDate Then
        GuessAccess2000Password = "Can't open database file. Maybe you have it open in exclusive mode"
        Exit Function
    End If
    ' Create a temporary file in the Temp directory of windows
    TempFile = TempPath & "temp.mdb"
    ' Remove the temporary file if it exists
    If Dir(TempFile) <> "" Then
        Kill TempFile
    End If

    ' Keep the system date, then set it to the same date as the protected database
    PreviousDate = Date
    Date = DateFile
 
    ' Create the database, which will have so the same date as the protected database
    CreateDatabase TempFile, dbLangGeneral

    ' We can restore now the real date
   Date = PreviousDate
    
    Handle1 = FreeFile
    Open TempFile For Binary As #Handle1
    Handle2 = FreeFile
    Open ProtectedFile For Binary As #Handle2
    Password = ""
    Seek #Handle1, &H43
    Seek #Handle2, &H43
    ' The maximum length of the password is 20 characters
    For n = 0 To 19
        x1 = Asc(Input(1, Handle1))
        x2 = Asc(Input(1, Handle2))
        If x1 <> x2 Then
          Password = Password & Chr(x1 Xor x2)
        End If
        ' We skip the even positions, because the password is stored using
        ' two bytes per character
        x1 = Asc(Input(1, Handle1))
        x2 = Asc(Input(1, Handle2))
    Next
    Close 1
    Close 2
    Kill TempFile
    GuessAccess2000Password = Password
End Function

Private Function TempPath() As String
' Generate a temporary file (path)\api????.TMP, where (path)
' is Windows's temporary file directory and ???? is a randomly assigned unique value.
' Then display the name of the created file on the screen.
Dim Path As String  ' receives name of temporary file path
Dim TempFile As String  ' receives name of temporary file
Dim slength As Long  ' receives length of string returned for the path
Dim lastfour As Long  ' receives hex value of the randomly assigned ????
    ' Get Windows's temporary file path
    Path = Space(255)  ' initialize the buffer to receive the path
    slength = GetTempPath(255, Path)  ' read the path name
    TempPath = Left(Path, slength)  ' extract data from the variable
End Function

Sub main()
' NOte 1: Put a reference to Microsoft DAO 3.6
' Note 2: You cannot have the database opened in exclusive mode while you execute this procedure
    'MsgBox "The password is: " & GuessAccess2000Password("c:\db1.mdb")

End Sub
