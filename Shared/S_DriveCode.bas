Attribute VB_Name = "S_DriveCode"
Option Private Module



Option Explicit


Public Const GENERIC_READ As Long = &H80000000
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const OPEN_EXISTING As Long = 3
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const MAX_PATH As Long = 260


' =================================================================================
' ====================     GENERAL DECLARATIONS        =========================
' =================================================================================


'Enum containing values representing
'the status of the file
Public Enum IsFileResults
   FILE_IN_USE = -1  'True
   FILE_FREE = 0     'False
   FILE_DOESNT_EXIST = -999 'arbitrary number, other than 0 or -1
End Enum

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Declare Function CreateFile Lib "kernel32" _
   Alias "CreateFileA" _
  (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long
    

  
Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    
Public Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Public Declare Function EnumProcesses Lib "PSAPI.DLL" ( _
   lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long

Public Declare Function EnumProcessModules Lib "PSAPI.DLL" ( _
    ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long

Public Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" ( _
    ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_QUERY_INFORMATION = &H400




Const SW_SHOW = 1
Const SW_SHOWMAXIMIZED = 3


Public Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpszOp As String, _
    ByVal lpszFile As String, ByVal lpszParams As String, _
    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
    As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Const SW_SHOWNORMAL = 1

Private Declare Function GetDriveType Lib "kernel32" _
  Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Function DriveType(DriveLetter As String) As String
'  Returns a string that describes the type of drive of DriveLetter
   DriveLetter = Left(DriveLetter, 1) & ":\"
   Select Case GetDriveType(DriveLetter)
      Case 0: DriveType = "Unknown"
      Case 1: DriveType = "Non-existent"
      Case 2: DriveType = "Removable drive"
      Case 3: DriveType = "Fixed drive"
      Case 4: DriveType = "Network drive"
      Case 5: DriveType = "CD-ROM drive"
      Case 6: DriveType = "RAM disk"
      Case Else: DriveType = "Unknown drive type"
  End Select
End Function
' =================================================================================
' ====================     BIT TESTING STUFF         =========================
' =================================================================================


Sub BitSet(X As Long, ByVal n As Long, ByVal Value As Boolean)
        If Value Then
          X = X Or BitMask(n)
        Else
          X = X And Not BitMask(n)
        End If
End Sub


Function BitTest(X As Long, ByVal n As Long) As Boolean
      ' Return False if invalid N
        BitTest = (X And BitMask(n)) <> 0
End Function


Function BitMask(ByVal n As Long) As Long
      Dim i As Long, Mask As Long
        If n < 0 Or n > 31 Then
          BitMask = 0
        ElseIf n = 31 Then
          BitMask = &H80000000
        Else
          Mask = 1
          For i = 1 To n
            Mask = Mask + Mask
          Next i
          BitMask = Mask
        End If
End Function


' =================================================================================
' ====================     GET FILES / FOLDERS =========================
' =================================================================================

Function GetFolder(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

If Right(strPath, 1) <> "\" Then strPath = strPath & "\"

With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With

NextCode:
GetFolder = sItem
Set fldr = Nothing
    
End Function


Public Sub OpenFolder(GetIn As String)
Dim Result

Dim ProjPath

On Error Resume Next

ProjPath = GetIn
If Len(ProjPath) And DirExists(ProjPath) Then
    Shell "C:\WINDOWS\explorer.exe """ & ProjPath & "", vbNormalFocus
ElseIf Len(ProjPath) And DirExists(ProjPath) = False Then
    Result = MsgBox("Folder does not exist!", vbCritical, ProgramName$)
Else
    Result = MsgBox("No selected folder to open!", vbCritical, ProgramName$)
End If
End Sub

Function GetFile(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFilePicker)

If Right(strPath, 1) <> "\" Then strPath = strPath & "\"

With fldr
    .Title = "Select a File"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    .Filters.Clear
    .InitialView = msoFileDialogViewDetails
  '  If Not IsMissing(FileType) Then
        'Add the filter supplied through the function.
 '       .Filters.Add "Files Required", FileType
  '  End If
 '   FileType = ALLFILETYPES
  '  .Filters.Add ALLFILES, ALLFILETYPES
  
  
'        ElseIf LCase(FileType) = "access" Then
'            .Filters.Add "Access Database", "*.mdb,*.accdb", 1
'            .Filters.Add "Access Database 2007-2010", "*.accdb", 1
    
    
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With

NextCode:
GetFile = sItem
Set fldr = Nothing
    
End Function


' =================================================================================
' ====================     CHECK FILES / FOLDERS =========================
' =================================================================================


Function IsFileOpen(FileName As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open FileName For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            MsgBox "Error Accessing File"
        '    Error errnum
            Stop
    End Select

End Function

Public Function IsFileInUse(sFile As String) As IsFileResults

   Dim hFile As Long
   
   If FileExists(sFile) Then
   
     'note that FILE_ATTRIBUTE_NORMAL (&H80) has
     'a different value than VB's constant vbNormal (0)!
      hFile = CreateFile(sFile, _
                         GENERIC_READ, _
                         0, 0, _
                         OPEN_EXISTING, _
                         FILE_ATTRIBUTE_NORMAL, 0&)

     'this will evaluate to either
     '-1 (FILE_IN_USE) or 0 (FILE_FREE)
      IsFileInUse = hFile = INVALID_HANDLE_VALUE

      CloseHandle hFile
   
   Else
   
     'the value of FILE_DOESNT_EXIST in the Enum
     'is arbitrary, as long as it's not 0 or -1
      IsFileInUse = FILE_DOESNT_EXIST
   
   End If
   
End Function



Public Function DirExists(ByVal sDirName As String) As Boolean
On Error Resume Next
DirExists = (GetAttr(sDirName) And vbDirectory) = vbDirectory
Err.Clear
End Function


Public Function FileExists(ByVal sPathName As String) As Boolean
On Error Resume Next
FileExists = (GetAttr(sPathName) And vbNormal) = vbNormal
Err.Clear
End Function




Public Sub OpenDocument(FileName$)
Dim nDT, nApp

If FileExists(FileName$) = False Then Exit Sub
If FileName$ = "" Then Exit Sub
nDT = GetDesktopWindow()
nApp = ShellExecute(nDT, "Open", FileName$, "", "C:\", SW_SHOWNORMAL)

End Sub



'Public Function HardDelete(ByVal strFolderPath As String) As Boolean
'HardDelete = True'

'On Error GoTo fail:
' If DirExists(strFolderPath) = False Then Exit Function
'
' Dim fso As New Scripting.FileSystemObject
'   'You can also pass a second argument and set it to True if you want to force the deletion of read-only files:''
'
'   'fso.DeleteFolder strFolderPath
'    fso.DeleteFolder strFolderPath, True'
'
'Exit Function'
'
'fail:
'HardDelete = False
'
'End Function


Sub CreateLNKFile(ByVal sShortcut As String, ByVal sFileLinkName As String, Optional inIcon As String)

    If Right(LCase(sShortcut), 4) <> ".lnk" Then sShortcut = sShortcut & ".lnk"
    Dim sh As Object
    Dim link As Object
    
    Set sh = CreateObject("WScript.Shell")
    Set link = sh.CreateShortcut(sShortcut)
    link.TargetPath = sFileLinkName
    link.Description = "Shortcut for Project Folder"
    
    Select Case LCase(inIcon)
        Case "web"
        link.IconLocation = "C:\WINDOWS\system32\shell32.dll, 13"
        
        Case "folder"
        link.IconLocation = "C:\WINDOWS\system32\shell32.dll, 4"
        
        Case Else
        link.IconLocation = "C:\WINDOWS\system32\shell32.dll, 23"
    End Select
    
    link.Save
    
End Sub




