VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type

Private mfEOF        As Boolean
Private mfFileOpen   As Boolean
Private mtFileName   As String
Private mlBufferSize As Long
Private mlFile       As Long
Private mlRet        As Long

Public Event FileNotFound(ptFileName As String)
Public Event FileIsAllreadyOpen()
Public Event FileIsAllreadyClosed()
Public Event FileCouldNotBeOpened()
Public Event FileCouldNotBeClosed()
Public Event FileIsNotOpen()
Public Event FileIsAtEOF()
Public Event IncorrectBufferSize(plBufferSize As Long)
Public Event FileNameUnchangedCouldNotCloseOpenFile()

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Function OpenFile() As Boolean
   Dim lSecAtt As SECURITY_ATTRIBUTES
   
   ' Standard result
   OpenFile = False
   
   ' Try to open file, if allready open, "raise error"
   If (mfFileOpen = True) Then
      RaiseEvent FileIsAllreadyOpen
   Else
      mfFileOpen = False
      
      If (IsWin2000 = True) Then
         ' Set for windows 2000, we need tho set
         ' some security attributes to open the file
         lSecAtt.nLength = Len(lSecAtt)   ' size of the structure
         lSecAtt.lpSecurityDescriptor = 0 ' default (normal) level of security
         lSecAtt.bInheritHandle = 1       ' this is the default setting
         
         mlFile = CreateFile(mtFileName, GENERIC_READ, FILE_SHARE_READ, ByVal lSecAtt, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
      Else
         mlFile = CreateFile(mtFileName, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
      End If
      
      If (mlFile <> -1) Then mfFileOpen = True
      
      ' File could not be opened, deleted ???
      If (mfFileOpen = False) Then
         RaiseEvent FileCouldNotBeOpened
      Else
         mfEOF = False
      End If

      ' Return true if the file was opened succesfully
      OpenFile = mfFileOpen
   End If
End Function

Public Function CloseFile() As Boolean
   ' Standard result
   CloseFile = False
   
   ' Only try to close file if the file is open
   If (mfFileOpen = True) Then
      ' Try to close file
      mlRet = CloseHandle(mlFile)
      
      ' If returned 1, then succesfull
      If (mlRet = 1) Then
         mfEOF = False
         CloseFile = True
         mfFileOpen = False
      Else
         RaiseEvent FileCouldNotBeClosed
      End If
   Else
      RaiseEvent FileIsAllreadyClosed
   End If
End Function

Public Function ReadBuffer() As String
   Dim lBytes  As Long
   Dim tBuffer As String
   
   If (mfFileOpen = True) Then
      If (mfEOF = False) Then
         tBuffer = String(mlBufferSize, Chr(0))
         mlRet = ReadFile(mlFile, ByVal tBuffer, mlBufferSize, lBytes, ByVal CLng(0))
         If lBytes = 0 Then
            ReadBuffer = ""
         Else
            ReadBuffer = Left(tBuffer, lBytes)
         End If
      
         mfEOF = (lBytes = 0)
      Else
         RaiseEvent FileIsAtEOF
      End If
   Else
      RaiseEvent FileIsNotOpen
   End If
End Function

Public Function ReadChars(mlCharsToRead As Long) As String
   Dim lBytes  As Long
   Dim lBuffer As Long
   Dim tBuffer As String
   
   lBuffer = mlCharsToRead
   If (lBuffer = 0) Then
      ReadChars = ""
      Exit Function
   Else
      If (mfFileOpen = True) Then
         If (mfEOF = False) Then
            tBuffer = String(lBuffer, Chr(0))
            mlRet = ReadFile(mlFile, ByVal tBuffer, lBuffer, lBytes, ByVal CLng(0))
            If lBytes = 0 Then
               ReadChars = ""
            Else
               ReadChars = Left(tBuffer, lBytes)
            End If

            mfEOF = (lBytes = 0)
         Else
            RaiseEvent FileIsAtEOF
         End If
      Else
         RaiseEvent FileIsNotOpen
      End If
   End If
End Function

Public Property Get FileName() As String
   FileName = mtFileName
End Property

Public Property Let FileName(ByVal ptFileName As String)
   ' Try to change filename, if fails, "raise error"
   If (FileExists(ptFileName) = True) Then
      
      ' if the file is open, close it
      If (mfFileOpen = True) Then
         
         ' if the file was closed succesfully,
         ' then set filename to new ptFileName
         If (CloseFile = True) Then
            mtFileName = ptFileName
         Else
            RaiseEvent FileNameUnchangedCouldNotCloseOpenFile
         End If
      
      Else
         mtFileName = ptFileName
      End If
   Else
      RaiseEvent FileNotFound(ptFileName)
   End If
End Property

Public Property Get BufferSize() As Long
   BufferSize = mlBufferSize
End Property

Public Property Let BufferSize(ByVal plBufferSize As Long)
   ' Try to change buffersize, if fails, "raise error"
   If (plBufferSize > 0) Then
      mlBufferSize = plBufferSize
   Else
      RaiseEvent IncorrectBufferSize(plBufferSize)
   End If
End Property

Public Property Get FileAtEOF() As Boolean
   FileAtEOF = mfEOF
End Property

Private Sub Class_Initialize()
   mfEOF = False
   mtFileName = ""
   mlBufferSize = 1048576
End Sub

Private Function IsWin2000() As Boolean
   Dim os   As OSVERSIONINFO
   Dim lRet As Long
   
   ' Default return
   IsWin2000 = False
   
   ' Check for windows 2000
   os.dwOSVersionInfoSize = Len(os)  ' set the size of the structure
   lRet = GetVersionEx(os)           ' read Windows's version information
   
   ' Check for Win32 & NT/2000
   If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      If (os.dwMajorVersion < 5) Then
         IsWin2000 = False
      Else
         IsWin2000 = True
      End If
   End If
End Function

Private Function FileExists(ptFileToCheck As String) As Boolean
    On Error Resume Next
    FileExists = False
    
    If (Dir(ptFileToCheck) <> "") Then FileExists = True
End Function

