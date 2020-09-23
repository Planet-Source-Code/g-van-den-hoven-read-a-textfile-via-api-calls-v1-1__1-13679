Attribute VB_Name = "modMain"
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

Public mlTimeTook  As Long
Public mlTimeStart As Long
Public mlTimeEnd   As Long
Public mlSpeed     As Double

Private Const BUFFER_SIZE_256k = 262144   ' 256k buffer (ehhh 256 * 1024)
Private Const BUFFER_SIZE_1Mb = 1048576   ' 1 MegaByte buffer (ehhh 1024 * 1024)
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Function ReadTextFile(ptFileName As String, pfSucces As Boolean) As String
   Dim lFile   As Long
   Dim lRet    As Long
   Dim lBytes  As Long
   Dim lSecAtt As SECURITY_ATTRIBUTES
   Dim tBuffer As String
   Dim tDestin As String
   tDestin = ""
   
   mlTimeStart = GetTickCount
   
   
   ' Try to open the file
   If IsWin2000 Then
      ' Set for windows 2000, we need tho set
      ' some security attributes to open the file
      lSecAtt.nLength = Len(lSecAtt)   ' size of the structure
      lSecAtt.lpSecurityDescriptor = 0 ' default (normal) level of security
      lSecAtt.bInheritHandle = 1       ' this is the default setting
      
      lFile = CreateFile(ptFileName, GENERIC_READ, FILE_SHARE_READ, ByVal lSecAtt, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
   Else
      lFile = CreateFile(ptFileName, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
   End If
   
   ' the file could not be opened
   If lFile = -1 Then
      ReadTextFile = ""
      pfSucces = False
      Exit Function
   End If
   
   tBuffer = String(BUFFER_SIZE_1Mb, Chr(0))
   lRet = ReadFile(lFile, ByVal tBuffer, BUFFER_SIZE_1Mb, lBytes, ByVal CLng(0))
   ' Check for EOF
   If lBytes = 0 Then
      ReadTextFile = ""
      pfSucces = False
   Else
      ' Not EOF, read the rest of the file...
      ReadTextFile = ""
      tDestin = Left(tBuffer, lBytes)
      Do While lBytes > 0
         lRet = ReadFile(lFile, ByVal tBuffer, BUFFER_SIZE_1Mb, lBytes, ByVal CLng(0))
         tDestin = tDestin & Left(tBuffer, lBytes)
      Loop
   End If
   
   ' Close the file
   lRet = CloseHandle(lFile)

   ' Calculate the time that it took
   mlTimeEnd = GetTickCount
   mlTimeTook = mlTimeEnd - mlTimeStart
   On Error Resume Next
   mlSpeed = (Len(tDestin) / mlTimeTook) / 1024
   Err.Clear

   ' Return contents
   ReadTextFile = tDestin
   pfSucces = True

   ' Clear memory
   tDestin = ""
End Function

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
