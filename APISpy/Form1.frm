VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get All windows system api --------------by Qixinghua "
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "c:\AllWinAPI"
      Top             =   720
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "Save To"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type IMAGE_EXPORT_DIRECTORY
  Characteristics As Long ' 4
  TimeDateStamp As Long ' 4
  MajorVersion As Integer '2
  MinorVersion As Integer '2
  Name As Long '4
  Base As Long '4
  NumberOfFunctions As Long '4
  NumberOfNames As Long '4
  AddressOfFunctions As Long  ' 4 RVA from base of image
  AddressOfNames As Long  ' RVA from base of image
  AddressOfNameOrdinals As Long ' RVA from base of image
End Type

Private Type IMAGE_DATA_DIRECTORY
     VirtualAddress As Long
     Size As Long
End Type
Private Type IMAGE_OPTIONAL_HEADER
  Magic As Long
  MajorLinkerVersion  As Byte
  MinorLinkerVersion As Byte
  SizeOfCode As Long
  SizeOfInitializedData As Long
  SizeOfUninitializedData As Long
  AddressOfEntryPoint As Long
  BaseOfCode As Long
  BaseOfData As Long
  ImageBase As Long
  SectionAlignment As Long
  FileAlignment As Long
  MajorOperatingSystemVersion As Integer
  MinorOperatingSystemVersion As Integer
  MajorImageVersion As Integer
  MinorImageVersion As Integer
  MajorSubsystemVersion As Integer
  MinorSubsystemVersion As Integer
  Win32VersionValue As Long
  SizeOfImage As Long
  SizeOfHeaders As Long
  CheckSum As Long
  Subsystem As Integer
  DllCharacteristics As Integer
  SizeOfStackReserve As Long
  SizeOfStackCommit As Long
  SizeOfHeapReserve As Long
  SizeOfHeapCommit As Long
  LoaderFlags As Long
  NumberOfRvaAndSizes As Long
  DataDirectory As Long '(0 To 256) As IMAGE_DATA_DIRECTORY
End Type
Private Type IMAGE_FILE_HEADER
   Machine As Integer
   NumberOfSections As Integer
   TimeDateStamp As Long
   PointerToSymbolTable As Long
   NumberOfSymbols As Long
   SizeOfOptionalHeader As Integer
   Characteristics As Integer
End Type
Private Type IMAGE_NT_HEADERS
   Signature As Long
   FileHeader As IMAGE_FILE_HEADER
   OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ImageNtHeader Lib "imagehlp.dll" (ByVal pImageBase As Long) As Long
Private Declare Function ImageDirectoryEntryToData Lib "imagehlp.dll" (ByVal pImageBase As Long, ByVal MappedAsImage As Boolean, ByVal DirectoryEntry As Long, Size As Long) As Long
Private Declare Function ImageRvaToVa Lib "imagehlp.dll" (ByVal NtHeaders As Long, ByVal pImageBase As Long, ByVal Rva As Long, LastRvaSection As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Private Const IMAGE_DIRECTORY_ENTRY_EXPORT = 0 ' Export Directory
Private Const IMAGE_DIRECTORY_ENTRY_IMPORT = 1          ' Import Directory
Private Const IMAGE_DIRECTORY_ENTRY_RESOURCE = 2        ' Resource Directory
Private Const IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3       ' Exception Directory
Private Const IMAGE_DIRECTORY_ENTRY_SECURITY = 4        ' Security Directory
Private Const IMAGE_DIRECTORY_ENTRY_BASERELOC = 5       ' Base Relocation Table
Private Const IMAGE_DIRECTORY_ENTRY_DEBUG = 6           ' Debug Directory
Private Const GENERIC_READ = &H80000000
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Const OPEN_EXISTING = 3
Private Const SECTION_MAP_READ = &H4
Private Const FILE_MAP_READ = SECTION_MAP_READ
Private Const INVALID_HANDLE_VALUE = -1
Private Const PAGE_READONLY = &H2
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Src As Any, ByVal Size As Long) As Long
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal DestStr As String, ByVal pSrc As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const MAX_PATH As Long = 260
'GetDriveType return values
Private Const DRIVE_REMOVABLE As Long = 2
Private Const DRIVE_FIXED As Long = 3
Private Const DRIVE_REMOTE As Long = 4
Private Const DRIVE_CDROM As Long = 5
Private Const DRIVE_RAMDISK As Long = 6
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
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
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Function SpyDll(ByVal strDll As String) As String
    Dim Size As Long
    Dim BaseAddress As Long
    Dim pExportDirectory As Long
    Dim ExportDirectory As IMAGE_EXPORT_DIRECTORY
    Dim ExportNamePointerTable() As Long
    Dim strResult As String
    Dim FuncName As String * 1000
    Dim I As Long
    On Error GoTo Errhandler
    
    BaseAddress = LoadLibrary(strDll)
   
    If BaseAddress <= 0 Then
      Exit Function
     End If
    pExportDirectory = ImageDirectoryEntryToData(BaseAddress, True, IMAGE_DIRECTORY_ENTRY_EXPORT, Size)
    If pExportDirectory = 0 Then GoTo Errhandler
    MoveMemory ExportDirectory, ByVal pExportDirectory, Len(ExportDirectory)
    If ExportDirectory.NumberOfNames = 0 Then GoTo Errhandler
    If IsBadReadPtr(ByVal (BaseAddress + ExportDirectory.AddressOfNames), ByVal ExportDirectory.NumberOfNames * 4) = 1 Then GoTo Errhandler
    ReDim ExportNamePointerTable(ExportDirectory.NumberOfNames)
    MoveMemory ExportNamePointerTable(0), _
        ByVal (BaseAddress + ExportDirectory.AddressOfNames), _
        ByVal ExportDirectory.NumberOfNames * 4
    strResult = ""
    For I = 0 To ExportDirectory.NumberOfNames - 1
        PtrToStr FuncName, BaseAddress + ExportNamePointerTable(I)
        strResult = strResult & "<tr>" & "<td><font size='3' face='Verdana' color='#000000'>" & CStr(I + 1) & "</font></td>" & "<td><font size='3' face='Verdana' color='#000000'>" & vbCrLf & Left(FuncName, InStr(FuncName, vbNullChar) - 1) & vbCrLf & "</font></td></tr>" & vbCrLf
        Label2 = Left(FuncName, InStr(FuncName, vbNullChar) - 1)
        DoEvents: DoEvents
    Next
    SpyDll = strResult
    FreeLibrary BaseAddress
    Erase ExportNamePointerTable
  Exit Function
Errhandler:
    If BaseAddress Then Call FreeLibrary(BaseAddress)
     Erase ExportNamePointerTable
    SpyDll = ""
    Exit Function
End Function
Private Function GetSystemDir() As String
Dim SystemDir As String * 255
GetSystemDirectory SystemDir, 255
GetSystemDir = Left$(SystemDir, InStr(SystemDir, vbNullChar) - 1)
End Function
Public Function ExportAPI(ByVal strDesDir As String, Optional strAPIHome As String)
Dim WFD As WIN32_FIND_DATA
Dim hFile As Long
Dim strPath As String
If Len(Trim(strAPIHome)) = 0 Then
   strPath = GetSystemDir
Else
   strPath = Trim(strAPIHome)
End If
If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
If Right$(strDesDir, 1) <> "\" Then strDesDir = strDesDir & "\"
Call MakeSureDirectoryPathExists(strDesDir)
hFile = FindFirstFile(strPath & "*.dll" & Chr$(0), WFD)
If hFile <> -1 Then
    sFile = TrimNull(WFD.cFileName)
    strTemp$ = SpyDll(strPath & sFile)
     If Len(strTemp$) > 0 Then
         Call SaveMe(strDesDir & sFile & ".html", strTemp$, sFile)
         strTemp$ = ""
      End If
 While FindNextFile(hFile, WFD)
      sFile = TrimNull(WFD.cFileName)
      DoEvents: DoEvents
      strTemp$ = SpyDll(strPath & sFile)
     If Len(strTemp$) > 0 Then
         Call SaveMe(strDesDir & sFile & ".html", strTemp$, sFile)
         strTemp$ = ""
     End If
 Wend
End If
  Call FindClose(hFile)
End Function
Private Function SaveMe(ByVal strDesFile As String, ByVal strContent As String, ByVal strTitle As String)
Dim fWriteHandle As Long
Dim fSuccess As Long
Dim lBytesWritten As Long
 fWriteHandle = CreateFile(strDesFile, GENERIC_WRITE Or GENERIC_READ, 0&, 0&, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  
   If fWriteHandle <> INVALID_HANDLE_VALUE Then
     strHeader$ = strHeader$ & "<html><Body><Center>" & vbCrLf
     strHeader$ = strHeader$ & "<font  face='Verdana' color='red' size='5'> All APIs in " & strTitle & "</font><br></Center><br>" & vbCrLf
     strHeader$ = strHeader$ & "<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0 WIDTH=100%>" & vbCrLf
     strHeader$ = strHeader$ & "<TD ALIGN=right VALIGN='top'><font  face='Verdana' color='Blue' size='2'> Author:Qi Xinghua</font><br>" & vbCrLf
     strHeader$ = strHeader$ & "<font  face='Verdana' color='Blue' size='2'> Almar Mater of Author:<a href=http://www.nju.edu.cn/>" & vbCrLf
     strHeader$ = strHeader$ & "Nanjing University</A></font><br>" & vbCrLf
     strHeader$ = strHeader$ & "<font face='Verdana' color='Blue' size='2'> Have questions?<a href=mailto:PRC_NJU_Qixh@Yahoo.com.cn>Email Me Now</A></font><br></Td></table>" & vbCrLf
     strHeader$ = strHeader$ & "<META HTTP-EQUIV='Content-Type' CONTENT='text/html';charset=ISO-8859-1>" & vbCrLf
     strHeader$ = strHeader$ & "<table   border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse'  bordercolor='#CCCCCC' width='100%'  bgcolor='#FFFFFF'>" & vbCrLf
     Call WriteFile(fWriteHandle, ByVal strHeader$, Len(strHeader$), lBytesWritten, 0)
     fSuccess = FlushFileBuffers(fWriteHandle)
     fSuccess = WriteFile(fWriteHandle, ByVal strContent, Len(strContent), lBytesWritten, 0)
     If fSuccess <> 0 Then
        fSuccess = FlushFileBuffers(fWriteHandle)
     End If
     strEnd$ = "</Body></html>" & vbCrLf
     Call WriteFile(fWriteHandle, ByVal strEnd$, Len(strEnd$), lBytesWritten, 0)
     fSuccess = FlushFileBuffers(fWriteHandle)
     fSuccess = CloseHandle(fWriteHandle)
     End If
End Function
Private Function TrimNull(startstr As String) As String
  Dim pos As Long
  pos = InStr(startstr, Chr$(0))
  If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
  End If
  TrimNull = startstr
End Function


Private Sub Command1_Click()
ExportAPI Trim(Text1)
End Sub
