Attribute VB_Name = "frmComerclarUsu"
Option Explicit
 
'typedef enum _MEMORY_INFORMATION_CLASS {
'    MemoryBasicInformation,
'    MemoryWorkingSetList,
'    MemorySectionName
'} MEMORY_INFORMATION_CLASS;
 
Public Enum MEMORY_INFORMATION_CLASS
    MemoryBasicInformation = 0
    MemoryWorkingSetList
    MemorySectionName
End Enum
 
'typedef struct _MEMORY_BASIC_INFORMATION {
'    PVOID BaseAddress;
'    PVOID AllocationBase;
'    DWORD AllocationProtect;
'    SIZE_T RegionSize;
'    DWORD State;
'    DWORD Protect;
'    DWORD Type;
'} MEMORY_BASIC_INFORMATION, *PMEMORY_BASIC_INFORMATION;
 
Public Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    Type As Long
End Type
 
'typedef struct _FUNCTION_INFORMATION {
'    char name[64];
'    ULONG_PTR VirtualAddress;
'} FUNCTION_INFORMATION, *PFUNCTION_INFORMATION;
 
Public Type FUNCTION_INFORMATION
    name As String * 64
    VirtualAddress As Long
End Type
 
'typedef struct _MODULE_INFORMATION
'{
'    PVOID BaseAddress;
'    PVOID AllocationBase;
'    DWORD AllocationProtect;
'    SIZE_T RegionSize;
'    DWORD State;
'    DWORD Protect;
'    DWORD Type;
'    WCHAR szPathName[MAX_PATH];
'    PVOID EntryAddress;
'    PFUNCTION_INFORMATION Functions;
'    DWORD FunctionCount;
'    DWORD SizeOfImage;
'}MODULE_INFORMATION, *PMODULE_INFORMATION;
 
Public Type MODULE_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    Type As Long
    szPathName(1 To 520) As Byte
    EntryAddress As Long
    Functions As Long 'VarPtr(MODULE_INFORMATION), es un puntero, PFUNCTION_INFORMATION Functions;
    FunctionCount As Long
    SizeOfImage As Long
End Type
 
'struct UNICODE_STRING {
'    USHORT  Length;
'    USHORT  MaximumLength;
'    PWSTR    Buffer;
'};
 
Public Type UNICODE_STRING
    length As Integer
    MaximumLength As Integer
    Buffer As Long 'PWSTR    Buffer;
End Type
 
'typedef UNICODE_STRING *PUNICODE_STRING;
 
Public Const PAGE_NOACCESS = &H1
Public Const PAGE_READONLY = &H2
Public Const PAGE_READWRITE = &H4
Public Const PAGE_WRITECOPY = &H8
Public Const PAGE_EXECUTE = &H10
Public Const PAGE_EXECUTE_READ = &H20
Public Const PAGE_EXECUTE_READWRITE = &H40
Public Const PAGE_EXECUTE_WRITECOPY = &H80
Public Const PAGE_GUARD = &H100
Public Const PAGE_NOCACHE = &H200
Public Const PAGE_WRITECOMBINE = &H400
Public Const MEM_COMMIT = &H1000
Public Const MEM_RESERVE = &H2000
Public Const MEM_DECOMMIT = &H4000
Public Const MEM_RELEASE = &H8000
Public Const MEM_FREE = &H10000
Public Const MEM_PRIVATE = 20000
Public Const MEM_MAPPED = &H40000
Public Const MEM_RESET = &H80000
Public Const MEM_TOP_DOWN = &H100000
Public Const MEM_WRITE_WATCH = &H200000
Public Const MEM_PHYSICAL = &H400000
Public Const MEM_ROTATE = &H800000
Public Const MEM_LARGE_PAGES = &H20000000
Public Const MEM_4MB_PAGES = &H80000000
 
'typedef LONG (WINAPI *ZWQUERYVIRTUALMEMORY)(
'    HANDLE ProcessHandle,
'    PVOID BaseAddress,
'    MEMORY_INFORMATION_CLASS MemoryInformationClass,
'    PVOID MemoryInformation,
'    ULONG MemoryInformationLength,
'    PULONG ReturnLength
');
 
Public Declare Function ZwQueryVirtualMemory Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal BaseAddress As Long, ByVal MemoryInformationClass As MEMORY_INFORMATION_CLASS, ByVal MemoryInformation As Long, ByVal MemoryInformationLength As Long, ByVal ReturnLength As Long) As Long
 
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
 
Public Declare Function VirtualQuery Lib "kernel32" (ByRef lpAddress As Any, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
 
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (destination As Any, ByVal length As Long)
 
Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
 
 
 
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
 
Public Const PROCESS_ALL_ACCESS = &H1F0FFF  'Specifies all possible access flags for the process object.
Public Const PROCESS_CREATE_THREAD = &H2   'Enables using the process handle in the CreateRemoteThread function to create a thread in the process.
Public Const PROCESS_DUP_HANDLE = &H40  'Enables using the process handle as either the source or target process in the DuplicateHandle function to duplicate a handle.
Public Const PROCESS_QUERY_INFORMATION = &H400 'Enables using the process handle in the GetExitCodeProcess and GetPriorityClass functions to read information from the process object.
Public Const PROCESS_SET_INFORMATION = &H200  'Enables using the process handle in the SetPriorityClass function to set the priority class of the process.
Public Const PROCESS_TERMINATE = &H1 'Enables using the process handle in the TerminateProcess function to terminate the process.
Public Const PROCESS_VM_OPERATION = &H8 'Enables using the process handle in the VirtualProtectEx and WriteProcessMemory functions to modify the virtual memory of the process.
Public Const PROCESS_VM_READ = &H10     'Enables using the process handle in the ReadProcessMemory function to read from the virtual memory of the process.
Public Const PROCESS_VM_WRITE = &H20 'Enables using the process handle in the WriteProcessMemory function to write to the virtual memory of the process.
Public Const SYNCHRONIZE = &H100000   'Enables using the process handle in any of the wait functions to wait for the process to terminate.
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
 
'The WideCharToMultiByte function maps a wide-character string to a new character string.
'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.
 
'CodePage
Private Const CP_ACP = 0 'ANSI
Private Const CP_MACCP = 2 'Mac
Private Const CP_OEMCP = 1 'OEM
Private Const CP_UTF7 = 65000
Private Const CP_UTF8 = 65001
 
'dwFlags
Private Const WC_NO_BEST_FIT_CHARS = &H400
Private Const WC_COMPOSITECHECK = &H200
Private Const WC_DISCARDNS = &H10
Private Const WC_SEPCHARS = &H20 'Default
Private Const WC_DEFAULTCHAR = &H40
 
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
 
Public Function ByteArrayToString(Bytes() As Byte) As String
Dim iUnicode As Long, i As Long, j As Long
 
On Error Resume Next
i = UBound(Bytes)
 
If (i < 1) Then
    'ANSI, just convert to unicode and return
    ByteArrayToString = StrConv(Bytes, vbUnicode)
    Exit Function
End If
i = i + 1
 
'Examine the first two bytes
CopyMemory iUnicode, Bytes(0), 2
 
If iUnicode = Bytes(0) Then 'Unicode
    'Account for terminating null
    If (i Mod 2) Then i = i - 1
    'Set up a buffer to recieve the string
    ByteArrayToString = String$(i / 2, 0)
    'Copy to string
    CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
Else 'ANSI
    ByteArrayToString = StrConv(Bytes, vbUnicode)
End If
End Function
 
Public Function StringToByteArray(strInput As String, Optional bReturnAsUnicode As Boolean = True, Optional bAddNullTerminator As Boolean = False) As Byte()
Dim lRet As Long
Dim bytBuffer() As Byte
Dim lLenB As Long
 
If bReturnAsUnicode Then
    'Number of bytes
    lLenB = LenB(strInput)
    'Resize buffer, do we want terminating null?
    If bAddNullTerminator Then
        ReDim bytBuffer(lLenB)
    Else
        ReDim bytBuffer(lLenB - 1)
    End If
    'Copy characters from string to byte array
    CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
Else
    'METHOD ONE
'        'Get rid of embedded nulls
'        strRet = StrConv(strInput, vbFromUnicode)
'        lLenB = LenB(strRet)
'        If bAddNullTerminator Then
'            ReDim bytBuffer(lLenB)
'        Else
'            ReDim bytBuffer(lLenB - 1)
'        End If
'        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
 
    'METHOD TWO
    'Num of characters
    lLenB = Len(strInput)
    If bAddNullTerminator Then
        ReDim bytBuffer(lLenB)
    Else
        ReDim bytBuffer(lLenB - 1)
    End If
    lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
End If
 
StringToByteArray = bytBuffer
End Function
 
