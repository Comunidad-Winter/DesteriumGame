Attribute VB_Name = "ModPersonajes"
Option Explicit
    
    
    
    Public Const TH32CS_SNAPPROCESS As Long = &H2
    Public Const MAX_PATH As Integer = 260
    
    Public Type PROCESSENTRY32
        dwSize As Long
        cntUsage As Long
        th32ProcessID As Long
        th32DefaultHeapID As Long
        th32ModuleID As Long
        cntThreads As Long
        th32ParentProcessID As Long
        pcPriClassBase As Long
        dwFlags As Long
        szExeFile As String * MAX_PATH
    End Type
    
    Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
        "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) _
        As Long
    
    Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
        (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    
    Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
        (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    
    Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

