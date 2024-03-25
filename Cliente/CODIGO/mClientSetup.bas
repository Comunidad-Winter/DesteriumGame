Attribute VB_Name = "mClientSetup"
Option Explicit

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Private Declare Sub ZeroMemory _
                Lib "kernel32" _
                Alias "RtlZeroMemory" (destination As Any, _
                                       ByVal Length As Long)

Private Declare Function CryptHashData _
                Lib "advapi32.dll" (ByVal hHash As Long, _
                                    pbData As Any, _
                                    ByVal dwDataLen As Long, _
                                    ByVal dwFlags As Long) As Long

Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long

Private Declare Function CryptCreateHash _
                Lib "advapi32.dll" (ByVal hProv As Long, _
                                    ByVal Algid As Long, _
                                    ByVal hKey As Long, _
                                    ByVal dwFlags As Long, _
                                    phHash As Long) As Long

Private Declare Function CryptGetHashParam _
                Lib "advapi32.dll" (ByVal hHash As Long, _
                                    ByVal dwParam As Long, _
                                    pbData As Any, _
                                    pdwDataLen As Long, _
                                    ByVal dwFlags As Long) As Long

Private Declare Function CryptAcquireContext _
                Lib "advapi32.dll" _
                Alias "CryptAcquireContextA" (phProv As Long, _
                                              ByVal pszContainer As Long, _
                                              ByVal pszProvider As Long, _
                                              ByVal dwProvType As Long, _
                                              ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext _
                Lib "advapi32.dll" (ByVal hProv As Long, _
                                    ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL       As Long = 1

Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const CALG_MD5            As Long = 32771

Private hashInitialized           As Boolean

Private savedHashValue            As String

Private Function CalculateHash(ByVal data As String) As String

    Dim hCryptProv     As Long

    Dim hHash          As Long

    Dim result         As Long

    Dim hashData()     As Byte

    Dim hashDataLength As Long
    
    result = CryptAcquireContext(hCryptProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)

    If result = 0 Then
        ' Manejar error al adquirir el contexto criptográfico
        ' ...
        Exit Function

    End If
    
    result = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)

    If result = 0 Then
        ' Manejar error al crear el hash
        ' ...
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If
    
    result = CryptHashData(hHash, ByVal StrPtr(data), Len(data), 0)

    If result = 0 Then
        ' Manejar error al calcular el hash
        ' ...
        CryptDestroyHash hHash
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If
    
    result = CryptGetHashParam(hHash, 21, 0, hashDataLength, 0)

    If result = 0 Then
        ' Manejar error al obtener la longitud del hash
        ' ...
        CryptDestroyHash hHash
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If
    
    ReDim hashData(hashDataLength - 1) As Byte
    
    result = CryptGetHashParam(hHash, 21, hashData(0), hashDataLength, 0)

    If result = 0 Then
        ' Manejar error al obtener el hash
        ' ...
        CryptDestroyHash hHash
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If

    CalculateHash = StrConv(hashData, vbUnicode)

    CryptDestroyHash hHash
    CryptReleaseContext hCryptProv, 0

End Function

Private Function GetMemoryHash() As String

    Dim processId  As Long

    Dim memoryData As String

    processId = GetCurrentProcessId()

    ' Leer los datos de memoria relevantes
    ' Esto puede incluir valores críticos, estructuras de datos o secciones específicas de la memoria
    ' En este ejemplo, solo se muestra un valor de ejemplo
    memoryData = CStr(processId) & "example_data"

    GetMemoryHash = CalculateHash(memoryData)

End Function

Private Sub CheckMemory()

    Dim currentHashValue As String

    currentHashValue = GetMemoryHash()

    If hashInitialized Then
        If currentHashValue <> savedHashValue Then
            MsgBox "Modificación de memoria detectada"

        End If

    Else
        savedHashValue = currentHashValue
        hashInitialized = True

    End If

End Sub

Public Function PATH_CLIENTSETUP() As String
    
    PATH_CLIENTSETUP = App.path & "\AO\resource\init\config.ini"
        
End Function

Public Sub ILoadClientSetup()

    '<EhHeader>
    On Error GoTo ILoadClientSetup_Err

    '</EhHeader>
 
    Dim A As Long
    
    ' Start Cursor
    Call StartAnimatedCursor(App.path & "\AO\resource\cursor\" & ClientSetup.CursorGeneral, IDC_ARROW)
    Call StartAnimatedCursor(App.path & "\AO\resource\cursor\" & ClientSetup.CursorSpell, IDC_CROSS)
    Call StartAnimatedCursor(App.path & "\AO\resource\cursor\" & ClientSetup.CursorHand, IDC_HAND)
    
    If FileExist(PATH_CLIENTSETUP, vbArchive) Then

        ClientSetup.CursorGeneral = GetVar(PATH_CLIENTSETUP, "CURSOR", "GENERAL")
        ClientSetup.CursorHand = GetVar(PATH_CLIENTSETUP, "CURSOR", "HAND")
        ClientSetup.CursorInv = GetVar(PATH_CLIENTSETUP, "CURSOR", "INV")
        ClientSetup.CursorSpell = GetVar(PATH_CLIENTSETUP, "CURSOR", "SPELL")

        ClientSetup.bMasterSound = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "MASTER"))
        ClientSetup.bSoundMusic = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "MUSIC"))
        ClientSetup.bSoundEffect = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "EFFECT"))
        ClientSetup.bSoundInterface = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "INTERFACE"))

        ClientSetup.bValueSoundMusic = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEMUSIC"))
        ClientSetup.bValueSoundEffect = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEEFFECT"))
        ClientSetup.bValueSoundInterface = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEINTERFACE"))
        ClientSetup.bValueSoundMaster = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEMASTER"))
                  
        ClientSetup.bResolution = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "RESOLUTION"))
                  
        ClientSetup.bFps = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "FPS"))
        ClientSetup.bAlpha = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "ALPHA"))
        ClientSetup.bResolution = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "RESOLUTION"))
                   
        For A = 1 To MAX_SETUP_MODS
            ClientSetup.bConfig(A) = Val(GetVar(PATH_CLIENTSETUP, "CONFIG", CStr(A)))
        Next
                
    Else
        Call MsgBox("Hubo un error crítico al cargar las opciones de juego. Contacta a los Administradores del Juego", vbCritical, App.Title)

        'End
    End If

    '<EhFooter>
    Exit Sub

ILoadClientSetup_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.mClientSetup.ILoadClientSetup " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub
    
Public Sub ISaveClientSetup()

    '<EhHeader>
    On Error GoTo ISaveClientSetup_Err

    '</EhHeader>
     
    Dim A As Long
     
    Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "GENERAL", ClientSetup.CursorGeneral)
    Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "HAND", ClientSetup.CursorHand)
    Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "INV", ClientSetup.CursorInv)
    Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "SPELL", ClientSetup.CursorSpell)

    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "MASTER", CStr(ClientSetup.bMasterSound))
    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "MUSIC", CStr(ClientSetup.bSoundMusic))
    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "EFFECT", CStr(ClientSetup.bSoundEffect))
    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "INTERFACE", CStr(ClientSetup.bSoundInterface))

    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEMUSIC", CStr(ClientSetup.bValueSoundMusic))
    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEEFFECT", CStr(ClientSetup.bValueSoundEffect))
    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEINTERFACE", CStr(ClientSetup.bValueSoundInterface))
    Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEMASTER", CStr(ClientSetup.bValueSoundMaster))
          
    Call WriteVar(PATH_CLIENTSETUP, "VIDEO", "FPS", CStr(ClientSetup.bFps))
    Call WriteVar(PATH_CLIENTSETUP, "VIDEO", "ALPHA", CStr(ClientSetup.bAlpha))
    Call WriteVar(PATH_CLIENTSETUP, "VIDEO", "RESOLUTION", CStr(ClientSetup.bResolution))
          
    For A = 1 To MAX_SETUP_MODS
        Call WriteVar(PATH_CLIENTSETUP, "CONFIG", CStr(A), CStr(ClientSetup.bConfig(A)))
    Next

    '<EhFooter>
    Exit Sub

ISaveClientSetup_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.mClientSetup.ISaveClientSetup " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

