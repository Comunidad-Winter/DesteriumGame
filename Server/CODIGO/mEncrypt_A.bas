Attribute VB_Name = "mEncrypt_A"
'--- mdAesCbc.bas
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- for CryptAcquireContext
Private Const PROV_RSA_AES                  As Long = 24

Private Const CRYPT_VERIFYCONTEXT           As Long = &HF0000000

'--- for CryptCreateHash
Private Const CALG_RC2                      As Long = &H6602&

Private Const CALG_AES_128                  As Long = &H660E&

Private Const CALG_AES_192                  As Long = &H660F&

Private Const CALG_AES_256                  As Long = &H6610&

Private Const CALG_HMAC                     As Long = &H8009&

Private Const CALG_SHA1                     As Long = &H8004&

Private Const CALG_SHA_256                  As Long = &H800C&

Private Const CALG_SHA_384                  As Long = &H800D&

Private Const CALG_SHA_512                  As Long = &H800E&

'--- for CryptGet/SetHashParam
Private Const HP_HASHVAL                    As Long = 2

Private Const HP_HMAC_INFO                  As Long = 5

'--- for CryptImportKey
Private Const PLAINTEXTKEYBLOB              As Long = 8

Private Const CUR_BLOB_VERSION              As Long = 2

Private Const CRYPT_IPSEC_HMAC_KEY          As Long = &H100

'--- for CryptSetKeyParam
Private Const KP_IV                         As Long = 1

Private Const KP_MODE                       As Long = 4

Private Const CRYPT_MODE_CBC                As Long = 1

'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1

'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000

Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Const LNG_FACILITY_WIN32            As Long = &H80070000

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal length As Long)

Private Declare Function htonl Lib "ws2_32" (ByVal hostlong As Long) As Long

'--- advapi32
Private Declare Function CryptAcquireContext _
                Lib "advapi32" _
                Alias "CryptAcquireContextW" (phProv As Long, _
                                              ByVal pszContainer As Long, _
                                              ByVal pszProvider As Long, _
                                              ByVal dwProvType As Long, _
                                              ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext _
                Lib "advapi32" (ByVal hProv As Long, _
                                ByVal dwFlags As Long) As Long

Private Declare Function CryptImportKey _
                Lib "advapi32" (ByVal hProv As Long, _
                                pbData As Any, _
                                ByVal dwDataLen As Long, _
                                ByVal hPubKey As Long, _
                                ByVal dwFlags As Long, _
                                phKey As Long) As Long

Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function CryptGetHashParam _
                Lib "advapi32" (ByVal hHash As Long, _
                                ByVal dwParam As Long, _
                                pbData As Any, _
                                pdwDataLen As Long, _
                                ByVal dwFlags As Long) As Long

Private Declare Function CryptSetHashParam _
                Lib "advapi32" (ByVal hHash As Long, _
                                ByVal dwParam As Long, _
                                pbData As Any, _
                                ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash _
                Lib "advapi32" (ByVal hProv As Long, _
                                ByVal AlgId As Long, _
                                ByVal hKey As Long, _
                                ByVal dwFlags As Long, _
                                phHash As Long) As Long

Private Declare Function CryptHashData _
                Lib "advapi32" (ByVal hHash As Long, _
                                pbData As Any, _
                                ByVal dwDataLen As Long, _
                                ByVal dwFlags As Long) As Long

Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long

Private Declare Function CryptSetKeyParam _
                Lib "advapi32" (ByVal hKey As Long, _
                                ByVal dwParam As Long, _
                                pbData As Any, _
                                ByVal dwFlags As Long) As Long

Private Declare Function CryptEncrypt _
                Lib "advapi32" (ByVal hKey As Long, _
                                ByVal hHash As Long, _
                                ByVal Final As Long, _
                                ByVal dwFlags As Long, _
                                pbData As Any, _
                                pdwDataLen As Long, _
                                ByVal dwBufLen As Long) As Long

Private Declare Function CryptDecrypt _
                Lib "advapi32" (ByVal hKey As Long, _
                                ByVal hHash As Long, _
                                ByVal Final As Long, _
                                ByVal dwFlags As Long, _
                                pbData As Any, _
                                pdwDataLen As Long) As Long

Private Declare Function RtlGenRandom _
                Lib "advapi32" _
                Alias "SystemFunction036" (RandomBuffer As Any, _
                                           ByVal RandomBufferLength As Long) As Long
#If Not ImplUseShared Then

    Private Declare Function CryptStringToBinary _
                    Lib "crypt32" _
                    Alias "CryptStringToBinaryW" (ByVal pszString As Long, _
                                                  ByVal cchString As Long, _
                                                  ByVal dwFlags As Long, _
                                                  ByVal pbBinary As Long, _
                                                  ByRef pcbBinary As Long, _
                                                  ByRef pdwSkip As Long, _
                                                  ByRef pdwFlags As Long) As Long

    Private Declare Function CryptBinaryToString _
                    Lib "crypt32" _
                    Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, _
                                                  ByVal cbBinary As Long, _
                                                  ByVal dwFlags As Long, _
                                                  ByVal pszString As Long, _
                                                  pcchString As Long) As Long

    Private Declare Function WideCharToMultiByte _
                    Lib "kernel32" (ByVal CodePage As Long, _
                                    ByVal dwFlags As Long, _
                                    ByVal lpWideCharStr As Long, _
                                    ByVal cchWideChar As Long, _
                                    lpMultiByteStr As Any, _
                                    ByVal cchMultiByte As Long, _
                                    ByVal lpDefaultChar As Long, _
                                    ByVal lpUsedDefaultChar As Long) As Long

    Private Declare Function MultiByteToWideChar _
                    Lib "kernel32" (ByVal CodePage As Long, _
                                    ByVal dwFlags As Long, _
                                    lpMultiByteStr As Any, _
                                    ByVal cchMultiByte As Long, _
                                    ByVal lpWideCharStr As Long, _
                                    ByVal cchWideChar As Long) As Long

    Private Declare Function FormatMessage _
                    Lib "kernel32" _
                    Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                            lpSource As Long, _
                                            ByVal dwMessageId As Long, _
                                            ByVal dwLanguageId As Long, _
                                            ByVal lpBuffer As String, _
                                            ByVal nSize As Long, _
                                            Args As Any) As Long
#End If

Private Type BLOBHEADER

    bType               As Byte
    bVersion            As Byte
    reserved            As Integer
    aiKeyAlg            As Long
    cbKeySize           As Long
    buffer(0 To 255)    As Byte

End Type

Private Const sizeof_BLOBHEADER As Long = 12

Private Type HMAC_INFO

    HashAlgid           As Long
    pbInnerString       As Long
    cbInnerString       As Long
    pbOuterString       As Long
    cbOuterString       As Long

End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const AES_BLOCK_SIZE   As Long = 16

Private Const AES_KEYLEN       As Long = 32                    '-- 32 -> AES-256, 24 -> AES-196, 16 -> AES-128

Private Const AES_IVLEN        As Long = AES_BLOCK_SIZE

Private Const KDF_SALTLEN      As Long = 8

Private Const KDF_ITER         As Long = 10000

Private Const KDF_HASH         As String = "SHA512"

Private Const OPENSSL_MAGIC    As String = "Salted__"          '-- for openssl compatibility

Private Const OPENSSL_MAGICLEN As Long = 8

Private Type UcsCryptoContextType

    hProv               As Long
    hKey                As Long

End Type

'=========================================================================
' Functions
'=========================================================================

'--- equivalent to `openssl aes256 -pbkdf2 -md sha512 -pass pass:{sPassword} -in {sText}.file -a`
Public Function AesEncryptString(sText As String, sPassword As String) As String

    Const PREFIXLEN                  As Long = OPENSSL_MAGICLEN + KDF_SALTLEN

    Dim baData()                     As Byte

    Dim baSalt(0 To KDF_SALTLEN - 1) As Byte

    Dim uCtx                         As UcsCryptoContextType

    Dim lSize                        As Long

    Dim lPadSize                     As Long

    Dim hResult                      As Long

    Dim sApiSource                   As String
    
    baData = ToUtf8Array(sText)
    lSize = UBound(baData) + 1

    If lSize = 0 Then
        GoTo QH

    End If

    Call RtlGenRandom(baSalt(0), KDF_SALTLEN)

    If Not pvCryptoAesCbcInit(uCtx, ToUtf8Array(sPassword), baSalt, AES_KEYLEN) Then
        GoTo QH

    End If

    lPadSize = (lSize + AES_BLOCK_SIZE - 1) And -AES_BLOCK_SIZE
    ReDim Preserve baData(0 To lPadSize - 1) As Byte

    If CryptEncrypt(uCtx.hKey, 0, 1, 0, baData(0), lSize, UBound(baData) + 1) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncrypt"
        GoTo QH

    End If

    ReDim Preserve baData(0 To UBound(baData) + PREFIXLEN) As Byte

    If UBound(baData) >= PREFIXLEN Then
        Call CopyMemory(baData(PREFIXLEN), baData(0), UBound(baData) + 1 - PREFIXLEN)

    End If

    Call CopyMemory(baData(OPENSSL_MAGICLEN), baSalt(0), KDF_SALTLEN)
    Call CopyMemory(baData(0), ByVal OPENSSL_MAGIC, 8)
    AesEncryptString = Replace(ToBase64Array(baData), vbCrLf, vbNullString)
QH:
    pvCryptoAesCbcTerminate uCtx

    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource, GetSystemMessage(hResult)

    End If

End Function

'--- equivalent to `openssl aes256 -pbkdf2 -md sha512 -pass pass:{sPassword} -in {sEncr}.file -a -d`
Public Function AesDecryptString(sEncr As String, sPassword As String) As String

    Const PREFIXLEN As Long = OPENSSL_MAGICLEN + KDF_SALTLEN

    Dim baData()    As Byte

    Dim baSalt()    As Byte

    Dim sMagic      As String

    Dim uCtx        As UcsCryptoContextType

    Dim lSize       As Long

    Dim hResult     As Long

    Dim sApiSource  As String
    
    baData = FromBase64Array(sEncr)
    baSalt = vbNullString

    If UBound(baData) >= PREFIXLEN - 1 Then
        sMagic = String$(OPENSSL_MAGICLEN, 0)
        Call CopyMemory(ByVal sMagic, baData(0), OPENSSL_MAGICLEN)

        If sMagic = OPENSSL_MAGIC Then
            ReDim baSalt(0 To KDF_SALTLEN - 1) As Byte
            Call CopyMemory(baSalt(0), baData(OPENSSL_MAGICLEN), KDF_SALTLEN)

            If UBound(baData) >= PREFIXLEN Then
                Call CopyMemory(baData(0), baData(PREFIXLEN), UBound(baData) + 1 - PREFIXLEN)
                ReDim Preserve baData(0 To UBound(baData) - PREFIXLEN) As Byte
            Else
                GoTo QH

            End If

        End If

    End If

    lSize = UBound(baData) + 1

    If lSize = 0 Then
        GoTo QH

    End If

    If Not pvCryptoAesCbcInit(uCtx, ToUtf8Array(sPassword), baSalt, AES_KEYLEN) Then
        GoTo QH

    End If

    If CryptDecrypt(uCtx.hKey, 0, 1, 0, baData(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptDecrypt"
        GoTo QH

    End If

    If lSize <> UBound(baData) + 1 Then
        ReDim Preserve baData(0 To lSize - 1) As Byte

    End If

    AesDecryptString = FromUtf8Array(baData)
QH:
    pvCryptoAesCbcTerminate uCtx

    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource, GetSystemMessage(hResult)

    End If

End Function

Private Function pvCryptoAesCbcInit(uCtx As UcsCryptoContextType, _
                                    baPass() As Byte, _
                                    baSalt() As Byte, _
                                    ByVal lKeyLen As Long) As Boolean

    Dim uBlob       As BLOBHEADER

    Dim baDerived() As Byte

    Dim hResult     As Long

    Dim sApiSource  As String
    
    With uCtx
        ReDim baDerived(0 To lKeyLen + AES_IVLEN - 1) As Byte

        If Not pvCryptoDeriveKeyPBKDF2(KDF_HASH, baPass, baSalt, KDF_ITER, baDerived) Then
            GoTo QH

        End If

        If CryptAcquireContext(.hProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptAcquireContext"
            GoTo QH

        End If

        uBlob.bType = PLAINTEXTKEYBLOB
        uBlob.bVersion = CUR_BLOB_VERSION

        Select Case lKeyLen

            Case 16
                uBlob.aiKeyAlg = CALG_AES_128

            Case 24
                uBlob.aiKeyAlg = CALG_AES_192

            Case Else
                uBlob.aiKeyAlg = CALG_AES_256

        End Select

        Debug.Assert UBound(uBlob.buffer) >= lKeyLen
        uBlob.cbKeySize = lKeyLen
        Call CopyMemory(uBlob.buffer(0), baDerived(0), lKeyLen)

        If CryptImportKey(.hProv, uBlob, sizeof_BLOBHEADER + uBlob.cbKeySize, 0, 0, .hKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptImportKey"
            GoTo QH

        End If

        If CryptSetKeyParam(.hKey, KP_MODE, CRYPT_MODE_CBC, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptSetKeyParam(KP_MODE)"
            GoTo QH

        End If

        If CryptSetKeyParam(.hKey, KP_IV, baDerived(lKeyLen), 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptSetKeyParam(KP_IV)"
            GoTo QH

        End If

    End With

    '--- success
    pvCryptoAesCbcInit = True
QH:

    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource, GetSystemMessage(hResult)

    End If

End Function

Private Sub pvCryptoAesCbcTerminate(uCtx As UcsCryptoContextType)

    With uCtx

        If .hKey <> 0 Then
            Call CryptDestroyKey(.hKey)
            .hKey = 0

        End If

        If .hProv <> 0 Then
            Call CryptReleaseContext(.hProv, 0)
            .hProv = 0

        End If

    End With

End Sub

Private Function pvCryptoDeriveKeyPBKDF2(sAlgId As String, _
                                         baPass() As Byte, _
                                         baSalt() As Byte, _
                                         ByVal lNumIter As Long, _
                                         baRetVal() As Byte) As Boolean

    Dim lSize      As Long

    Dim lHashAlgId As Long

    Dim lHashSize  As Long

    Dim hProv      As Long

    Dim uBlob      As BLOBHEADER

    Dim hKey       As Long

    Dim baHmac()   As Byte

    Dim lIdx       As Long

    Dim lRemaining As Long

    Dim hResult    As Long

    Dim sApiSource As String
    
    lSize = UBound(baRetVal) + 1

    Select Case UCase$(sAlgId)

        Case "SHA256"
            lHashAlgId = CALG_SHA_256
            lHashSize = 32

        Case "SHA384"
            lHashAlgId = CALG_SHA_384
            lHashSize = 48

        Case "SHA512"
            lHashAlgId = CALG_SHA_512
            lHashSize = 64

        Case Else
            lHashAlgId = CALG_SHA1
            lHashSize = 20

    End Select

    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptAcquireContext"
        GoTo QH

    End If

    uBlob.bType = PLAINTEXTKEYBLOB
    uBlob.bVersion = CUR_BLOB_VERSION
    uBlob.aiKeyAlg = CALG_RC2
    Debug.Assert UBound(uBlob.buffer) >= UBound(baPass)
    uBlob.cbKeySize = UBound(baPass) + 1
    Call CopyMemory(uBlob.buffer(0), baPass(0), uBlob.cbKeySize)

    If CryptImportKey(hProv, uBlob, sizeof_BLOBHEADER + uBlob.cbKeySize, 0, CRYPT_IPSEC_HMAC_KEY, hKey) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptImportKey"
        GoTo QH

    End If

    ReDim baHmac(0 To lHashSize - 1) As Byte

    For lIdx = 0 To (lSize + lHashSize - 1) \ lHashSize - 1

        If Not pvCryptoDeriveKeyHmacPrf(hProv, hKey, lHashAlgId, baSalt, htonl(lIdx + 1), lNumIter, baHmac) Then
            GoTo QH

        End If

        lRemaining = lSize - lIdx * lHashSize

        If lRemaining > lHashSize Then
            lRemaining = lHashSize

        End If

        Call CopyMemory(baRetVal(lIdx * lHashSize), baHmac(0), lRemaining)
    Next
    '--- success
    pvCryptoDeriveKeyPBKDF2 = True
QH:

    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)

    End If

    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)

    End If

    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource, GetSystemMessage(hResult)

    End If

End Function

Private Function pvCryptoDeriveKeyHmacPrf(ByVal hProv As Long, _
                                          ByVal hKey As Long, _
                                          ByVal lHashAlgId As Long, _
                                          baSalt() As Byte, _
                                          ByVal lCounter As Long, _
                                          ByVal lNumIter As Long, _
                                          baRetVal() As Byte) As Boolean

    Dim hHash      As Long

    Dim uInfo      As HMAC_INFO

    Dim baTemp()   As Byte

    Dim lIdx       As Long

    Dim lJdx       As Long

    Dim hResult    As Long

    Dim sApiSource As String
    
    uInfo.HashAlgid = lHashAlgId
    baTemp = baRetVal

    For lIdx = 0 To lNumIter - 1

        If CryptCreateHash(hProv, CALG_HMAC, hKey, 0, hHash) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptCreateHash(CALG_HMAC)"
            GoTo QH

        End If

        If CryptSetHashParam(hHash, HP_HMAC_INFO, uInfo, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptSetHashParam(HP_HMAC_INFO)"
            GoTo QH

        End If

        If lIdx = 0 Then
            If UBound(baSalt) >= 0 Then
                If CryptHashData(hHash, baSalt(0), UBound(baSalt) + 1, 0) = 0 Then
                    hResult = Err.LastDllError
                    sApiSource = "CryptHashData(baSalt)"
                    GoTo QH

                End If

            End If

            If CryptHashData(hHash, lCounter, 4, 0) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CryptHashData(lCounter)"
                GoTo QH

            End If

        Else

            If CryptHashData(hHash, baTemp(0), UBound(baTemp) + 1, 0) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CryptHashData(baTemp)"
                GoTo QH

            End If

        End If

        If CryptGetHashParam(hHash, HP_HASHVAL, baTemp(0), UBound(baTemp) + 1, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptGetHashParam(HP_HASHVAL)"
            GoTo QH

        End If

        If hHash <> 0 Then
            Call CryptDestroyHash(hHash)
            hHash = 0

        End If

        If lIdx = 0 Then
            baRetVal = baTemp
        Else

            For lJdx = 0 To UBound(baTemp)
                baRetVal(lJdx) = baRetVal(lJdx) Xor baTemp(lJdx)
            Next

        End If

    Next
    '--- success
    pvCryptoDeriveKeyHmacPrf = True
QH:

    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)

    End If

    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource, GetSystemMessage(hResult)

    End If

End Function

'= shared ================================================================

#If Not ImplUseShared Then
    Public Function ToBase64Array(baData() As Byte) As String

        Dim lSize As Long
    
        If UBound(baData) >= 0 Then
            ToBase64Array = String$(2 * UBound(baData) + 6, 0)
            lSize = Len(ToBase64Array) + 1
            Call CryptBinaryToString(VarPtr(baData(0)), UBound(baData) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64Array), lSize)
            ToBase64Array = Left$(ToBase64Array, lSize)

        End If

    End Function

Public Function FromBase64Array(sText As String) As Byte()

    Dim lSize      As Long

    Dim baOutput() As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, 0)

    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        FromBase64Array = baOutput
    Else
        FromBase64Array = vbNullString

    End If

End Function

Public Function ToUtf8Array(sText As String) As Byte()

    Dim baRetVal() As Byte

    Dim lSize      As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)

    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString

    End If

    ToUtf8Array = baRetVal

End Function

Public Function FromUtf8Array(baText() As Byte) As String

    Dim lSize As Long
    
    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * UBound(baText), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)

    End If

End Function

Public Function GetSystemMessage(ByVal lLastDllError As Long) As String

    Dim lSize As Long
   
    GetSystemMessage = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)

    If lSize > 2 Then
        If mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2

        End If

    End If

    GetSystemMessage = "[" & lLastDllError & "] " & Left$(GetSystemMessage, lSize)

End Function

#End If

