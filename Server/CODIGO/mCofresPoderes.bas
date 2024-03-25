Attribute VB_Name = "mCofresPoderes"
Option Explicit

Private Enum eCofres

    cBronce = 1
    cPlata = 2
    cOro = 3
    cPremium = 4
    cStreamer = 5

End Enum

Public Function UseCofrePoder(ByVal UserIndex As Integer, _
                              ByVal CofreIndex As Byte) As Boolean

    '<EhHeader>
    On Error GoTo UseCofrePoder_Err

    '</EhHeader>

    Dim ObjRequired As Obj

    Dim TempSTR     As String

    Dim Ft          As FontTypeNames
    
    With UserList(UserIndex)
    
        ' REQUISITOS
        Select Case CofreIndex
            
            Case eCofres.cStreamer

                If .flags.Streamer = 1 Then
                    WriteConsoleMsg UserIndex, "Ya eres considerado un usuario Streamer", FontTypeNames.FONTTYPE_INFORED
                    Exit Function

                End If
                
                TempSTR = "Servidor> El usuario " & .Name & " ha sido considerado como streamer de la comunidad."
                Ft = FontTypeNames.FONTTYPE_GUILD
                
            Case eCofres.cOro
                
                TempSTR = "¡Te has convertido en una Leyenda!"
                Ft = FontTypeNames.FONTTYPE_USERGOLD
                
            Case eCofres.cBronce
                
                TempSTR = "¡Te has convertido en un Aventurero!"
                Ft = FontTypeNames.FONTTYPE_USERBRONCE
                
            Case eCofres.cPlata
                TempSTR = "¡Te vas convertido en un Héroe!"
                Ft = FontTypeNames.FONTTYPE_USERPLATA
            
            Case eCofres.cPremium

                If .flags.Premium = 1 Then
                    WriteConsoleMsg UserIndex, "Tu personaje ya posee el poder del cofre seleccionado.", FontTypeNames.FONTTYPE_INFO

                    Exit Function

                End If
                
                TempSTR = "Te has convertido en un PERSONAJE PREMIUM"
                Ft = FontTypeNames.FONTTYPE_USERPREMIUM

        End Select

        ' APLICAMOS
        Select Case CofreIndex

            Case eCofres.cStreamer
                .flags.Streamer = 1
                
            Case eCofres.cOro
                .flags.Oro = 1

            Case eCofres.cPremium
                .flags.Premium = 1

            Case eCofres.cPlata
                .flags.Plata = 1

            Case eCofres.cBronce
                .flags.Bronce = 1

        End Select
        
        Call WriteConsoleMsg(UserIndex, TempSTR, Ft)
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(TempSTR, Ft))
        UseCofrePoder = True

    End With
    
    '<EhFooter>
    Exit Function

UseCofrePoder_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mCofresPoderes.UseCofrePoder " & "at line " & Erl
        
    '</EhFooter>
End Function
