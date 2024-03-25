Attribute VB_Name = "mSecurity"
'

Option Explicit

Public Enum eInterval

    iUseItem = 0
    iUseItemClick = 1
    iUseSpell = 2

End Enum

Public Type tInterval

    Default As Long
    ModifyTimer As Long
    UseInvalid As Byte

End Type

' Maximo de Intervalos
Public Const MAX_INTERVAL As Byte = 2

' Calcular Tiempos
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Constants Key Code Packets
Public Const MAX_KEY_PACKETS    As Byte = 2

Public Const MAX_KEY_CHANGE     As Byte = 30

' Constants Pointers
Public Const MAX_POINTERS       As Byte = 2

Public Const LIMIT_POINTER      As Byte = 9

Public Const LIMIT_FLOD_POINTER As Byte = 10

' Position of the pointer
Public Enum ePoint

    Point_Spell = 1
    Point_Inv = 2

End Enum

' Cursor X, Y
Public Type tPoint

    X(LIMIT_POINTER) As Long
    Y(LIMIT_POINTER) As Long
    cant(LIMIT_POINTER) As Byte

End Type

Public Type tPackets

    Value As Byte
    cant As Byte

End Type

' Key Code of the special Packets
Public Enum eKeyPackets

    Key_UseItem = 0
    Key_UseSpell = 1
    Key_UseWeapon = 2

End Enum

' Check Key Code
Public Function CheckKeyPacket(ByVal UserIndex As Integer, _
                               ByVal Packet As eKeyPackets, _
                               ByVal KeyPacket As Long) As Boolean

    '<EhHeader>
    On Error GoTo CheckKeyPacket_Err

    '</EhHeader>
                            
    With UserList(UserIndex)
    
        If .KeyPackets(Packet).Value = 0 Then
            'UpdateKeyPacket UserIndex, Packet
            CheckKeyPacket = True

            Exit Function

        End If
        
        If .KeyPackets(Packet).Value <> KeyPacket Then Exit Function
        
        .KeyPackets(Packet).cant = .KeyPackets(Packet).cant + 1
        
        If .KeyPackets(Packet).cant = MAX_KEY_CHANGE Then

            'UpdateKeyPacket UserIndex, Packet
        End If
        
    End With
    
    CheckKeyPacket = True
    '<EhFooter>
    Exit Function

CheckKeyPacket_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mSecurity.CheckKeyPacket " & "at line " & Erl
        
    '</EhFooter>
End Function
                        
' Reset Key Code
Public Sub ResetKeyPackets(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetKeyPackets_Err

    '</EhHeader>

    Dim A As Long
    
    With UserList(UserIndex)
    
        For A = 0 To MAX_KEY_PACKETS
            .KeyPackets(A).Value = 0
            .KeyPackets(A).cant = 0
        Next A
        
    End With

    '<EhFooter>
    Exit Sub

ResetKeyPackets_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mSecurity.ResetKeyPackets " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Reset Pointer
Public Sub ResetPointer(ByVal UserIndex As Integer, ByVal Point As ePoint)

    '<EhHeader>
    On Error GoTo ResetPointer_Err

    '</EhHeader>

    Dim A As Long
    
    With UserList(UserIndex).Pointers(Point)
    
        For A = 0 To LIMIT_POINTER
            .cant(A) = 0
            .X(A) = 0
            .Y(A) = 0
        Next A
        
    End With

    '<EhFooter>
    Exit Sub

ResetPointer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mSecurity.ResetPointer " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub UpdatePointer(ByVal UserIndex As Integer, _
                         ByVal Point As ePoint, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Ident As String)

    '<EhHeader>
    On Error GoTo UpdatePointer_Err

    '</EhHeader>
                        
    Dim A           As Integer

    Dim PointerCero As Byte
    
    With UserList(UserIndex).Pointers(Point)

        For A = 0 To LIMIT_POINTER
            
            ' Pointer cero
            If PointerCero = 0 And (.X(A) = 0 And .Y(A) = 0) Then
                
                PointerCero = A

            End If
            
            ' Pointer repetido
            If .X(A) = X And .Y(A) = Y Then
                .cant(A) = .cant(A) + 1
                
                ' (Máximo permitido en el Point de igualdad)
                If .cant(A) = LIMIT_FLOD_POINTER Then
                    Call Logs_Security(eLog.eSecurity, eLogSecurity.eAntiCheat, Ident & ": Misma posición de marcado " & UserList(UserIndex).Account.Email & " NICK: " & UserList(UserIndex).Name & " IP: " & UserList(UserIndex).IpAddress)
                    
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Ident & ": Misma posición de marcado " & UserList(UserIndex).Account.Email & " NICK: " & UserList(UserIndex).Name & " IP: " & UserList(UserIndex).IpAddress, FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                    
                    ResetPointer UserIndex, Point

                    Exit Sub

                End If
                
                Exit Sub

            End If

        Next A
        
        If PointerCero = 0 Then
            ResetPointer UserIndex, Point
            .X(0) = X
            .Y(0) = Y
            .cant(0) = 1
        Else
            .X(PointerCero) = X
            .Y(PointerCero) = Y
            .cant(PointerCero) = 1

        End If
        
    End With

    '<EhFooter>
    Exit Sub

UpdatePointer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mSecurity.UpdatePointer " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Initialize_Security()

    '<EhHeader>
    On Error GoTo Initialize_Security_Err

    '</EhHeader>
    SecurityKey(0) = 67
    SecurityKey(1) = 24
    SecurityKey(2) = 47
    SecurityKey(3) = 147
    SecurityKey(4) = 29
    SecurityKey(5) = 81
    SecurityKey(6) = 110
    SecurityKey(7) = 94
    SecurityKey(8) = 105
    SecurityKey(9) = 166
    SecurityKey(10) = 4
    SecurityKey(11) = 27
    SecurityKey(12) = 245
    SecurityKey(13) = 252
    SecurityKey(14) = 85
    SecurityKey(15) = 111
    SecurityKey(16) = 94
    SecurityKey(17) = 204
    SecurityKey(18) = 5
    SecurityKey(19) = 66
    SecurityKey(20) = 131
    SecurityKey(21) = 201
    SecurityKey(22) = 11
    SecurityKey(23) = 123
    SecurityKey(24) = 57
    SecurityKey(25) = 195
    SecurityKey(26) = 7
    SecurityKey(27) = 10
    SecurityKey(28) = 64
    SecurityKey(29) = 203
    SecurityKey(30) = 213
    SecurityKey(31) = 44
    SecurityKey(32) = 118
    SecurityKey(33) = 152
    SecurityKey(34) = 98
    SecurityKey(35) = 234
    SecurityKey(36) = 75
    SecurityKey(37) = 41
    SecurityKey(38) = 190
    SecurityKey(39) = 227
    SecurityKey(40) = 117
    SecurityKey(41) = 172
    SecurityKey(42) = 115
    SecurityKey(43) = 76
    SecurityKey(44) = 229
    SecurityKey(45) = 159
    SecurityKey(46) = 22
    SecurityKey(47) = 53
    SecurityKey(48) = 249
    SecurityKey(49) = 53
    SecurityKey(50) = 27
    SecurityKey(51) = 14
    SecurityKey(52) = 243
    SecurityKey(53) = 251
    SecurityKey(54) = 237
    SecurityKey(55) = 105
    SecurityKey(56) = 170
    SecurityKey(57) = 187
    SecurityKey(58) = 62
    SecurityKey(59) = 1
    SecurityKey(60) = 127
    SecurityKey(61) = 160
    SecurityKey(62) = 156
    SecurityKey(63) = 252
    SecurityKey(64) = 147
    SecurityKey(65) = 156
    SecurityKey(66) = 70
    SecurityKey(67) = 109
    SecurityKey(68) = 55
    SecurityKey(69) = 7
    SecurityKey(70) = 61
    SecurityKey(71) = 29
    SecurityKey(72) = 165
    SecurityKey(73) = 137
    SecurityKey(74) = 158
    SecurityKey(75) = 139
    SecurityKey(76) = 237
    SecurityKey(77) = 57
    SecurityKey(78) = 46
    SecurityKey(79) = 49
    SecurityKey(80) = 49
    SecurityKey(81) = 6
    SecurityKey(82) = 55
    SecurityKey(83) = 9
    SecurityKey(84) = 21
    SecurityKey(85) = 86
    SecurityKey(86) = 147
    SecurityKey(87) = 80
    SecurityKey(88) = 114
    SecurityKey(89) = 238
    SecurityKey(90) = 6
    SecurityKey(91) = 138
    SecurityKey(92) = 74
    SecurityKey(93) = 18
    SecurityKey(94) = 208
    SecurityKey(95) = 198
    SecurityKey(96) = 1
    SecurityKey(97) = 237
    SecurityKey(98) = 131
    SecurityKey(99) = 35
    SecurityKey(100) = 202
    SecurityKey(101) = 50
    SecurityKey(102) = 19
    SecurityKey(103) = 147
    SecurityKey(104) = 95
    SecurityKey(105) = 4
    SecurityKey(106) = 44
    SecurityKey(107) = 223
    SecurityKey(108) = 245
    SecurityKey(109) = 20
    SecurityKey(110) = 200
    SecurityKey(111) = 236
    SecurityKey(112) = 111
    SecurityKey(113) = 9
    SecurityKey(114) = 0
    SecurityKey(115) = 60
    SecurityKey(116) = 210
    SecurityKey(117) = 36
    SecurityKey(118) = 34
    SecurityKey(119) = 183
    SecurityKey(120) = 249
    SecurityKey(121) = 36
    SecurityKey(122) = 5
    SecurityKey(123) = 172
    SecurityKey(124) = 137
    SecurityKey(125) = 103
    SecurityKey(126) = 153
    SecurityKey(127) = 19
    SecurityKey(128) = 40
    SecurityKey(129) = 83
    SecurityKey(130) = 194
    SecurityKey(131) = 21
    SecurityKey(132) = 234
    SecurityKey(133) = 244
    SecurityKey(134) = 103
    SecurityKey(135) = 205
    SecurityKey(136) = 12
    SecurityKey(137) = 230
    SecurityKey(138) = 197
    SecurityKey(139) = 81
    SecurityKey(140) = 229
    SecurityKey(141) = 118
    SecurityKey(142) = 10
    SecurityKey(143) = 236
    SecurityKey(144) = 25
    SecurityKey(145) = 4
    SecurityKey(146) = 31
    SecurityKey(147) = 174
    SecurityKey(148) = 16
    SecurityKey(149) = 171
    SecurityKey(150) = 197
    SecurityKey(151) = 39
    SecurityKey(152) = 167
    SecurityKey(153) = 36
    SecurityKey(154) = 227
    SecurityKey(155) = 111
    SecurityKey(156) = 37
    SecurityKey(157) = 232
    SecurityKey(158) = 30
    SecurityKey(159) = 105
    SecurityKey(160) = 112
    SecurityKey(161) = 149
    SecurityKey(162) = 171
    SecurityKey(163) = 73
    SecurityKey(164) = 128
    SecurityKey(165) = 147
    SecurityKey(166) = 97
    SecurityKey(167) = 84
    SecurityKey(168) = 21
    SecurityKey(169) = 247
    SecurityKey(170) = 19
    SecurityKey(171) = 231
    SecurityKey(172) = 165
    SecurityKey(173) = 168
    SecurityKey(174) = 28
    SecurityKey(175) = 187
    SecurityKey(176) = 153
    SecurityKey(177) = 192
    SecurityKey(178) = 59
    SecurityKey(179) = 103
    SecurityKey(180) = 184
    SecurityKey(181) = 53
    SecurityKey(182) = 162
    SecurityKey(183) = 39
    SecurityKey(184) = 228
    SecurityKey(185) = 184
    SecurityKey(186) = 73
    SecurityKey(187) = 219
    SecurityKey(188) = 4
    SecurityKey(189) = 221
    SecurityKey(190) = 136
    SecurityKey(191) = 83
    SecurityKey(192) = 65
    SecurityKey(193) = 125
    SecurityKey(194) = 229
    SecurityKey(195) = 201
    SecurityKey(196) = 117
    SecurityKey(197) = 88
    SecurityKey(198) = 42
    SecurityKey(199) = 175
    SecurityKey(200) = 224
    SecurityKey(201) = 255
    SecurityKey(202) = 187
    SecurityKey(203) = 171
    SecurityKey(204) = 29
    SecurityKey(205) = 242
    SecurityKey(206) = 39
    SecurityKey(207) = 225
    SecurityKey(208) = 85
    SecurityKey(209) = 5
    SecurityKey(210) = 253
    SecurityKey(211) = 112
    SecurityKey(212) = 179
    SecurityKey(213) = 8
    SecurityKey(214) = 225
    SecurityKey(215) = 63
    SecurityKey(216) = 24
    SecurityKey(217) = 166
    SecurityKey(218) = 223
    SecurityKey(219) = 249
    SecurityKey(220) = 15
    SecurityKey(221) = 142
    SecurityKey(222) = 254
    SecurityKey(223) = 86
    SecurityKey(224) = 3
    SecurityKey(225) = 209
    SecurityKey(226) = 25
    SecurityKey(227) = 157
    SecurityKey(228) = 175
    SecurityKey(229) = 139
    SecurityKey(230) = 234
    SecurityKey(231) = 102
    SecurityKey(232) = 215
    SecurityKey(233) = 198
    SecurityKey(234) = 104
    SecurityKey(235) = 165
    SecurityKey(236) = 54
    SecurityKey(237) = 155
    SecurityKey(238) = 83
    SecurityKey(239) = 228
    SecurityKey(240) = 183
    SecurityKey(241) = 154
    SecurityKey(242) = 13
    SecurityKey(243) = 208
    SecurityKey(244) = 232
    SecurityKey(245) = 108
    SecurityKey(246) = 171
    SecurityKey(247) = 247
    SecurityKey(248) = 171
    SecurityKey(249) = 183
    SecurityKey(250) = 76
    SecurityKey(251) = 208
    SecurityKey(252) = 46
    SecurityKey(253) = 66
    SecurityKey(254) = 169
    SecurityKey(255) = 252
    SecurityKey(256) = 30
    SecurityKey(257) = 90
    SecurityKey(258) = 238
    SecurityKey(259) = 203
    SecurityKey(260) = 24
    SecurityKey(261) = 116
    SecurityKey(262) = 200
    SecurityKey(263) = 2
    SecurityKey(264) = 97
    SecurityKey(265) = 19
    SecurityKey(266) = 192
    SecurityKey(267) = 220
    SecurityKey(268) = 214
    SecurityKey(269) = 237
    SecurityKey(270) = 199
    SecurityKey(271) = 78
    SecurityKey(272) = 38
    SecurityKey(273) = 73
    SecurityKey(274) = 18
    SecurityKey(275) = 143
    SecurityKey(276) = 62
    SecurityKey(277) = 171
    SecurityKey(278) = 40
    SecurityKey(279) = 216
    SecurityKey(280) = 5
    SecurityKey(281) = 179
    SecurityKey(282) = 57
    SecurityKey(283) = 104
    SecurityKey(284) = 74
    SecurityKey(285) = 67
    SecurityKey(286) = 177
    SecurityKey(287) = 204
    SecurityKey(288) = 250
    SecurityKey(289) = 224
    SecurityKey(290) = 13
    SecurityKey(291) = 93
    SecurityKey(292) = 151
    SecurityKey(293) = 91
    SecurityKey(294) = 237
    SecurityKey(295) = 10
    SecurityKey(296) = 229
    SecurityKey(297) = 176
    SecurityKey(298) = 107
    SecurityKey(299) = 88
    SecurityKey(300) = 231
    SecurityKey(301) = 46
    SecurityKey(302) = 172
    SecurityKey(303) = 166
    SecurityKey(304) = 9
    SecurityKey(305) = 216
    SecurityKey(306) = 180
    SecurityKey(307) = 182
    SecurityKey(308) = 159
    SecurityKey(309) = 12
    SecurityKey(310) = 127
    SecurityKey(311) = 105
    SecurityKey(312) = 142
    SecurityKey(313) = 98
    SecurityKey(314) = 77
    SecurityKey(315) = 202
    SecurityKey(316) = 73
    SecurityKey(317) = 215
    SecurityKey(318) = 61
    SecurityKey(319) = 78
    SecurityKey(320) = 0
    SecurityKey(321) = 43
    SecurityKey(322) = 29
    SecurityKey(323) = 90
    SecurityKey(324) = 19
    SecurityKey(325) = 135
    SecurityKey(326) = 129
    SecurityKey(327) = 6
    SecurityKey(328) = 205
    SecurityKey(329) = 99
    SecurityKey(330) = 18
    SecurityKey(331) = 33
    SecurityKey(332) = 79
    SecurityKey(333) = 167
    SecurityKey(334) = 41
    SecurityKey(335) = 117
    SecurityKey(336) = 202
    SecurityKey(337) = 16
    SecurityKey(338) = 157
    SecurityKey(339) = 76
    SecurityKey(340) = 242
    SecurityKey(341) = 214
    SecurityKey(342) = 216
    SecurityKey(343) = 50
    SecurityKey(344) = 175
    SecurityKey(345) = 140
    SecurityKey(346) = 49
    SecurityKey(347) = 253
    SecurityKey(348) = 21
    SecurityKey(349) = 71
    SecurityKey(350) = 117
    SecurityKey(351) = 11
    SecurityKey(352) = 150
    SecurityKey(353) = 2
    SecurityKey(354) = 199
    SecurityKey(355) = 203
    SecurityKey(356) = 118
    SecurityKey(357) = 65
    SecurityKey(358) = 171
    SecurityKey(359) = 127
    SecurityKey(360) = 128
    SecurityKey(361) = 245
    SecurityKey(362) = 93
    SecurityKey(363) = 64
    SecurityKey(364) = 248
    SecurityKey(365) = 160
    SecurityKey(366) = 103
    SecurityKey(367) = 66
    SecurityKey(368) = 208
    SecurityKey(369) = 185
    SecurityKey(370) = 114
    SecurityKey(371) = 89
    SecurityKey(372) = 30
    SecurityKey(373) = 82
    SecurityKey(374) = 93
    SecurityKey(375) = 188
    SecurityKey(376) = 206
    SecurityKey(377) = 248
    SecurityKey(378) = 140
    SecurityKey(379) = 9
    SecurityKey(380) = 148
    SecurityKey(381) = 219
    SecurityKey(382) = 131
    SecurityKey(383) = 138
    SecurityKey(384) = 37
    SecurityKey(385) = 46
    SecurityKey(386) = 179
    SecurityKey(387) = 183
    SecurityKey(388) = 167
    SecurityKey(389) = 209
    SecurityKey(390) = 147
    SecurityKey(391) = 252
    SecurityKey(392) = 102
    SecurityKey(393) = 46
    SecurityKey(394) = 243
    SecurityKey(395) = 188
    SecurityKey(396) = 200
    SecurityKey(397) = 96
    SecurityKey(398) = 141
    SecurityKey(399) = 149
    SecurityKey(400) = 131
    SecurityKey(401) = 155
    SecurityKey(402) = 222
    SecurityKey(403) = 230
    SecurityKey(404) = 13
    SecurityKey(405) = 200
    SecurityKey(406) = 52
    SecurityKey(407) = 142
    SecurityKey(408) = 84
    SecurityKey(409) = 111
    SecurityKey(410) = 7
    SecurityKey(411) = 247
    SecurityKey(412) = 176
    SecurityKey(413) = 218
    SecurityKey(414) = 140
    SecurityKey(415) = 83
    SecurityKey(416) = 22
    SecurityKey(417) = 120
    SecurityKey(418) = 136
    SecurityKey(419) = 38
    SecurityKey(420) = 142
    SecurityKey(421) = 127
    SecurityKey(422) = 98
    SecurityKey(423) = 5
    SecurityKey(424) = 231
    SecurityKey(425) = 213
    SecurityKey(426) = 125
    SecurityKey(427) = 157
    SecurityKey(428) = 169
    SecurityKey(429) = 49
    SecurityKey(430) = 196
    SecurityKey(431) = 246
    SecurityKey(432) = 75
    SecurityKey(433) = 125
    SecurityKey(434) = 135
    SecurityKey(435) = 249
    SecurityKey(436) = 166
    SecurityKey(437) = 127
    SecurityKey(438) = 133
    SecurityKey(439) = 49
    SecurityKey(440) = 170
    SecurityKey(441) = 185
    SecurityKey(442) = 74
    SecurityKey(443) = 206
    SecurityKey(444) = 80
    SecurityKey(445) = 142
    SecurityKey(446) = 187
    SecurityKey(447) = 239
    SecurityKey(448) = 207
    SecurityKey(449) = 165
    SecurityKey(450) = 239
    SecurityKey(451) = 33
    SecurityKey(452) = 19
    SecurityKey(453) = 147
    SecurityKey(454) = 64
    SecurityKey(455) = 34
    SecurityKey(456) = 107
    SecurityKey(457) = 180
    SecurityKey(458) = 162
    SecurityKey(459) = 235
    SecurityKey(460) = 130
    SecurityKey(461) = 89
    SecurityKey(462) = 52
    SecurityKey(463) = 238
    SecurityKey(464) = 144
    SecurityKey(465) = 41
    SecurityKey(466) = 21
    SecurityKey(467) = 157
    SecurityKey(468) = 209
    SecurityKey(469) = 193
    SecurityKey(470) = 121
    SecurityKey(471) = 43
    SecurityKey(472) = 54
    SecurityKey(473) = 158
    SecurityKey(474) = 252
    SecurityKey(475) = 150
    SecurityKey(476) = 91
    SecurityKey(477) = 61
    SecurityKey(478) = 53
    SecurityKey(479) = 229
    SecurityKey(480) = 186
    SecurityKey(481) = 128
    SecurityKey(482) = 143
    SecurityKey(483) = 174
    SecurityKey(484) = 30
    SecurityKey(485) = 84
    SecurityKey(486) = 84
    SecurityKey(487) = 220
    SecurityKey(488) = 90
    SecurityKey(489) = 145
    SecurityKey(490) = 11
    SecurityKey(491) = 175
    SecurityKey(492) = 58
    SecurityKey(493) = 33
    SecurityKey(494) = 4
    SecurityKey(495) = 4
    SecurityKey(496) = 186
    SecurityKey(497) = 101
    SecurityKey(498) = 49
    SecurityKey(499) = 215
    SecurityKey(500) = 118
    '<EhFooter>
    Exit Sub

Initialize_Security_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mSecurity.Initialize_Security " & "at line " & Erl
         
    '</EhFooter>
End Sub

Private Function PacketID_Change(ByVal Selected As Byte) As Integer

    '<EhHeader>
    On Error GoTo PacketID_Change_Err

    '</EhHeader>
    
    Dim Temp     As Integer

    Dim KeyText  As String

    Dim KeyValue As String
    
    Select Case Selected

        Case 75
            KeyValue = "GAHBDEWIDKFLSQ2DIWJNE"

        Case 150
            KeyValue = "AGSQEFHFFDFSDQETUHFLSJNE"

        Case 99
            KeyValue = "13SDDJS2s"

        Case 105
            KeyValue = "ADSDEWEFFDFGRT"

    End Select
    
    Temp = 127
    Temp = Temp Xor 45
    
    If Len(KeyValue) > 10 Then
        Temp = Temp Xor 4 Xor Selected
    Else
        Temp = Temp Xor 75

    End If
    
    '<EhFooter>
    Exit Function

PacketID_Change_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mSecurity.PacketID_Change " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function ReadPacketID(ByVal PacketID As Integer) As Integer

    '<EhHeader>
    On Error GoTo ReadPacketID_Err

    '</EhHeader>
    
    Dim KeyTempOne   As Integer

    Dim KeyTempTwo   As Integer

    Dim KeyTempThree As Integer
    
    Dim KeyOne       As String: KeyOne = "137"

    Dim KeyTwo       As String: KeyTwo = "215"

    Dim KeyThree     As String: KeyThree = "45"

    Dim KeyFour      As String: KeyFour = "12"

    Dim KeyFive      As String: KeyFive = "197"
    
    PacketID = PacketID Xor 127
    KeyTempOne = 127
    PacketID = PacketID Xor 67
    PacketID = PacketID Xor Len(KeyOne)
    KeyTempOne = KeyTempOne Xor 12
    
    PacketID = PacketID Xor PacketID_Change(99)
    
    If PacketID Then
        PacketID = PacketID Xor Len(KeyTwo)
        PacketID = PacketID Xor Len(KeyThree)
        
        PacketID = PacketID Xor PacketID_Change(75)
    Else
        PacketID = PacketID Xor Len(KeyOne)
        PacketID = PacketID Xor Len(KeyThree)
        PacketID = PacketID Xor PacketID_Change(99)

    End If
    
    KeyTempOne = KeyTempOne Xor PacketID
    
    If KeyTempOne > 55 Then
        KeyTempTwo = KeyTempTwo Xor 49
        KeyTempThree = KeyTempThree Xor 75
    ElseIf KeyTempOne > 150 Then
        KeyTempTwo = KeyTempTwo Xor 49
        KeyTempThree = KeyTempThree Xor 75
    ElseIf KeyTempOne > 250 Then
        KeyTempTwo = KeyTempTwo Xor 49

    End If
    
    PacketID = PacketID Xor KeyOne
    KeyTempTwo = KeyTempTwo Xor KeyTempOne Xor PacketID_Change(150)
    KeyTempThree = KeyTempOne Xor KeyTempTwo
    PacketID = PacketID Xor 75 Xor PacketID_Change(105)
    
    KeyTempTwo = PacketID Xor KeyTempThree
    PacketID = PacketID Xor 21
    
    PacketID = PacketID Xor Len(KeyFive)
    
    ReadPacketID = PacketID
    '<EhFooter>
    Exit Function

ReadPacketID_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mSecurity.ReadPacketID " & "at line " & Erl
        
    '</EhFooter>
End Function
