VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComerciarUsu.frx":0000
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInvOroProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3450
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   930
      Width           =   960
   End
   Begin VB.TextBox txtAgregar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4500
      TabIndex        =   6
      Top             =   2295
      Width           =   1035
   End
   Begin VB.PictureBox picInvOroOfertaOtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   5040
      Width           =   960
   End
   Begin VB.PictureBox picInvOfertaOtro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   7080
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   5040
      Width           =   2400
   End
   Begin VB.PictureBox picInvOfertaProp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   7080
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   3
      Top             =   930
      Width           =   2400
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   495
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   7965
      Width           =   6060
   End
   Begin VB.PictureBox picInvComercio 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   480
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   945
      Width           =   2400
   End
   Begin VB.PictureBox picInvOroOfertaProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   930
      Width           =   960
   End
   Begin RichTextLib.RichTextBox CommerceConsole 
      Height          =   1620
      Left            =   495
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   6030
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   2858
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmComerciarUsu.frx":23C60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   360
      Tag             =   "1"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgRechazar 
      Height          =   360
      Left            =   8220
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image imgConfirmar 
      Height          =   360
      Left            =   7440
      Tag             =   "2"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   6750
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image imgAgregar 
      Height          =   255
      Left            =   4800
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgQuitar 
      Height          =   255
      Left            =   4800
      Top             =   2760
      Width           =   375
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmComerciarUsu.frm
'
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton
Private cBotonRechazar As clsGraphicalButton
Private cBotonConfirmar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Const GOLD_OFFER_SLOT As Byte = INV_OFFER_SLOTS + 1

Private sCommerceChat As String

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()
10        If Not cBotonAceptar.IsEnabled Then Exit Sub  ' Deshabilitado
          
20        Call WriteUserCommerceOk
30        HabilitarAceptarRechazar False
          
End Sub

Private Sub imgAgregar_Click()
         
          ' No tiene seleccionado ningun item
10        If InvComUsu.SelectedItem = 0 Then
20            Call PrintCommerceMsg("¡No tienes ningún item seleccionado!", _
                  FontTypeNames.FONTTYPE_FIGHT)
30            Exit Sub
40        End If
          
          ' Numero invalido
50        If Not IsNumeric(txtAgregar.Text) Then Exit Sub
          
60        HabilitarConfirmar True
          
          Dim OfferSlot As Byte
          Dim Amount As Long
          Dim InvSlot As Byte
              
70        With InvComUsu
80            If .SelectedItem = FLAGORO Then
90                If Val(txtAgregar.Text) > InvOroComUsu(0).Amount(1) Then
100                   Call PrintCommerceMsg("¡No tienes esa cantidad!", _
                          FontTypeNames.FONTTYPE_FIGHT)
110                   Exit Sub
120               End If
                  
130               Amount = InvOroComUsu(1).Amount(1) + Val(txtAgregar.Text)
          
                  ' Le aviso al otro de mi cambio de oferta
140               Call WriteUserCommerceOffer(FLAGORO, Val(txtAgregar.Text), _
                      GOLD_OFFER_SLOT)
                  
                  ' Actualizo los inventarios
150               Call InvOroComUsu(0).ChangeSlotItemAmount(1, _
                      InvOroComUsu(0).Amount(1) - Val(txtAgregar.Text))
160               Call InvOroComUsu(1).ChangeSlotItemAmount(1, Amount)
                  
170               Call PrintCommerceMsg("¡Agregaste " & Val(txtAgregar.Text) & _
                      " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & _
                      " de oro a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
                  
180           ElseIf .SelectedItem > 0 Then
190                If Val(txtAgregar.Text) > .Amount(.SelectedItem) Then
200                   Call PrintCommerceMsg("¡No tienes esa cantidad!", _
                          FontTypeNames.FONTTYPE_FIGHT)
210                   Exit Sub
220               End If
                   
230               OfferSlot = CheckAvailableSlot(.SelectedItem, Val(txtAgregar.Text))
                  
                  ' Hay espacio o lugar donde sumarlo?
240               If OfferSlot > 0 Then
                  
250                   Call PrintCommerceMsg("¡Agregaste " & Val(txtAgregar.Text) & " " _
                          & .ItemName(.SelectedItem) & " a tu oferta!!", _
                          FontTypeNames.FONTTYPE_GUILD)
                      
                      ' Le aviso al otro de mi cambio de oferta
260                   Call WriteUserCommerceOffer(.SelectedItem, Val(txtAgregar.Text), _
                          OfferSlot)
                      
                      ' Actualizo el inventario general de comercio
270                   Call .ChangeSlotItemAmount(.SelectedItem, _
                          .Amount(.SelectedItem) - Val(txtAgregar.Text))
                      
280                   Amount = InvOfferComUsu(0).Amount(OfferSlot) + _
                          Val(txtAgregar.Text)
                      
                      ' Actualizo los inventarios
290                   If InvOfferComUsu(0).ObjIndex(OfferSlot) > 0 Then
                          ' Si ya esta el item, solo actualizo su cantidad en el invenatario
300                       Call InvOfferComUsu(0).ChangeSlotItemAmount(OfferSlot, _
                              Amount)
310                   Else
320                       InvSlot = .SelectedItem
                          ' Si no agrego todo
330                       Call InvOfferComUsu(0).SetItem(OfferSlot, _
                              .ObjIndex(InvSlot), Amount, 0, .GrhIndex(InvSlot), _
                              .ObjType(InvSlot), .MaxHit(InvSlot), .MinHit(InvSlot), _
                              .MaxDef(InvSlot), .MinDef(InvSlot), .Valor(InvSlot), _
                              .ItemName(InvSlot))
340                   End If
350               End If
360           End If
370       End With
End Sub

Private Sub imgCancelar_Click()
10        Call WriteUserCommerceEnd
End Sub

Private Sub imgConfirmar_Click()
10        If Not cBotonConfirmar.IsEnabled Then Exit Sub  ' Deshabilitado
          
20        HabilitarConfirmar False
30        imgAgregar.Visible = False
40        imgQuitar.Visible = False
50        txtAgregar.Enabled = False
          
60        Call PrintCommerceMsg("¡Has confirmado tu oferta! Ya no puedes cambiarla.", _
              FontTypeNames.FONTTYPE_CONSE)
70        Call WriteUserCommerceConfirm
End Sub

Private Sub imgQuitar_Click()
          Dim Amount As Long
          Dim InvComSlot As Byte

          ' No tiene seleccionado ningun item
10        If InvOfferComUsu(0).SelectedItem = 0 Then
20            Call PrintCommerceMsg("¡No tienes ningún ítem seleccionado!", _
                  FontTypeNames.FONTTYPE_FIGHT)
30            Exit Sub
40        End If
          
          ' Numero invalido
50        If Not IsNumeric(txtAgregar.Text) Then Exit Sub

          ' Comparar con el inventario para distribuir los items
60        If InvOfferComUsu(0).SelectedItem = FLAGORO Then
70            Amount = IIf(Val(txtAgregar.Text) > InvOroComUsu(1).Amount(1), _
                  InvOroComUsu(1).Amount(1), Val(txtAgregar.Text))
              ' Estoy quitando, paso un valor negativo
80            Amount = Amount * (-1)
              
              ' No tiene sentido que se quiten 0 unidades
90            If Amount <> 0 Then
                  ' Le aviso al otro de mi cambio de oferta
100               Call WriteUserCommerceOffer(FLAGORO, Amount, GOLD_OFFER_SLOT)
                  
                  ' Actualizo los inventarios
110               Call InvOroComUsu(0).ChangeSlotItemAmount(1, _
                      InvOroComUsu(0).Amount(1) - Amount)
120               Call InvOroComUsu(1).ChangeSlotItemAmount(1, _
                      InvOroComUsu(1).Amount(1) + Amount)
              
130               Call PrintCommerceMsg("¡¡Quitaste " & Amount * (-1) & " moneda" & _
                      IIf(Val(txtAgregar.Text) = 1, "", "s") & " de oro de tu oferta!!", _
                      FontTypeNames.FONTTYPE_GUILD)
140           End If
150       Else
160           Amount = IIf(Val(txtAgregar.Text) > _
                  InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), _
                  InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), _
                  Val(txtAgregar.Text))
              ' Estoy quitando, paso un valor negativo
170           Amount = Amount * (-1)
              
              ' No tiene sentido que se quiten 0 unidades
180           If Amount <> 0 Then
190               With InvOfferComUsu(0)
                      
200                   Call PrintCommerceMsg("¡¡Quitaste " & Amount * (-1) & " " & _
                          .ItemName(.SelectedItem) & " de tu oferta!!", _
                          FontTypeNames.FONTTYPE_GUILD)
          
                      ' Le aviso al otro de mi cambio de oferta
210                   Call WriteUserCommerceOffer(0, Amount, .SelectedItem)
                  
                      ' Actualizo el inventario general
220                   Call UpdateInvCom(.ObjIndex(.SelectedItem), Abs(Amount))
                       
                       ' Actualizo el inventario de oferta
230                    If .Amount(.SelectedItem) + Amount = 0 Then
                           ' Borro el item
240                        Call .SetItem(.SelectedItem, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                               "")
250                    Else
                           ' Le resto la cantidad deseada
260                        Call .ChangeSlotItemAmount(.SelectedItem, _
                               .Amount(.SelectedItem) + Amount)
270                    End If
280               End With
290           End If
300       End If
          
          ' Si quito todos los items de la oferta, no puede confirmarla
310       If Not HasAnyItem(InvOfferComUsu(0)) And Not HasAnyItem(InvOroComUsu(1)) _
              Then HabilitarConfirmar (False)
End Sub

Private Sub imgRechazar_Click()
10        If Not cBotonRechazar.IsEnabled Then Exit Sub  ' Deshabilitado
          
20        Call WriteUserCommerceReject
End Sub

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me

          'Me.Picture = LoadPicture(DirGraficos & "VentanaComercioUsuario.jpg")
          
30        LoadButtons
          
40        Call _
              PrintCommerceMsg("> Una vez termines de formar tu oferta, debes presionar en ""Confirmar"", tras lo cual ya no podrás modificarla.", FontTypeNames.FONTTYPE_GUILDMSG)
50        Call _
              PrintCommerceMsg("> Luego que el otro usuario confirme su oferta, podrás aceptarla o rechazarla. Si la rechazas, se terminará el comercio.", _
              FontTypeNames.FONTTYPE_GUILDMSG)
60        Call _
              PrintCommerceMsg("> Cuando ambos acepten la oferta del otro, se realizará el intercambio.", _
              FontTypeNames.FONTTYPE_GUILDMSG)
70        Call _
              PrintCommerceMsg("> Si se intercambian más ítems de los que pueden entrar en tu inventario, es probable que caigan al suelo, así que presta mucha atención a esto.", _
              FontTypeNames.FONTTYPE_GUILDMSG)
          
End Sub

Private Sub LoadButtons()

          Dim GrhPath As String
10        GrhPath = DirGraficos
          
20        Set cBotonAceptar = New clsGraphicalButton
30        Set cBotonConfirmar = New clsGraphicalButton
40        Set cBotonRechazar = New clsGraphicalButton
50        Set cBotonCancelar = New clsGraphicalButton
          
60        Set LastButtonPressed = New clsGraphicalButton
          
End Sub

Private Sub Form_LostFocus()
10        Me.SetFocus
End Sub

Private Sub SubtxtAgregar_Change()
10        If Val(txtAgregar.Text) < 1 Then txtAgregar.Text = "1"

20        If Val(txtAgregar.Text) > 2147483647 Then txtAgregar.Text = "2147483647"
End Sub

Private Sub picInvComercio_Click()
10        Call InvOroComUsu(0).DeselectItem
End Sub

Private Sub picInvComercio_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
10        LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOfertaOtro_MouseMove(Button As Integer, Shift As Integer, X _
    As Single, Y As Single)
10        LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOfertaProp_Click()
10        InvOroComUsu(1).DeselectItem
End Sub

Private Sub picInvOfertaProp_MouseMove(Button As Integer, Shift As Integer, X _
    As Single, Y As Single)
10        LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOroOfertaOtro_Click()
          ' No se puede seleccionar el oro que oferta el otro :P
10        InvOroComUsu(2).DeselectItem
End Sub

Private Sub picInvOroOfertaProp_Click()
10        InvOfferComUsu(0).SelectGold
End Sub

Private Sub picInvOroProp_Click()
10        InvComUsu.SelectGold
End Sub

Private Sub SendTxt_Change()
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 03/10/2009
      '**************************************************************
10        If Len(SendTxt.Text) > 160 Then
20            sCommerceChat = "Soy un cheater, avisenle a un gm"
30        Else
              'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
              Dim i As Long
              Dim TempStr As String
              Dim CharAscii As Integer
              
40            For i = 1 To Len(SendTxt.Text)
50                CharAscii = Asc(mid$(SendTxt.Text, i, 1))
60                If CharAscii >= vbKeySpace And CharAscii <= 250 Then
70                    TempStr = TempStr & Chr$(CharAscii)
80                End If
90            Next i
              
100           If TempStr <> SendTxt.Text Then
                  'We only set it if it's different, otherwise the event will be raised
                  'constantly and the client will crush
110               SendTxt.Text = TempStr
120           End If
              
130           sCommerceChat = SendTxt.Text
140       End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
10        If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii _
              <= 250) Then KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
          'Send text
10        If KeyCode = vbKeyReturn Then
20            If LenB(sCommerceChat) <> 0 Then Call WriteCommerceChat(sCommerceChat)
              
30            sCommerceChat = ""
40            SendTxt.Text = ""
50            KeyCode = 0
60        End If
End Sub

Private Sub txtAgregar_Change()
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 03/10/2009
      '**************************************************************
          'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
          Dim i As Long
          Dim TempStr As String
          Dim CharAscii As Integer
          
10        For i = 1 To Len(txtAgregar.Text)
20            CharAscii = Asc(mid$(txtAgregar.Text, i, 1))
              
30            If CharAscii >= 48 And CharAscii <= 57 Then
40                TempStr = TempStr & Chr$(CharAscii)
50            End If
60        Next i
          
70        If TempStr <> txtAgregar.Text Then
              'We only set it if it's different, otherwise the event will be raised
              'constantly and the client will crush
80            txtAgregar.Text = TempStr
90        End If
End Sub

Private Sub txtAgregar_KeyDown(KeyCode As Integer, Shift As Integer)
10    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or KeyCode = _
          vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
20        KeyCode = 0
30    End If

End Sub

Private Sub txtAgregar_KeyPress(KeyAscii As Integer)
10    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii _
          = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
          'txtCant = KeyCode
20        KeyAscii = 0
30    End If

End Sub

Private Function CheckAvailableSlot(ByVal InvSlot As Byte, ByVal Amount As _
    Long) As Byte
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 30/11/2009
      'Search for an available slot to put an item. If found returns the slot, else returns 0.
      '***************************************************
          Dim Slot As Long
10    On Error GoTo Err
          ' Primero chequeo si puedo sumar esa cantidad en algun slot que ya tenga ese item
20        For Slot = 1 To INV_OFFER_SLOTS
30            If InvComUsu.ObjIndex(InvSlot) = InvOfferComUsu(0).ObjIndex(Slot) Then
40                If InvOfferComUsu(0).Amount(Slot) + Amount <= MAX_INVENTORY_OBJS _
                      Then
                      ' Puedo sumarlo aca
50                    CheckAvailableSlot = Slot
60                    Exit Function
70                End If
80            End If
90        Next Slot
          
          ' No lo puedo sumar, me fijo si hay alguno vacio
100       For Slot = 1 To INV_OFFER_SLOTS
110           If InvOfferComUsu(0).ObjIndex(Slot) = 0 Then
                  ' Esta vacio, lo dejo aca
120               CheckAvailableSlot = Slot
130               Exit Function
140           End If
150       Next Slot
160       Exit Function
Err:
170       Debug.Print "Slot: " & Slot
End Function

Public Sub UpdateInvCom(ByVal ObjIndex As Integer, ByVal Amount As Long)
          Dim Slot As Byte
          Dim RemainingAmount As Long
          Dim DifAmount As Long
          
10        RemainingAmount = Amount
          
20        For Slot = 1 To MAX_INVENTORY_SLOTS
              
30            If InvComUsu.ObjIndex(Slot) = ObjIndex Then
40                DifAmount = Inventario.Amount(Slot) - InvComUsu.Amount(Slot)
50                If DifAmount > 0 Then
60                    If RemainingAmount > DifAmount Then
70                        RemainingAmount = RemainingAmount - DifAmount
80                        Call InvComUsu.ChangeSlotItemAmount(Slot, _
                              Inventario.Amount(Slot))
90                    Else
100                       Call InvComUsu.ChangeSlotItemAmount(Slot, _
                              InvComUsu.Amount(Slot) + RemainingAmount)
110                       Exit Sub
120                   End If
130               End If
140           End If
150       Next Slot
End Sub

Public Sub PrintCommerceMsg(ByRef msg As String, ByVal FontIndex As Integer)
          
10        With FontTypes(FontIndex)
20            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, msg, .red, _
                  .green, .blue, .bold, .italic)
30        End With
          
End Sub

#If Wgl = 0 Then
Public Function HasAnyItem(ByRef Inventory As clsGrapchicalInventory) As Boolean
#Else
Public Function HasAnyItem(ByRef Inventory As clsGrapchicalInventoryWgl) As Boolean
#End If
          Dim Slot As Long
          
10        For Slot = 1 To Inventory.MaxObjs
20            If Inventory.Amount(Slot) > 0 Then HasAnyItem = True: Exit Function
30        Next Slot
          
End Function

Public Sub HabilitarConfirmar(ByVal Habilitar As Boolean)
10        Call cBotonConfirmar.EnableButton(Habilitar)
End Sub

Public Sub HabilitarAceptarRechazar(ByVal Habilitar As Boolean)
10        Call cBotonAceptar.EnableButton(Habilitar)
20        Call cBotonRechazar.EnableButton(Habilitar)
End Sub


