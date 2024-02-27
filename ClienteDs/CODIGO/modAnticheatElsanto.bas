Attribute VB_Name = "modAnalisis"
Option Explicit
'Autor: El_Santo43
'Para: Desterium AO
'Este modulo analiza padrones de intervalos entre clicks y lugares especificos donde se clickea (X e Y)
'para detectar distintos tipos de macros, mouse gamers, y otros automatizadores de eventos de click
'Tiene llamadas en los respectivos subs del frmMain, en botones y inventario. Hasta la vista macreros :)
'Aunque esta seguridad no es infalible, es dificil de saltar ya que raramente encontramos alguien que
'sepa usar un editor de memorias
'Lo mas importante es no mostrarselo a nadie nunca, sino es probable que encuentren la forma de sobrepasar
'los analisis de padrones de clicks.
'Buen y sano agite :)
'Cuando se sobrepasan cierta cantidad de similaridades en los padrones de comportamiento, se manda un paquete
'al servidor para guardar en la carpeta logs/CHEATERS.log
'El paquete no tiene ningun nombre que lo delate como para alguien q se ponga a analizar los paquetes


'Comprobar la velocidad de los clicks, para ver si existen padrones, lo que va a mostrar si los clicks son hechos
'por un humano o por un software. Primer capa de seguridad antimacros

  
'   * Colocar un control Timer
'-------------------------------------------------
  
'Estructura de coordenadas para el api GetCursorPos
Private Type POINTAPI
    X As Long
    Y As Long
End Type
  
  
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Enum eTipo
    BotonLanzar = 1
    BotonHechizos = 2
    BotonInventario = 3
    ListaLanzar = 4
    InventarioObj = 5
    InvObjHechizo = 6
    LanzarAutocurar = 7
End Enum


Public Ac(1 To 10) As clsAnalizarPatrones

' Se almazenan los intervalos entre el click en la lista de hechizos y el boton lanzar
' y luego de 5 intervalos, se comparan

Public Sub initAntiCheat()

    Dim i As Long
    For i = 1 To 10
        Set Ac(i) = New clsAnalizarPatrones
    Next i
    
    Call Ac(eTipo.BotonLanzar).Inicializar(False, eTipo.BotonLanzar)
    Call Ac(eTipo.BotonHechizos).Inicializar(False, eTipo.BotonHechizos)
    Call Ac(eTipo.BotonInventario).Inicializar(False, eTipo.BotonInventario)
    
    Call Ac(eTipo.ListaLanzar).Inicializar(True, eTipo.ListaLanzar)
    Call Ac(eTipo.InventarioObj).Inicializar(True, eTipo.InventarioObj)
    Call Ac(eTipo.InvObjHechizo).Inicializar(True, eTipo.InvObjHechizo)
    Call Ac(eTipo.LanzarAutocurar).Inicializar(True, eTipo.LanzarAutocurar)

End Sub
Public Sub ClickCambioInv() 'cambio al menu inventario
    Call Ac(eTipo.InventarioObj).PrimerEvento
    
End Sub

Public Sub ClickEnInv() 'Click en algun objeto del inventario
    Call Ac(eTipo.InventarioObj).SegundoEvento
    
    Call Ac(eTipo.InvObjHechizo).PrimerEvento
End Sub

Public Sub ClickCambioHech() ' cambio al menu de hechizos
    Call Ac(eTipo.InvObjHechizo).SegundoEvento
End Sub

Public Sub ClickLista()
    Call Ac(eTipo.ListaLanzar).PrimerEvento
End Sub


Public Sub ClickLanzar()
    Call Ac(eTipo.ListaLanzar).SegundoEvento
    Call Ac(eTipo.LanzarAutocurar).PrimerEvento
End Sub

Public Sub ClickEnObjetoPos(ByVal Tipo As eTipo, ByVal X As Single, ByVal Y As Single)
    Call Ac(Tipo).ClickEnObjetoPos(X, Y)
End Sub


