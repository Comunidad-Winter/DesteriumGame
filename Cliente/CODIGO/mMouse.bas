Attribute VB_Name = "mMouse"
'------------------------------------------------------------------------------
' M�dulo para subclasificaci�n (subclassing)                        (26/Jun/98)
' Revisado (probado) para publicar en mis p�ginas                   (18/Abr/01)
'
' Modificado para usar con la clase clsMouse                    (21/Mar/99)
'
' �Guillermo 'guille' Som, 1998-2001
'
' Para m�s informaci�n sobre subclasificaci�n:
' En la documentaci�n de Visual Basic:
'   Pasar punteros de funci�n a los procedimientos de DLL y a las bibliotecas de tipos
' En la MSDN Library (o en la Knowledge Base):
'   HOWTO: Subclass a UserControl
'       Article ID: Q179398
'   HOWTO: Hook Into a Window's Messages Using AddressOf
'       Article ID: Q168795
'   HOWTO: Build a Windows Message Handler with AddressOf in VB5
'       Article ID: Q170570
'------------------------------------------------------------------------------
Option Explicit

' Un array de la clase que se usar� para subclasificar ventanas
' y el �ltimo elemento de clases en el array; empieza a contar por uno
Private mWSC() As clsMouse          ' Array de clases

Private mnWSC  As Long                   ' N�mero de ventanas subclasificadas

Private Const GWL_WNDPROC = (-4&)

Public Declare Function CallWindowProc _
               Lib "user32" _
               Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                        ByVal hWnd As Long, _
                                        ByVal msg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long

Public Declare Function SetWindowLong _
               Lib "user32" _
               Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long

Public Function WndProc(ByVal hWnd As Long, _
                        ByVal uMSG As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

    ' Los mensajes de windows llegar�n aqu�.
    ' Lo que hay que hacer es "capturar" los que se necesiten,
    ' en este caso se devuelven los mensajes a la clase, usando para
    ' ello un procedimiento p�blico llamado unMSG con los siguientes par�metros:
    ' ByVal uMSG As Long, ByVal wParam As Long, ByVal lParam As Long
    '
    ' Para un mejor uso, usar la clase en el formato:
    '   Dim WithEvents laClase As clsMouse
    '
    Static i As Long
    
    ' Buscar el �ndice de esta clase en el array
    ' NOTA: Esto se har� para cada uno de los mensajes recibidos,
    '       por tanto, no ser�a conveniente tener demasiadas ventanas o controles
    '       subclasificados, con idea de que no tarde demasiado en procesarlos
    '$Por hacer:
    '   Ser�a conveniente poner un l�mite m�ximo de ventanas a subclasificar
    i = IndiceClase(hWnd)
    
    If i Then

        With mWSC(i)
            WndProc = CallWindowProc(.PrevWndProc, hWnd, uMSG, wParam, lParam)
            ' Producir el evento del mensaje recibido
            .unMSG uMSG, wParam, lParam

        End With

    End If

End Function

Public Sub Hook(ByVal WSC As clsMouse, ByVal unControl As Object)
    ' Subclasificar la ventana o control indicado
    
    '--------------------------------------------------------------------------
    ' Nota:
    ' En este procedimiento no se hace chequeo de que el objeto pasado tenga
    ' la propiedad hWnd, ya que se comprueba en el m�todo Hook de la clase,
    ' por tanto no se deber�a llamar a este m�todo sin antes hacer una comprobaci�n
    ' de que estamos pasando un objeto-ventana (que tenga la propiedad hWnd)
    '
    
    ' Comprobar si ya est� subclasificada esta ventana
    Dim claseActual As Long

    Dim claseLibre  As Long
    
    ' Buscar el �ndice de esta clase en el array
    ' y si hay alguna clase liberada anteriormente
    claseActual = IndiceClase(unControl.hWnd, claseLibre)
    
    If claseActual = 0 Then

        ' Si hay un �ndice que ya no se usa...
        If claseLibre Then
            ' se usar� ese �ndice
            claseActual = claseLibre
        Else
            ' Crear una nueva clase
            mnWSC = mnWSC + 1
            ReDim Preserve mWSC(1 To mnWSC)
            claseActual = mnWSC

        End If

    End If
    
    ' Aqu� se est� haciendo referencia a una clase ya existente,
    ' para que no queden referencias "sueltas", en el evento Terminate de la clase
    ' se llama al procedimiento de liberaci�n de la subclasificaci�n en el que se
    ' borrar� la referencia a la clase indicada, por tanto no se deber�a modificar
    ' esa forma de actuar.
    '--------------------------------------------------------------------------
    ' Nota:
    '   En lugar de hacer una referencia a la clase, se podr�a usar un puntero a
    '   la misma usando ObjPtr, pero esto implicar�a usar la funci�n CopyMemory
    '   para poder acceder a las propiedades de la clase, y no s� si esto
    '   incrementar�a el tiempo de procesamiento, pero...
    '   "los expertos" as� lo hacen... as� que se supone que tendr� sus ventajas;
    '   aunque si se siguen "las reglas" indicadas, no tendr�a que dar problemas.
    '   Adem�s la intenci�n de esta clase es formar parte de un componente (DLL)
    '   y el c�digo no estar�a disponible a las aplicaciones cliente...
    '   Por eso, te aconsejo que no hagas experimentos,
    '   si no sabes las consecuencias que esa pruebas pueden tener, el que avisa...
    '
    ' Ver el siguiente art�culo en la Knowledge Base de Microsoft para un ejemplo
    ' de un UserControl subclasificado usando punteros a objetos:
    '   HOWTO: Subclass a UserControl, Article ID: Q179398
    '
    Set mWSC(claseActual) = WSC
    
    ' Subclasificar la ventana, (form o control), pasada como par�metro y
    ' guardar el procedimiento anterior
    With mWSC(claseActual)
        .hWnd = unControl.hWnd
        .PrevWndProc = SetWindowLong(.hWnd, GWL_WNDPROC, AddressOf WndProc)

    End With

End Sub

Public Sub unHook(ByVal WSC As clsMouse)

    ' Des-subclasificar la clase indicada
    Static claseActual As Long
    
    ' Buscar el �ndice de esta clase en el array
    claseActual = IndiceClase(WSC.hWnd)
    
    ' Si ya estaba subclasificada esta clase
    If claseActual Then

        With mWSC(claseActual)
            ' Restaurar la funci�n anterior de procesamiento de mensajes
            Call SetWindowLong(.hWnd, GWL_WNDPROC, .PrevWndProc)
            ' Poner a cero el indicador de que se est� usando
            .hWnd = 0&

        End With
        
        ' Quitar la referencia a esta clase
        Set mWSC(claseActual) = Nothing
        
        ' Si es la �ltima del array...
        If mnWSC = claseActual Then
            ' Eliminar este item y ajustar el array
            mnWSC = mnWSC - 1

            ' Si no hay m�s, eliminar el array
            If mnWSC = 0 Then
                Erase mWSC
            Else
                ' Ajustar el n�mero de elementos del array
                ReDim Preserve mWSC(1 To mnWSC)

            End If

        End If

    End If

End Sub

Private Function IndiceClase(ByVal elhWnd As Long, _
                             Optional ByRef Libre As Long = 0) As Long

    ' Este procedimiento buscar� el �ndice de la clase que tiene el hWnd indicado
    ' Tambi�n, si se especifica, devolver� el �ndice de una clase que est� libre.
    ' Nota: Es importante que el �ltimo par�metro sea por referencia,
    '       ya que en �l se devolver� el valor del �ndice libre.
    '
    Static i As Long
    
    IndiceClase = 0
    
    ' Recorrer todo el array
    For i = 1 To mnWSC

        With mWSC(i)

            ' Si coinciden los hWnd, es que ya se est� usando una subclasificaci�n
            If .hWnd = elhWnd Then

                ' usar esta misma clase
                ' pero si el hWnd es cero, ser� uno libre
                If elhWnd = 0 Then
                    Libre = i
                Else
                    IndiceClase = i

                End If

                Exit For
                
                ' Comprobar si hay alg�n "hueco" en el array,
                ' por ejemplo de una clase previamente liberada.
                ' Hay que tener en cuenta que estos procedimientos est�n en un BAS
                ' y sus valores se mantienen entre varias llamadas a las clases.
            ElseIf .hWnd = 0& Then
                Libre = i

            End If

        End With

    Next

End Function

