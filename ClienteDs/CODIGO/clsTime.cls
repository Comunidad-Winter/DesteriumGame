VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' cGetTimer                                                         (16/Jun/98)
' Clase para calcular el tiempo entre eventos, etc.
'
' Ampliada para usar distintas funciones del API                    (29/Ago/98)
' Nuevos métodos/propiedades usando APIs especiales para NT/2000    (27/Oct/00)
' Comprobado en Windows XP Professional                             (09/Jul/03)
'
' Para usarla:
'   StartTimer          Iniciar el temporizador
'   StopTimer           Finalizar la cuenta
'   ElapsedTimer        Mostrar el tiempo transcurrido, formateado
'                       --- Nuevas propiedades ---
'   TimerType           El temporizador a usar
'   TotalElapsedTime    El número de milisegundos transcurridos
'
'   TimerPerformanceFrecuency
'   TimerPerformanceCounter
'
' ©Guillermo 'guille' Som, 1998-2003
'
' En esta clase se usa el tiempo del sistema, es decir el número de
' milisegundos transcurridos desde que se inició Windows.
' Este temporizador vuelve a cero una vez transcurridos 49.7 días,
' (2^32 milisegundos),
' por tanto el cálculo del tiempo transcurrido puede no ser exacto,
' dependiendo de cuándo se inició y terminó la cuenta.
'------------------------------------------------------------------------------
Option Explicit

'
'------------------------------------------------------------------------------
' Se puede usar timeGetTime si se quiere mayor precisión,
' ya que esta función tiene una precisión de:
' 1 milisegundo en Windows 95, y de
' 5 milisegundos o más, (es configurable), en Windows NT
'
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'
' Sólo para Windows NT/2000/XP
' Para averiguar si se dispone de temporizador de alta precisión
' y configurar la precisión de timeGetTime()
'
' Nota: El valor asignado a los parámetros Currency hay que multiplicarlo por 10000
Private Declare Function QueryPerformanceCounter Lib "kernel32" _
    (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" _
    (ByRef lpFrequency As Currency) As Long
'
' Texto basado en la descripción de:
'   High-Resolution Timer (Platform SDK / Windows User Interface)
' La función QueryPerformanceCounter devuelve el valor actual del temporizador
' de alta precisión (si es que el sistema tiene uno).
' Se puede usar para calcular intervalos de tiempo usando la precisión de
' un temporizador de alta precisión (o resolución).
' Por ejemplo, suponiendo que QueryPerformanceFrequency indicara que la frecuencia
' del temporizador de alta precisión es de 50000 por segundo. Si la aplicación
' llamara a QueryPerformanceCounter inmediatamente antes y justo después de la
' sección de código a temporizar, y que estos valores fuesen 1500 y 3500,
' esto indicaría que el tiempo transcurrido es de .04 segundos (2000 ciclos)
'
'
'Private Declare Function timeBeginPeriod Lib "winmm.dll" _
'    (ByVal uPeriod As Long) As Long
'Private Declare Function timeEndPeriod Lib "winmm.dll" _
'    (ByVal uPeriod As Long) As Long
'
'------------------------------------------------------------------------------
' La precisión de GetTickCount es (aproximadamente)
' de 10 milisegundos en Windows NT 3.5 o superior
' de 16 ms. en Windows NT 3.1
' de 55 ms. en Windows 95 o superior
'
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

'------------------------------------------------------------------------------
' Valores del tipo de timer a usar
Public Enum eTipoTimer
    eGetTickCount '= 0
    eTimeGetTime '= 1
End Enum

' Variables locales para los cálculos
'
Private mc_StartTimer As Currency
Private mc_EndTimer As Currency
Private mc_Frecuencia As Currency
'
Private m_TimerType As eTipoTimer
Private m_StartTimer As Long
Private m_EndTimer As Long
Private m_LastTime As Long      ' Último tiempo calculado           (11/Mar/99)

Public Function ElapsedTime(Optional Intervalo As Long = 1000) As String
    ' Mostrar el tiempo transcurrido
    '
    ' Intervalo se usará para la forma de mostrar el tiempo transcurrido
    '
    ' Detener el timer
    If m_EndTimer = 0 Then
        StopTimer
    End If
    '
    ' No admitir números negativos ni cero
    If Intervalo < 1 Then Intervalo = 1000
    '
    ' Devolver una cadena formateada con el tiempo
    ElapsedTime = Format$(CDbl(m_EndTimer) / Intervalo, "0.0####")
    ' si se quiere usar el TimerPerformanceCounter
    'ElapsedTime = Format$((mc_EndTimer / mc_Frecuencia) * 1000 / Intervalo, "0.0####")
End Function

Public Sub StopTimer()
    ' Detener el timer
    '
    ' El tiempo de diferencia puede no ser "exacto",
    ' ya que se reinicia cada 49.7 dias... (2^32 milisegundos)
    ' El timer se pone a cero cuando se inicia el sistema.
    If m_TimerType = eTimeGetTime Then
        'mc_EndTimer = TimerPerformanceCounter - mc_StartTimer
        m_EndTimer = timeGetTime() - m_StartTimer
    Else
        m_EndTimer = GetTickCount() - m_StartTimer
    End If
    '
    m_LastTime = m_EndTimer
End Sub

Public Sub StartTimer()
    ' Iniciar el timer
    m_EndTimer = 0
    If m_TimerType = eTimeGetTime Then
        'mc_Frecuencia = TimerPerformanceFrecuency
        'mc_StartTimer = TimerPerformanceCounter
        m_StartTimer = timeGetTime()
    Else
        m_StartTimer = GetTickCount()
    End If
End Sub

Public Property Get TimerType() As eTipoTimer
    ' Devolver el tipo de timer que se usa
    TimerType = m_TimerType
End Property

Public Property Let TimerType(ByVal vNewValue As eTipoTimer)
    ' asignar el tipo de timer a usar
    m_TimerType = vNewValue
End Property

Private Sub Class_Initialize()
    ' Valor por defecto
    'm_TimerType = eGetTickCount
    m_TimerType = eTimeGetTime      ' Este es más preciso               (11/Mar/99)
End Sub

Public Property Get TotalElapsedTime() As Long
    ' Mostrar el tiempo transcurrido en el último cálculo,
    ' incluso mientras se está calculando
    TotalElapsedTime = m_LastTime ' m_EndTimer
End Property

Public Property Get TimerPerformanceFrecuency() As Long
    ' Devuelve la frecuencia del temporizador de alta precisión     (27/Oct/00)
    '
    Dim curFrecuency As Currency
    Dim ret As Long
    '
    ret = QueryPerformanceFrequency(curFrecuency)
    If ret = 0 Then
        ' No hay un temporizador de alta precisión
        TimerPerformanceFrecuency = 0
    Else
        ' Devolver el valor de la frecuencia por cada segundo
        ' (multiplicar por 10000 ya que el valor devuelto viene con decimales)
        TimerPerformanceFrecuency = curFrecuency * 10000
    End If
End Property

Public Property Get TimerPerformanceCounter() As Currency
    ' Contador del temporizador de alta precisión actual            (27/Oct/00)
    ' Esta función se usará junto con TimerPerformanceFrecuency,
    ' para averiguar el lapso transcurrido entre dos periodos de tiempo
    '
    ' Los milisegundos transcurridos será:
    '   la diferencia entre dos llamadas a esta propiedad
    '   dividido por el valor indicado por TimerPerformanceFrecuency
    '
    ' Nota: Este resultado será casi igual que una llamada a timeGetTime,
    '       aunque algo más precisa.
    '
    Dim curFrecuency As Currency
    Dim ret As Long
    '
    ret = QueryPerformanceCounter(curFrecuency)
    If ret = 0 Then
        ' No hay un temporizador de alta precisión
        TimerPerformanceCounter = 0
    Else
        ' Devolver el valor del contador
        ' (multiplicar por 10000 ya que el valor devuelto viene con decimales)
        TimerPerformanceCounter = curFrecuency * 10000
    End If
End Property

