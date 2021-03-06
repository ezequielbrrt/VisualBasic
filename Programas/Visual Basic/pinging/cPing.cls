VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Clase para hacer ping desde Visual Basic"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales que contienen valores de propiedad
Private mvarIPDestino As String 'copia local
Private mvarLongitudDatos As Long 'copia local
Private mvarTimeOut As Long 'copia local
Private mvarEstado As Long 'copia local
Private mvarDescripcion As String 'copia local
Private mvarTiempo As Long 'copia local

'c�digos de error
Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Const PING_TIMEOUT = 200
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
'estructuras
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
Private ICMPOPT As ICMP_OPTIONS
Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

'funciones del api
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, _
    ByVal RequestData As String, ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Function AddressStringToLong(ByVal tmp As String) As Long
Dim i As Integer
Dim parts(1 To 4) As String
i = 0
'extrae cada parte del string '123.456.789.123
'delimitado por puntos
While InStr(tmp, ".") > 0
    i = i + 1
    parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
    tmp = Mid(tmp, InStr(tmp, ".") + 1)
Wend
i = i + 1
parts(i) = tmp
If i <> 4 Then
    AddressStringToLong = 0
    Exit Function
End If
'build the long value out of the  'hex of the extracted strings
AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
                         Right("00" & Hex(parts(3)), 2) & _
                         Right("00" & Hex(parts(2)), 2) & _
                         Right("00" & Hex(parts(1)), 2))
End Function


Private Function HiByte(ByVal wParam As Integer)
HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
LoByte = wParam And &HFF&
End Function

Private Function SocketsCleanup() As Boolean
Dim X As Long
X = WSACleanup()
If X <> 0 Then
    MsgBox "Windows Sockets error " & Trim$(Str$(X)) & _
           " occurred in Cleanup.", vbExclamation
    SocketsCleanup = False
Else
    SocketsCleanup = True
End If
End Function

Private Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
Dim X As Integer
Dim szLoByte As String, szHiByte As String, szBuf As String
X = WSAStartup(WS_VERSION_REQD, WSAD)
If X <> 0 Then
    MsgBox "Windows Sockets for 32 bit Windows " & _
           "environments is not successfully responding."
    SocketsInitialize = False
    Exit Function
End If
If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
    HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
    szHiByte = Trim$(Str$(HiByte(WSAD.wVersion)))
    szLoByte = Trim$(Str$(LoByte(WSAD.wVersion)))
    szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
    szBuf = szBuf & " is not supported by Windows " & _
                        "Sockets for 32 bit Windows environments."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
End If
If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
    szBuf = "This application requires a minimum of " & _
    Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
End If
SocketsInitialize = True
End Function



Public Sub Ping()
Attribute Ping.VB_UserMemId = 0
    Dim hPort As Long
    Dim dwAddress As Long
    Dim sDataToSend As String
    Dim iOpt As Long
    Dim ECHO As ICMP_ECHO_REPLY
    Dim res As Boolean
    
    res = SocketsInitialize
    If res Then
        sDataToSend = String$(mvarLongitudDatos, "A")
        dwAddress = AddressStringToLong(mvarIPDestino)
        hPort = IcmpCreateFile()
        If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), _
                        0, ECHO, Len(ECHO), mvarTimeOut) Then
            'the ping succeeded,.Status will be 0
            '.RoundTripTime is the time in ms for the ping to complete,
            '.Data is the data returned (NULL terminated)
            '.Address is the Ip address that actually replied
            '.DataSize is the size of the string in .Data
             mvarEstado = 0
             mvarTiempo = ECHO.RoundTripTime
        Else
            mvarTiempo = 0
            mvarEstado = Abs(ECHO.status)
            If mvarEstado = 0 Then mvarEstado = IP_GENERAL_FAILURE
        End If
        Call IcmpCloseHandle(hPort)
        res = SocketsCleanup
Else
    mvarEstado = IP_GENERAL_FAILURE
    mvarTiempo = 0
End If
mvarDescripcion = GetStatusCode(mvarEstado)
End Sub

Private Function GetStatusCode(status As Long) As String
Dim msg As String

Select Case status
    Case IP_SUCCESS:               msg = ""
    Case IP_BUF_TOO_SMALL:         msg = "buffer demasiado peque�o"
    Case IP_DEST_NET_UNREACHABLE:  msg = "red de destino no encontrada"
    Case IP_DEST_HOST_UNREACHABLE: msg = "host destino no encontrado"
    Case IP_DEST_PROT_UNREACHABLE: msg = "dest prot unreachable"
    Case IP_DEST_PORT_UNREACHABLE: msg = "puerto destino no encontrado"
    Case IP_NO_RESOURCES:          msg = "sin recursos"
    Case IP_BAD_OPTION:            msg = "opci�n inv�lida"
    Case IP_HW_ERROR:              msg = "error hardware"
    Case IP_PACKET_TOO_BIG:        msg = "paquete demasiado grande"
    Case IP_REQ_TIMED_OUT:         msg = "timeout"
    Case IP_BAD_REQ:               msg = "respuesta incorrecta"
    Case IP_BAD_ROUTE:             msg = "ruta inv�lida"
    Case IP_TTL_EXPIRED_TRANSIT:   msg = "ttl finalizado"
    Case IP_TTL_EXPIRED_REASSEM:   msg = "ttl expired reassem"
    Case IP_PARAM_PROBLEM:         msg = "error en par�metros"
    Case IP_SOURCE_QUENCH:         msg = "source quench"
    Case IP_OPTION_TOO_BIG:        msg = "opci�n demasiado grande"
    Case IP_BAD_DESTINATION:       msg = "destino incorrecto"
    Case IP_ADDR_DELETED:          msg = "addr deleted"
    Case IP_SPEC_MTU_CHANGE:       msg = "spec mtu change"
    Case IP_MTU_CHANGE:            msg = "ip mtu_change"
    Case IP_UNLOAD:                msg = "unload"
    Case IP_ADDR_ADDED:            msg = "addr added"
    Case IP_GENERAL_FAILURE:       msg = "fallo general"
    Case IP_PENDING:               msg = "pendiente"
    Case PING_TIMEOUT:             msg = "ping timeout"
    Case Else:                     msg = "recibido mensaje desconocido"
End Select
GetStatusCode = msg
End Function


Public Property Get Tiempo() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Tiempo
    Tiempo = mvarTiempo
End Property





Public Property Get Descripcion() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Descripcion
    Descripcion = mvarDescripcion
End Property





Public Property Get Estado() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.Estado
    Estado = mvarEstado
End Property



Public Property Let Timeout(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.TimeOut = 5
    mvarTimeOut = vData
    If mvarTimeOut < 1 Then mvarTimeOut = 1
End Property


Public Property Get Timeout() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.TimeOut
    Timeout = mvarTimeOut
End Property



Public Property Let LongitudDatos(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.LongitudDatos = 5
    mvarLongitudDatos = vData
    If mvarLongitudDatos > 250 Then mvarLongitudDatos = 250
    If mvarLongitudDatos < 1 Then mvarLongitudDatos = 1
End Property


Public Property Get LongitudDatos() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.LongitudDatos
    LongitudDatos = mvarLongitudDatos
End Property



Public Property Let IPDestino(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignaci�n.
'Syntax: X.IPDestino = 5
    mvarIPDestino = vData
End Property


Public Property Get IPDestino() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignaci�n.
'Syntax: Debug.Print X.IPDestino
    IPDestino = mvarIPDestino
End Property



Private Sub Class_Initialize()
mvarLongitudDatos = 32
mvarTimeOut = 1500
End Sub


