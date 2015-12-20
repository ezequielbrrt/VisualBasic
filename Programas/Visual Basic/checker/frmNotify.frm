VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-mail checker"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4665
   Icon            =   "frmNotify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock wsock 
      Left            =   240
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.PictureBox TrayIcon 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   1260
      Picture         =   "frmNotify.frx":030A
      ScaleHeight     =   555
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdOpenEmail 
      Caption         =   "Ver e-mail"
      Height          =   435
      Left            =   3480
      TabIndex        =   1
      Top             =   1740
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   2280
      TabIndex        =   0
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Timer tmrCheck 
      Interval        =   60000
      Left            =   720
      Top             =   1680
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   60
      Width           =   2595
   End
   Begin VB.Image imgNewMail 
      Height          =   675
      Left            =   360
      Top             =   540
      Width           =   675
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   540
      Width           =   2745
   End
   Begin VB.Menu mnuOptions 
      Caption         =   " "
      Enabled         =   0   'False
      Begin VB.Menu mnuOptionsCheckNow 
         Caption         =   "Chequear ahora"
      End
      Begin VB.Menu mnuOptionsExecutemail 
         Caption         =   "Ejecutar programa mail"
      End
      Begin VB.Menu mnuOptionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsConfigurar 
         Caption         =   "Configurar..."
      End
      Begin VB.Menu mnuOptionsHabilitado 
         Caption         =   "Habilitado"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsAbout 
         Caption         =   "Acerca de ..."
      End
      Begin VB.Menu mnuOptionsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsCerrar 
         Caption         =   "Cerrar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'e-checker
'Programa para chequear e-mail de una cuenta pop3
'al iniciarse se coloca en el system tray
'
'por Julio Daniel Moreyra
'21/07/98
'
'Este programa y su código fuente es freeware
'se puede usar y modificar libremente, si
'se cita el autor (aunque sea en los comentarios).
'
Option Explicit
Dim result As Long
Dim Response As String
Dim TimeToCheck As Integer
Dim ShowAlert As Boolean

'Codigo tomado de Brian Harper
'www.brianharper.demon.co.uk
'Gracias Brian !!
'Pone el icono del programa en el system tray
Private Sub ShowProgramInTray()
    NI.cbSize = Len(NI) 'set the length of this structure
    NI.hwnd = TrayIcon.hwnd 'control to receive messages from
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP 'operation flags
    NI.uCallbackMessage = WM_MOUSEMOVE 'recieve messages from mouse activities
    TrayIcon.Picture = LoadResPicture(20, vbResIcon)
    NI.hIcon = TrayIcon.Picture  'the location of the icon to display
    NI.szTip = "No hay mensajes nuevos" + Chr$(0) 'the tool tip to display
    result = Shell_NotifyIconA(NIM_ADD, NI) 'add the icon to the system tray
End Sub
'Cambia el icono y el tip en el system tray
Private Sub ShowIconInTray(NroIcon As Integer, msg As String)
    NI.szTip = msg + Chr(0)
    TrayIcon.Picture = LoadResPicture(NroIcon, vbResIcon)
    NI.hIcon = TrayIcon.Picture
    result = Shell_NotifyIconA(NIM_MODIFY, NI) 'add the icon to the system tray
End Sub
'Espera sincronicamente por una respuesta
'asigna a Respuesta lo leido del socket
Function WaitFor(ResponseCode As String, Respuesta As String) As Boolean
    Dim start As Single, Tmr As Single
    Static nIcon As Integer
    
    If nIcon = 0 Then nIcon = 40
    start = Timer ' Controlar que no sea forever
    
    While Len(Response) = 0
        Tmr = Timer - start
    
        DoEvents '** IMPORTANTE: Dejar que el sistema siga
        'Aviso en el system tray, para que no se desespere el usuario
        ShowIconInTray nIcon, "E-checker:esperando respuesta del servidor"
        If Tmr > Val(Timeout) Then  ' Time in seconds to wait
           'MsgBox "POP3 service error, timeout esperando respuesta del servidor", vbExclamation
           Exit Function
        End If
        
        Sleep 200       'espero un 2/10 de segundo
        nIcon = nIcon + 10 'cambio icono
        If nIcon > 70 Then nIcon = 40
          
    Wend

    Respuesta = Response
    Response = "" ' **IMPORTANTE: poner en blanco
    WaitFor = True
End Function

'Lee la configuracion del registry
Private Sub LeerConfiguracion()
    
    pop3Host = GetSetting(App.EXEName, "Config", "Host")
    Do While pop3Host = ""
        If pop3Host = "" Then
            MsgBox "Debe configurar el programa", vbExclamation
            frmConfigurar.Show 1
        End If
    Loop
    
    pop3User = GetSetting(App.EXEName, "Config", "User")
    pop3Passwd = GetSetting(App.EXEName, "Config", "Passwd")
    Interval = GetSetting(App.EXEName, "Config", "Interval")
    EmailProgram = GetSetting(App.EXEName, "Config", "Program")
    Arguments = GetSetting(App.EXEName, "Config", "Arguments")
    Timeout = GetSetting(App.EXEName, "Config", "TimeOut", "30")
    
End Sub
'Ya vio el aviso - ocultar formulario
Private Sub cmdAceptar_Click()
    result = SetWindowPos(frmMain.hwnd, -2, 0, 0, 0, 0, 3)
    frmMain.Visible = False
End Sub
'Llamar al programa de e-mail
Private Sub cmdOpenEmail_Click()
    mnuOptionsExecutemail_Click
End Sub

Private Sub Form_Load()
    
    ShowProgramInTray    'mostrar el icono en el System Tray
    App.TaskVisible = False
    LeerConfiguracion    'leer seteos del programa
    mnuOptionsCheckNow_Click 'chequear, que para eso estamos
    
End Sub
'Sacar el icono del system tray
Private Sub Form_Unload(Cancel As Integer)

    result = Shell_NotifyIconA(NIM_DELETE, NI) 'removes the icon from the tray

End Sub

'About del programa
Private Sub mnuOptionsAbout_Click()
    frmAbout.Show 1
End Sub
'Salir del programa
Private Sub mnuOptionsCerrar_Click()
    SaveSetting App.EXEName, "Config", "Host", pop3Host
    SaveSetting App.EXEName, "Config", "User", pop3User
    SaveSetting App.EXEName, "Config", "Passwd", pop3Passwd
    SaveSetting App.EXEName, "Config", "Interval", Interval
    SaveSetting App.EXEName, "Config", "Program", EmailProgram
    SaveSetting App.EXEName, "Config", "Arguments", Arguments
    SaveSetting App.EXEName, "Config", "Timeout", Timeout
    Unload Me
End Sub
'Configurar las opciones de chequeo
Private Sub mnuOptionsConfigurar_Click()
    frmConfigurar.Show 1
    'Por si se cambio el intervalo de chequeo
    TimeToCheck = Val(Interval)
End Sub
'A chequear se ha dicho
Private Sub mnuOptionsCheckNow_Click()
    Dim Respuesta As String
    Dim cantmensajes As String
    
    On Error GoTo errsock
    wsock.RemoteHost = pop3Host
    wsock.RemotePort = POP3Port
    wsock.LocalPort = 0
    'De otra forma no es posible chequear a menos que pasen
    '4 minutos entre aperturas y cierres de sockets
    'esto es una "caracteristica" de diseño del control
    wsock.Connect
    
    If Not WaitFor("+OK", Respuesta) Then
        MsgBox "El servidor de correo no contesta", vbCritical
        ShowIconInTray 30, "e-checker: el servidor de correo no contesta"
        wsock.Close
        Exit Sub
    End If
    wsock.SendData "USER " & pop3User + vbCrLf
    If Not WaitFor("+OK", Respuesta) Then
        MsgBox "El usuario POP3 es inválido", vbCritical
        ShowIconInTray 30, "e-checker:usuario POP3 inválido"
        wsock.Close
        Exit Sub
    End If
    wsock.SendData "PASS " & pop3Passwd + vbCrLf
    If Not WaitFor("+OK", Respuesta) Then
        MsgBox "El password del usuario POP3 es inválido", vbCritical
        ShowIconInTray 30, "e-checker:password POP3 inválido"
        wsock.Close
        Exit Sub
    End If
    wsock.SendData "STAT" + vbCrLf
    If Not WaitFor("+OK", Respuesta) Then
        MsgBox "El servidor no responde al comando STAT", vbCritical
        ShowIconInTray 30, "e-checker: no puede ejecutar STAT"
        wsock.Close
        Exit Sub
    End If
    cantmensajes = Mid$(Respuesta, 5, InStr(5, Respuesta, " ", vbTextCompare) - 5)
    lblMsg(0).Caption = "Tiene " + cantmensajes + " mensajes nuevos."
    lblMsg(1).Caption = Format$(Now, "General Date")
    imgNewMail.Picture = LoadResPicture(IIf(cantmensajes > 0, 80, 90), vbResIcon)
    If Val(cantmensajes) > 0 Then
        ShowIconInTray 10, lblMsg(0).Caption
    Else
        ShowIconInTray 20, lblMsg(0).Caption
    End If
    wsock.SendData "QUIT" + vbCrLf
    wsock.Close
    TimeToCheck = Val(Interval)
    'Si fue por que transcurrio el tiempo, o hay mensajes
    If ShowAlert Or cantmensajes > 0 Then
        tmrCheck.Enabled = False
        frmMain.Visible = True
        result = SetWindowPos(frmMain.hwnd, -1, 0, 0, 0, 0, 3)
        tmrCheck.Enabled = True
    End If
    ShowAlert = True
    Exit Sub
    
errsock:
    MsgBox Err.Description, vbCritical
    ShowIconInTray 30, "e-checker: problemas en conexión"
    wsock.Close
    Exit Sub

End Sub
'Llamar al programa de e-mail
Private Sub mnuOptionsExecutemail_Click()
    Dim rc As Double
    
    On Error Resume Next
    If EmailProgram <> "" Then
        Screen.MousePointer = vbHourglass
        rc = Shell(EmailProgram + " " + Arguments, vbMaximizedFocus)
        Screen.MousePointer = vbNormal
        If rc = 0 Then
            MsgBox "Hubo un error al llamar al programa" + Chr(13) + "de e-mail. Verifique el path", vbExclamation
        End If
    End If
End Sub
'Habilita / deshabilita el timer
Private Sub mnuOptionsHabilitado_Click()
    
    mnuOptionsHabilitado.Checked = Not mnuOptionsHabilitado.Checked
    tmrCheck.Enabled = mnuOptionsHabilitado.Checked
    
End Sub
'Verifico cuando llega el momento
'de chequear el mail
Private Sub tmrCheck_Timer()
        
    TimeToCheck = TimeToCheck - 1
    If TimeToCheck = 0 Then
        ShowAlert = False
        mnuOptionsCheckNow_Click
    End If
    
End Sub
'Captura de los mensajes del mouse
Private Sub Trayicon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim msg As Long
    msg = (x And &HFF) * &H100

    Select Case msg
        Case 0 'mouse moves
        
        Case &HF00  'left mouse button down
        
        Case &H1E00 'left mouse button up
        
        Case &H3C00  'right mouse button down
        PopupMenu mnuOptions 'show the popoup menu
        Case &H2D00 'left mouse button double click
        mnuOptionsCheckNow_Click
        Case &H4B00 'right mouse button up
        
        Case &H5A00 'right mouse button double click
        
    End Select
   
End Sub

Private Sub wsock_Connect()
    
    'MsgBox "Conexion establecida con el servidor", vbInformation
    
End Sub

Private Sub wsock_DataArrival(ByVal bytesTotal As Long)
    
    wsock.GetData Response
    
End Sub
