VERSION 5.00
Begin VB.Form frmConfigurar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar e-checker"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6000
   Icon            =   "frmConfigurar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   180
      TabIndex        =   20
      Top             =   60
      Width           =   4395
      Begin VB.OptionButton optBTN 
         Caption         =   "Programa"
         Height          =   735
         Index           =   3
         Left            =   2520
         MaskColor       =   &H00808000&
         Picture         =   "frmConfigurar.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Programa de e-mail"
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optBTN 
         Caption         =   "Intervalo"
         Height          =   735
         Index           =   2
         Left            =   1260
         MaskColor       =   &H00808000&
         Picture         =   "frmConfigurar.frx":0C8E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Intervalo de chequeo"
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optBTN 
         Caption         =   "Identidad"
         Height          =   735
         Index           =   1
         Left            =   0
         MaskColor       =   &H00808000&
         Picture         =   "frmConfigurar.frx":16C0
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Identidad POP3"
         Top             =   60
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   3480
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cerrar"
      Height          =   435
      Left            =   4740
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame fraBTN 
      Caption         =   "Servidor y buzón POP3"
      Height          =   3375
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   5715
      Begin VB.TextBox txtPop3Host 
         Height          =   315
         Left            =   2580
         TabIndex        =   0
         Top             =   1260
         Width           =   2355
      End
      Begin VB.TextBox txtPop3User 
         Height          =   315
         Left            =   2580
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1860
         Width           =   2055
      End
      Begin VB.TextBox txtPop3Passwd 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2580
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2460
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese aquí los datos para establecer la conexión POP3 con su servidor de correo electrónico."
         Height          =   495
         Index           =   3
         Left            =   300
         TabIndex        =   21
         Top             =   420
         Width           =   4635
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del servidor de correo o dirección IP:"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del buzón:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   1860
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Contraseña del buzón:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   2520
         Width           =   1635
      End
   End
   Begin VB.Frame fraBTN 
      Caption         =   "Programa E-mail"
      Height          =   3375
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   5715
      Begin VB.TextBox txtProgram 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   900
         Width           =   5415
      End
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar ..."
         Height          =   315
         Left            =   4500
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtArguments 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   4035
      End
      Begin VB.Label Label1 
         Caption         =   "Indique la ruta completa al programa de e-mail (por ejemplo C:\EUDORA\EUDORA.EXE)"
         Height          =   435
         Index           =   5
         Left            =   60
         TabIndex        =   19
         Top             =   420
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Opcional: indique los argumentos de la línea de comandos"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   4155
      End
   End
   Begin VB.Frame fraBTN 
      Caption         =   "Intervalos"
      Height          =   3375
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   5715
      Begin VB.TextBox txtTimeout 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "4"
         Top             =   2580
         Width           =   270
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "3"
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "Establezca el intervalo de tiempo que debe esperar e-checker por la respuesta del servidor"
         Height          =   615
         Index           =   10
         Left            =   300
         TabIndex        =   24
         Top             =   1740
         Width           =   3915
      End
      Begin VB.Label Label1 
         Caption         =   "segundos."
         Height          =   195
         Index           =   9
         Left            =   2580
         TabIndex        =   23
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "cada"
         Height          =   195
         Index           =   4
         Left            =   1740
         TabIndex        =   22
         Top             =   1140
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Establezca el intervalo de tiempo que debe transcurrir para que e-checker verifique la existencia de nuevo correo:"
         Height          =   675
         Index           =   8
         Left            =   360
         TabIndex        =   16
         Top             =   300
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "minutos."
         Height          =   195
         Index           =   7
         Left            =   2580
         TabIndex        =   15
         Top             =   1140
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmConfigurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Configuracion de e-checker
Option Explicit
'Leer los seteos de configuración del programa
Private Sub LeerSeteos()
    txtPop3Host = pop3Host
    txtPop3User = pop3User
    txtPop3Passwd = pop3Passwd
    txtInterval = Interval
    txtProgram = EmailProgram
    txtArguments = Arguments
    txtTimeout = Timeout
End Sub
'Guarda los seteos de configuración del programa
Private Sub GuardarSeteos()
    pop3Host = txtPop3Host
    pop3User = txtPop3User
    pop3Passwd = txtPop3Passwd
    Interval = txtInterval
    EmailProgram = txtProgram
    Arguments = txtArguments
    Timeout = txtTimeout
End Sub
'Chequeo que todos los campos esten completos
Private Function TodoOk() As Boolean
        
    'Host POP3
    If txtPop3Host = "" Then
        MsgBox "Debe indicar el nombre o direccion IP" + Chr(13) + "del servidor de correo", vbExclamation
        txtPop3Host.SetFocus
        fraBTN(1).Visible = True
        fraBTN(1).ZOrder (0)
        Exit Function
    End If
    'Usuario POP3
    If txtPop3User = "" Then
        MsgBox "Debe indicar el usuario POP3", vbExclamation
        txtPop3User.SetFocus
        fraBTN(1).Visible = True
        fraBTN(1).ZOrder (0)
        Exit Function
    End If
    'No chequeo que complete el password
    'puede haber una cuenta sin password
    If Val(txtInterval) = 0 Then
        MsgBox "Debe indicar un intervalo mayor que cero", vbExclamation
        txtInterval.SetFocus
        fraBTN(2).Visible = True
        fraBTN(2).ZOrder (0)
        Exit Function
    End If
    If Val(txtTimeout) = 0 Then
        MsgBox "Debe indicar un timeout mayor que cero", vbExclamation
        txtTimeout.SetFocus
        fraBTN(2).Visible = True
        fraBTN(2).ZOrder (0)
        Exit Function
    End If
    
    TodoOk = True
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub
'Busca la ruta del programa de e-mail
Private Sub cmdExaminar_Click()
    Dim ofn As OPENFILENAME
    Dim rtn As String
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "Todos los archivos" + Chr(0)
    ofn.lpstrFile = Space(254) + Chr(0)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254) + Chr(0)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = "c:\archiv~1" + Chr(0)
    ofn.lpstrTitle = "Programa de e-mail" + Chr(0)
    ofn.flags = OFNFileMustExist + OFNHideReadOnly + OFNPathMustExist
    
    rtn = GetOpenFileName(ofn)
    
    If rtn >= 1 Then
       txtProgram.Text = ofn.lpstrFile
    End If
    
End Sub
Private Sub cmdOK_Click()

    If TodoOk() Then
        GuardarSeteos
        Unload Me
    End If
    
End Sub
'Activo el frame visible de acuerdo al boton presionado
Private Sub optBTN_Click(Index As Integer)
    Dim i As Integer
    
    fraBTN(Index).Visible = True
    For i = 1 To 3
        If i <> Index Then fraBTN(i).Visible = False
    Next
    Select Case Index
        Case 1
            txtPop3Host.SetFocus
        Case 2
            txtInterval.SetFocus
        Case 3
            txtProgram.SetFocus
    End Select
    
End Sub
Private Sub Form_Load()
    'centra el formulario
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    LeerSeteos
    fraBTN(1).Visible = True
    fraBTN(1).ZOrder (0)
End Sub
Private Sub txtInterval_KeyPress(KeyAscii As Integer)

    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
    
End Sub
