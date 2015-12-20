VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Barra de Tareas (c)1997 J.LeVasseur"
   ClientHeight    =   1830
   ClientLeft      =   1560
   ClientTop       =   2775
   ClientWidth     =   3945
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "EjemploBT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1830
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picGancho 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1680
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   720
      Width           =   555
   End
   Begin VB.Menu mnuBar 
      Caption         =   ""
      Enabled         =   0   'False
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' EjemploBT ver1.0
' 1997 J.LeVasseur lvasseur@tiac.net a0@null.net
' Un ejemplo de Usar la barra de tareas en Win95/NT4
' El PictureBox picGancho sirve como gancho de los
' mensajes CallBack del API Shell_NotifyIcon. Tiene
' que ser un control con un hWnd. Todo lo interesante
' esta en el picGancho_MouseMove . Como pueden ver, un
' control MsgHook o MsgBlaster aqui sobra...
'---------------
Private Type TIPONOTIFICARICONO
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'------------------
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
'--------------------
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    pnid As TIPONOTIFICARICONO) As Boolean
'--------------------
Private Declare Function WinExec& Lib "kernel32" _
    (ByVal lpCmdLine As String, ByVal nCmdShow As Long)
'--------------------
Dim t As TIPONOTIFICARICONO


Private Sub Form_Click()
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then  ' Como tener "Cancel"
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        mnuAcerca_Click
        Unload Me
        End
    End If
'---------------------------------
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
'---------------------------------
    t.szTip = "Ejemplo de barra de tareas..." & Chr$(0) ' Es un string de "C" ( \0 )
    Shell_NotifyIcon NIM_ADD, t
    Me.Hide
    App.TaskVisible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub




Private Sub mnuAcerca_Click()
' Un consejo, mover un Form en estado minimizado
' da un GPF...
Dim ValDev As Long
With Form1
    picGancho.Picture = Me.Icon
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
    Show
End With
End Sub




Private Sub mnuSalir_Click(Index As Integer)
    Unload Me
End Sub

Private Sub picGancho_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, Msg As Long, ValDev As Long
    Msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                'MsgBox ("Rayma")
                 ValDev = WinExec("CONTROL.EXE DESK.CPL", 1) ' aca mi prg
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                 ' PopUp menu,2 significa Izq/Der botones en el menu, mnuAbout es BOLD
                 Me.PopupMenu mnuBar, 2, , , mnuAcerca
            End Select
        rec = False
    End If
End Sub


