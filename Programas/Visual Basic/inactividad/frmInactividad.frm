VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de inactividad"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   1605
      TabIndex        =   2
      Top             =   2625
      Width           =   1470
   End
   Begin VB.TextBox Text1 
      Height          =   1620
      Left            =   375
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   690
      Width           =   3930
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   420
      Top             =   2565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "lblCoordenadas"
      Height          =   255
      Left            =   1343
      TabIndex        =   0
      ToolTipText     =   "Coordenadas del cursor del mouse."
      Top             =   195
      Width           =   1995
   End
   Begin VB.Menu mnuEjemplo 
      Caption         =   "Ejemplo de menú (&Archivo)"
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnuGuardarComo 
         Caption         =   "G&uardar como"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
'Control de inactividad                      '
''''''''''''''''''''''''''''''''''''''''''''''
'Ejemplo que muestra como se puede usar la   '
'API de Windows y la propiedad KeyPreview    '
'para detectar si se está trabajando en el   '
'programa o no (similar al funcionamiento del'
'protector de pantalla de Windows).          '
''''''''''''''''''''''''''''''''''''''''''''''
'Escrito por Gustavo Alegre                  '
'Junio del 2003                              '
'Descargado de Visual Basic siglo XII        '
'www.vbsiglo21.com                           '
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

Dim pt32 As POINTAPI
Dim ptx As Long
Dim pty As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim contador As Integer

Private Sub cmdSalir_Click()
End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
contador = 0
End Sub

Private Sub mnuAbrir_Click()
MsgBox "Hiciste click en el menú Abrir", vbInformation
End Sub

Private Sub mnuGuardar_Click()
MsgBox "Hiciste click en el menú Guardar", vbInformation
End Sub

Private Sub mnuGuardarComo_Click()
MsgBox "Hiciste click en el menú Guardar como", vbInformation
End Sub

Private Sub mnuImprimir_Click()
MsgBox "Hiciste click en el menú Imprimir", vbInformation
End Sub

Private Sub mnuNuevo_Click()
MsgBox "Hiciste click en el menú Nuevo", vbInformation
End Sub

Private Sub mnuSalir_Click()
If MsgBox("Hiciste click en el menú Salir" & vbCrLf & vbCrLf & "...pero, ¿Realmente deseas salir?", vbQuestion + vbYesNo) = vbYes Then End
End Sub

Private Sub Timer1_Timer()
ptx = pt32.X
pty = pt32.Y
Call GetCursorPos(pt32)
Label1.Caption = pt32.X & " " & pt32.Y
If pt32.X = ptx And pt32.Y = pty Then
    contador = contador + 1
    If contador = 50 Then
        MsgBox "Cada 5 segundos aparecera este mensaje", vbInformation
        contador = 0
    End If
Else
    contador = 0
End If
End Sub

