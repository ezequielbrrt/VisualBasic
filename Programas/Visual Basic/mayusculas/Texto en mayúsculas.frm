VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Texto en MAYÚSCULAS"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   435
      Left            =   1545
      TabIndex        =   2
      Top             =   1283
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1298
      TabIndex        =   1
      Top             =   728
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "&Escribe cualquier texto en la caja de texto:"
      Height          =   345
      Left            =   795
      TabIndex        =   0
      Top             =   173
      Width           =   3090
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
'Texto en MAYÚSCULAS                         '
''''''''''''''''''''''''''''''''''''''''''''''
'Ejemplo que muestra como se puede usar las  '
'funciones UCase, Asc, Chr, los eventos      '
'KeyPress y la variables Ascii para          '
'transformar el texto escrito a mayúsculas   '
'Esta es una mejora de la otra versión ya que'
'llamaba a la propiedad Text1.Text varias    '
'veces y en procesadores lentos se notaba el '
'cambio de minúscula a mayúscula. Ahora se   '
'preprocesa la variable KeyAscii para que    '
'el cambio sea más rápido.                   '
''''''''''''''''''''''''''''''''''''''''''''''
'Escrito por Gustavo Alegre                  '
'Junio del 2003                              '
'Descargado de Visual Basic siglo XII        '
'www.vbsiglo21.com                           '
''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
