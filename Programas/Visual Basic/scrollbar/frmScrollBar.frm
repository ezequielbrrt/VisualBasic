VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Uso del ScrollBar"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   0
      ScaleHeight     =   5565
      ScaleWidth      =   4365
      TabIndex        =   1
      Top             =   0
      Width           =   4365
      Begin VB.CommandButton cmdEjemplos 
         Caption         =   "Más ejemplos"
         Height          =   405
         Left            =   450
         TabIndex        =   3
         Top             =   4920
         Width           =   1380
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   405
         Left            =   2250
         TabIndex        =   4
         Top             =   4920
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Este es un ejemplo del uso del control ScrollBar para mover controles en un formulario."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4620
         Left            =   165
         TabIndex        =   2
         Top             =   195
         Width           =   3930
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3190
      LargeChange     =   250
      Left            =   4380
      SmallChange     =   150
      TabIndex        =   0
      Top             =   0
      Width           =   293
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEjemplos_Click()
'Muestra la ventana de ejemplos
Form2.Show
End Sub

Private Sub Command1_Click()
'Cierra la ventana
Unload Me
End Sub

Private Sub Form_Load()
'Se establece como máximo valor al ScrollBar
'la resta entre la altura del formulario y
'la altura del PictureBox
VScroll1.Max = Picture1.ScaleHeight - Me.ScaleHeight
End Sub

Private Sub VScroll1_Change()
'Cuando el ScrollBar cambia de posición
'se mueve el PictureBox según el valor
'del ScrollBar: A más valor el Picture se
'mueve hacia arriba para dejar visible su
'parte inferior.
Picture1.Move 0, -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
'Cuando el ScrollBar se está deslizando
'se mueve el PictureBox según el valor
'del ScrollBar: A más valor el Picture se
'mueve hacia arriba para dejar visible su
'parte inferior.
Picture1.Move 0, -VScroll1.Value
End Sub
