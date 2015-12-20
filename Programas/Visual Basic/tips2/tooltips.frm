VERSION 2.00
Begin Form FormTips 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Formulario para mostrar tips"
   ClientHeight    =   3345
   ClientLeft      =   1980
   ClientTop       =   1545
   ClientWidth     =   6405
   Height          =   3750
   Left            =   1920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   6405
   Top             =   1200
   Width           =   6525
   Begin PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   345
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6375
      TabIndex        =   12
      Top             =   3000
      Width           =   6405
      Begin Label LabelAbajo 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   50
         TabIndex        =   13
         Top             =   30
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5040
      Top             =   1800
   End
   Begin TextBox texto1 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Tag             =   "Un texto"
      Text            =   "Text1"
      Top             =   1860
      Width           =   1995
   End
   Begin CommandButton boton1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Tag             =   "Un botón"
      Top             =   1800
      Width           =   1395
   End
   Begin PictureBox Panel3D8 
      Align           =   1  'Align Top
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      Begin PictureBox btnImprimir 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   3600
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Tag             =   "Imprimir tema"
         Top             =   30
         Width           =   510
      End
      Begin PictureBox btnGrabar 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   5040
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Tag             =   "Grabar tema"
         Top             =   30
         Width           =   510
      End
      Begin PictureBox btnSalir 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   5760
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Tag             =   "Salir sin grabar"
         Top             =   30
         Width           =   510
      End
      Begin PictureBox btnPrimero 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   240
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         Tag             =   "Primer tema"
         Top             =   30
         Width           =   510
      End
      Begin PictureBox btnAnterior 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   960
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Tag             =   "Tema anterior"
         Top             =   30
         Width           =   510
      End
      Begin PictureBox btnSiguiente 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   1680
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Tag             =   "Siguiente tema"
         Top             =   30
         Width           =   510
      End
      Begin PictureBox btnUltimo 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   2400
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Tag             =   "Último tema"
         Top             =   30
         Width           =   510
      End
      Begin PictureBox btnBorrar 
         AutoSize        =   -1  'True
         Height          =   510
         Left            =   4320
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Tag             =   "Eliminar tema"
         Top             =   30
         Width           =   510
      End
   End
   Begin Label Burbuja 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Burbuja"
      Height          =   225
      Left            =   570
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   570
   End
End
'Attribute VB_Name = "FormTips"
'Attribute VB_Creatable = False
'Attribute VB_Exposed = False
Option Explicit

Sub boton1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja boton1
End Sub

Sub btnAnterior_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnAnterior
End Sub

Sub btnBorrar_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnBorrar
End Sub

Sub btnGrabar_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnGrabar
End Sub

Sub btnImprimir_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnImprimir
End Sub

Sub btnPrimero_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnPrimero
End Sub

Sub btnSalir_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnSalir
End Sub

Sub btnSiguiente_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnSiguiente
End Sub

Sub btnUltimo_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja btnUltimo
End Sub

Sub Form_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
QuitarTip
End Sub

Sub Panel3D8_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
QuitarTip
End Sub

Sub QuitarTip ()
Burbuja.Visible = False
LabelAbajo.Visible = Burbuja.Visible
Timer1.Enabled = False
End Sub

Sub SacaBurbuja (boton As Control)
Dim cartel As String
If Timer1.Enabled Then Exit Sub
Burbuja = boton.Tag + " "
If boton.Left + Burbuja.Width > Me.Width Then
   Burbuja.Move boton.Left + boton.Width - Burbuja.Width, panel3d8.Top + boton.Top + boton.Height * 1.2
Else
   Burbuja.Move boton.Left, panel3d8.Top + boton.Top + boton.Height * 1.2
End If
Burbuja.Visible = True
Burbuja.ZOrder 'En realidad, al ser un label, no tiene mayor influencia
LabelAbajo = Burbuja.Caption
LabelAbajo.Visible = Burbuja.Visible
Timer1.Enabled = True
End Sub

Sub texto1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja texto1
End Sub

Sub Timer1_Timer ()
Static Ya As Integer
If Ya Then
   Ya = False
   Burbuja.Visible = False
End If
Ya = Not Ya
LabelAbajo.Visible = Burbuja.Visible
End Sub

