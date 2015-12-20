VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   2415
   ClientTop       =   1560
   ClientWidth     =   9540
   Height          =   6225
   Left            =   2355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9540
   Top             =   1215
   Width           =   9660
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
   Begin SSPanel Panel3D8 
      ForeColor       =   &H00FF0000&
      Height          =   580
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin SSCommand btnImprimir 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   5520
         Outline         =   0   'False
         Picture         =   FORM1.FRX:0000
         TabIndex        =   8
         Tag             =   "Imprimir tema"
         Top             =   30
         Width           =   510
      End
      Begin SSCommand btnGrabar 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   7620
         Outline         =   0   'False
         Picture         =   FORM1.FRX:0302
         TabIndex        =   7
         Tag             =   "Grabar tema"
         Top             =   30
         Width           =   510
      End
      Begin SSCommand btnSalir 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   8760
         Outline         =   0   'False
         Picture         =   FORM1.FRX:0604
         TabIndex        =   6
         Tag             =   "Salir sin grabar"
         Top             =   30
         Width           =   510
      End
      Begin SSCommand btnPrimero 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   240
         Outline         =   0   'False
         Picture         =   FORM1.FRX:0906
         TabIndex        =   5
         Tag             =   "Primer tema"
         Top             =   30
         Width           =   510
      End
      Begin SSCommand btnAnterior 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   960
         Outline         =   0   'False
         Picture         =   FORM1.FRX:0C08
         TabIndex        =   4
         Tag             =   "Tema anterior"
         Top             =   30
         Width           =   510
      End
      Begin SSCommand btnSiguiente 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   1680
         Outline         =   0   'False
         Picture         =   FORM1.FRX:0F0A
         TabIndex        =   3
         Tag             =   "Siguiente tema"
         Top             =   30
         Width           =   510
      End
      Begin SSCommand btnUltimo 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   2400
         Outline         =   0   'False
         Picture         =   FORM1.FRX:120C
         TabIndex        =   2
         Tag             =   "Último tema"
         Top             =   30
         Width           =   510
      End
      Begin SSCommand btnBorrar 
         AutoSize        =   2  'Adjust Button Size To Picture
         BevelWidth      =   0
         Height          =   510
         Left            =   6600
         Outline         =   0   'False
         Picture         =   FORM1.FRX:150E
         TabIndex        =   1
         Tag             =   "Eliminar tema"
         Top             =   30
         Width           =   510
      End
   End
   Begin Label Burbuja 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Burbuja"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Width           =   1035
   End
End
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
Burbuja.Visible = False
End Sub

Sub Panel3D8_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
Burbuja.Visible = False
End Sub

Sub SacaBurbuja (boton As Control)
Dim cartel As String
If Burbuja.Visible And Burbuja.Caption = boton.Tag Then Exit Sub
cartel = boton.Tag
Burbuja = cartel
Burbuja.Width = TextWidth(cartel)
If cartel = "Salir sin grabar" Then
    Burbuja.Move boton.Left + boton.Width - Burbuja.Width, panel3d8.Top + boton.Top + boton.Height * 1.1
Else
    Burbuja.Move boton.Left + boton.Height * .75, panel3d8.Top + boton.Top + boton.Height * 1.2
End If
Burbuja.Visible = True
Burbuja.ZOrder
End Sub

Sub Text1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja texto1
End Sub

Sub texto1_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
SacaBurbuja texto1
End Sub

