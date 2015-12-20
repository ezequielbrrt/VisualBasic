VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "juego.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   4080
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FF0000&
      Caption         =   "&Tirar"
      Height          =   735
      Left            =   1560
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Juego de Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Dim numero1, suma, numero2 As Integer
cmd2.Enabled = True
Randomize
numero1 = CInt(Rnd() * 5) + 1
numero2 = CInt(Rnd() * 5) + 1
lbl2.Caption = numero1
lbl3.Caption = numero2
suma = numero1 + numero2

If suma = 11 Then
MsgBox "GANASTE", vbOKOnly + 48, "mensaje"

Else
MsgBox "PERDISTE", vbOKOnly + 16, "mensaje"
lbl2.Caption = ""
lbl3.Caption = ""
End If


End Sub

Private Sub cmd2_Click()
End
End Sub

Private Sub Timer1_Timer()
lbl1.Caption = "JUGAR AHORA"

End Sub
