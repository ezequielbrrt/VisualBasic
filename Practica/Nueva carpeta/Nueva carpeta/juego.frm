VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "juego.frx":0000
   ScaleHeight     =   5145
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6120
      Top             =   4200
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FF0000&
      Caption         =   "&Tirar"
      Height          =   735
      Left            =   2040
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Image imgdado2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   360
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image imgdado1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   6120
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
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
      Left            =   3960
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
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
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
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
      Left            =   1920
      TabIndex        =   0
      Top             =   120
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
MsgBox "GANASTE", vbOKOnly + 48, "Mensaje"

Else
MsgBox "PERDISTE", vbOKOnly + 16, "Mensaje"
lbl2.Caption = ""
lbl3.Caption = ""

If numero1 = 1 Then
imgdado2.Picture = LoadPicture("E:\dado 1.jpg")
numero2 = &H1

ElseIf numero1 = 2 Then
imgdado2.Picture = LoadPicture("E:\dado 2.jpg")
numero2 = &H8

ElseIf numero1 = 3 Then
imgdado2.Picture = LoadPicture("E:\dado 3.jpg")
numero2 = &H9

ElseIf numero1 = 4 Then
imgdado2.Picture = LoadPicture("E:\dado 4.jpg")
numero2 = &HC

ElseIf numero1 = 5 Then
imgdado2.Picture = LoadPicture("E:\dado 5.jpg")
numero2 = &HD

ElseIf numero1 = 6 Then
imgdado2.Picture = LoadPicture("E:\dado 6.jpg")
numero2 = &HE

End If


If numero2 = 1 Then
imgdado1.Picture = LoadPicture("E:\dado 1.jpg")
dsalida = numero2 + &H10


ElseIf numero2 = 2 Then
imgdado1.Picture = LoadPicture("E:\dado 2.jpg")
numero2 = numero2 + &H80
ElseIf numero2 = 3 Then

imgdado1.Picture = LoadPicture("E:\dado 3.jpg")
numero2 = numero2 + &H90

ElseIf numero2 = 4 Then

imgdado1.Picture = LoadPicture("E:\dado 4.jpg")
numero2 = numero2 + &HC0

ElseIf numero2 = 5 Then

imgdado1.Picture = LoadPicture("E:\dado 5.jpg")
numero2 = numero2 + &HD0

ElseIf numero2 = 6 Then

imgdado1.Picture = LoadPicture("E:\dado 6.jpg")
numero2 = numero2 + &HE0



End If

Out &H378, numero2
End If


End Sub

Private Sub cmd2_Click()
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Image1_Click()
If numero1 = 1 Then

End Sub

Private Sub Image2_Click()

End Sub

Private Sub img1_Click()

End Sub

Private Sub img_Click()

End Sub

Private Sub imd_Click()

End Sub

Private Sub lbl2_Click()
End If

End Sub

Private Sub lbl3_Click()
End If

End Sub

Private Sub Timer1_Timer()
lbl1.Caption = "JUGAR AHORA"

End Sub
