VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   Picture         =   "programas for.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd4 
      Caption         =   "&Habilitar"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "&Habilitar"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "&Habilitar"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "&Habilitar"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox cmb4 
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Text            =   "Tolerancia"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmb3 
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Text            =   "Linea 3"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmb2 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Text            =   "Linea 2 "
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Text            =   "Linea 1"
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Dim i As Integer
For i = 0 To 1 Step 1


cmb1.AddItem ("Negro")

cmb1.AddItem ("Cafe")

cmb1.AddItem ("Rojo")

cmb1.AddItem ("Amarillo")

cmb1.AddItem ("Verde")

cmb1.AddItem ("Azul")

cmb1.AddItem ("Violeta")

cmb1.AddItem ("Gris")

cmb1.AddItem ("Blanco")
Next i

End Sub

Private Sub cmd2_Click()
Dim i2 As Integer

For i2 = 0 To 1 Step 1
If i2 = 1 Then
cmb2.AddItem ("Negro")

cmb2.AddItem ("Cafe")

cmb2.AddItem ("Rojo")

cmb2.AddItem ("Amarillo")

cmb2.AddItem ("Verde")

cmb2.AddItem ("Azul")

cmb2.AddItem ("Violeta")

cmb2.AddItem ("Gris")

cmb2.AddItem ("Blanco")


End If
Next i2
End Sub

Private Sub cmd3_Click()
Dim i3 As Integer

For i3 = 1 To 10 Step 1
Select Case i3
Case 1
    cmb3.AddItem ("Negro")
Case 2
    cmb3.AddItem ("Cafe")
Case 3
    cmb3.AddItem ("Rojo")
Case 4
    cmb3.AddItem ("Amarillo")
Case 5
    cmb3.AddItem ("Verde")
Case 6
    cmb3.AddItem ("Azul")
Case 7
    cmb3.AddItem ("Violeta")
Case 8
    cmb3.AddItem ("Gris")
Case 9
    cmb3.AddItem ("Blanco")
End Select
Next i3

End Sub

Private Sub cmd4_Click()
Dim i4 As Integer

For i4 = 1 To 10 Step 1
Select Case i4
Case 1
    cmb4.AddItem ("Verde")
Case 2
    cmb4.AddItem ("Cafe")
Case 3
    cmb4.AddItem ("Rojo")
Case 4
    cmb4.AddItem ("Oro")
Case 5
    cmb4.AddItem ("Plata")
    
Case 6
    cmb4.AddItem ("Nada")

End Select
Next i4


End Sub
