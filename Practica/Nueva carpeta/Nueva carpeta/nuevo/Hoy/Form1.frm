VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.VScrollBar vsb1 
      Height          =   2655
      Left            =   4800
      Max             =   255
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Color de Fondo"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Color de Texto"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      DragIcon        =   "Form1.frx":89088
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Text            =   "Olguin Sanchez Ruben"
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Option1_Click()

End Sub

Private Sub cmd_Click()
MsgBox "Hasta Pronto", vbOKOnly + vbCritical, "Adios"
Beep
End
End Sub

Private Sub opt_Click()




End Sub

Private Sub opt1_Click()
txt1.BackColor = RGB(100, vsb1.Value, 50)

End Sub

Private Sub opt2_Click()
txt1.ForeColor = RGB(vsb1.Value, 200, 50)




End Sub




Private Sub vsb1_Change()
If opt1.Value = True Then
txt1.ForeColor = RGB(vsb1.Value, 200, 50)
Else
txt1.BackColor = RGB(100, vsb1.Value, 100)
End If


End Sub
