VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   2655
   ClientTop       =   2520
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   Picture         =   "led.frx":0000
   ScaleHeight     =   5370
   ScaleWidth      =   4155
   Begin VB.CommandButton cmd3 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.HScrollBar vsb1 
      Height          =   375
      Left            =   600
      Max             =   500
      TabIndex        =   2
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Off"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "On"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "INTENSIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Shape shp2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   1560
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape shp1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   1680
      Y1              =   1320
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   2160
      X2              =   2160
      Y1              =   1320
      Y2              =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd1_Click()
shp1.FillColor = vbRed
shp2.FillColor = vbRed
Out &H378, 1


End Sub

Private Sub cmd2_Click()
shp1.FillColor = vbBlack
shp2.FillColor = vbBlack
Out &H378, 0


End Sub

Private Sub cmd3_Click()
MsgBox "Hasta Pronto", vbOKOnly + vbCritical, "Adios"
Beep
End
End Sub

Private Sub HScroll1_Change()


shp1.FillColor = RGB(vsb1.Value, 200, 50)
shp2.FillColor = RGB(vsb1.Value, 200, 50)


shp1.BackColor = RGB(100, vsb1.Value, 100)



End Sub

Private Sub hsb1_Change()

End Sub

Private Sub vsb1_Change()
shp1.FillColor = RGB(vsb1.Value, 0, 0)
shp2.FillColor = RGB(vsb1.Value, 0, 0)
Out &H378, 1









End Sub
