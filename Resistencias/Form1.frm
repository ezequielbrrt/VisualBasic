VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   6450
   ClientTop       =   2730
   ClientWidth     =   7200
   FillStyle       =   0  'Solid
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   7200
   Begin VB.ComboBox cbo6 
      Height          =   315
      ItemData        =   "Form1.frx":19C52
      Left            =   720
      List            =   "Form1.frx":19C68
      TabIndex        =   11
      Text            =   "Temperatura"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cbo5 
      Height          =   315
      ItemData        =   "Form1.frx":19C94
      Left            =   720
      List            =   "Form1.frx":19CB6
      TabIndex        =   10
      Text            =   "Franja 4"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cbo7 
      Height          =   315
      ItemData        =   "Form1.frx":19D04
      Left            =   720
      List            =   "Form1.frx":19D11
      TabIndex        =   9
      Text            =   "Bandas"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.ComboBox cbo4 
      Height          =   315
      ItemData        =   "Form1.frx":19D33
      Left            =   720
      List            =   "Form1.frx":19D49
      TabIndex        =   3
      Text            =   "Tolerancia"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ComboBox cbo3 
      Height          =   315
      ItemData        =   "Form1.frx":19D75
      Left            =   720
      List            =   "Form1.frx":19D8E
      TabIndex        =   2
      Text            =   "Multiplicador "
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ComboBox cbo2 
      Height          =   315
      ItemData        =   "Form1.frx":19DC5
      Left            =   720
      List            =   "Form1.frx":19DE7
      TabIndex        =   1
      Text            =   "Franja 2"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ComboBox cbo1 
      Height          =   315
      ItemData        =   "Form1.frx":19E35
      Left            =   720
      List            =   "Form1.frx":19E54
      TabIndex        =   0
      Text            =   "Franja 1"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lbltemp1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   20
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbltemp 
      Caption         =   "Temperatura"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lbl12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   18
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lbl11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lbl10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lbl9 
      Caption         =   "&Tolerancia"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lbl8 
      Caption         =   "&Resistencia"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lbl6 
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl5 
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape shp6 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   4560
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape shp5 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1680
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   360
      X2              =   1320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   5160
      X2              =   6120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080C0FF&
      Height          =   975
      Left            =   1320
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H000080FF&
      Height          =   975
      Left            =   4920
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0FF&
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H008080FF&
      Height          =   1215
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3375
   End
   Begin VB.Shape shp4 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3960
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape shp3 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3120
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape shp2 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   2640
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape shp1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   2160
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public N1, N2, N3, N4, R As Double
Private Sub Command1_Click()
End
End Sub

Private Sub cbo1_Click()
Select Case cbo1.Text

Case "Cafe"
shp1.FillColor = RGB(55, 5, 25)
lbl1.Caption = "1"
N1 = 1

Case "Rojo"
shp1.FillColor = &HC0&
lbl1.Caption = "2"
N1 = 2

Case "Naranja"
shp1.FillColor = &H80FF&
lbl1.Caption = "3"
N1 = 3

Case "Amarillo"
shp1.FillColor = vbYellow
lbl1.Caption = "4"
N1 = 4

Case "Verde"
shp1.FillColor = &HC000&
lbl1.Caption = "5"
N1 = 5

Case "Azul"
shp1.FillColor = &HC0C000
lbl1.Caption = "6"
N1 = 6

Case "Violeta"
shp1.FillColor = &HFF80FF
lbl1.Caption = "7"
N1 = 7

Case "Gris"
shp1.FillColor = &HC0C0C0
lbl1.Caption = "8"
N1 = 8

Case "Blanco"
shp1.FillColor = vbWhite
lbl1.Caption = "9"
N1 = 9

End Select

End Sub


Private Sub cbo2_Click()
Select Case cbo2.Text

Case "Negro"
shp2.FillColor = vbBlack
lbl2.Caption = "0"
N2 = N1 * 10

Case "Cafe"
shp2.FillColor = &H404080
lbl2.Caption = "1"
N2 = N1 * 10 + 1

Case "Rojo"
shp2.FillColor = &HC0&
lbl2.Caption = "2"
N2 = N1 * 10 + 2

Case "Naranja"
shp2.FillColor = &H80FF&
lbl2.Caption = "3"
N2 = N1 * 10 + 3

Case "Amarillo"
shp2.FillColor = vbYellow
lbl2.Caption = "4"
N2 = N1 * 10 + 4

Case "Verde"
shp2.FillColor = &HC000&
lbl2.Caption = "5"
N2 = N1 * 10 + 5

Case "Azul"
shp2.FillColor = &HFFFF80
lbl2.Caption = "6"
N2 = N1 * 10 + 6

Case "Violeta"
shp2.FillColor = &HFF80FF
lbl2.Caption = "7"
N2 = N1 * 10 + 7

Case "Gris"
shp2.FillColor = &HC0C0C0
lbl2.Caption = "8"
N2 = N1 * 10 + 8

Case "Blanco"
shp2.FillColor = vbWhite
lbl2.Caption = "9"
N2 = N1 * 10 + 9

End Select

End Sub

Private Sub cbo3_Click()

Select Case cbo3.Text

Case "Negro"
shp3.FillColor = vbBlack
lbl3.Caption = "1"
N3 = N2 * 1

Case "Cafe"
shp3.FillColor = &H404080
lbl3.Caption = "10"
N3 = N2 * 10

Case "Rojo"
shp3.FillColor = &HC0&
lbl3.Caption = "100"
N3 = N2 * 100

Case "Naranja"
shp3.FillColor = &H80FF&
lbl3.Caption = "1000"
N3 = N2 * 1000

Case "Amarillo"
shp3.FillColor = vbYellow
lbl3.Caption = "10000"
N3 = N2 * 10000


Case "Verde"
shp3.FillColor = &HC000&
lbl3.Caption = "100000"
N3 = N2 * 100000

Case "Azul"
shp3.FillColor = &HFFFF80
lbl3.Caption = "1000000"
N3 = N2 * 1000000

End Select
If N3 >= 1000 And N3 < 1000000 Then
R = N3 / 1000
lbl10.Caption = R
lbl12.Caption = "KOhms"

ElseIf N3 >= 1000000 Then
R = N3 / 1000000
lbl10.Caption = R
lbl12.Caption = "MOhms"

Else
lbl10.Caption = N3
lbl12.Caption = "Ohms"

End If

End Sub


Private Sub cbo4_Click()

Select Case cbo4.Text

Case "Verde"
shp4.FillColor = &HC000&
lbl4.Caption = "0.5%"
lbl11.Caption = "0.5%"

Case "Cafe"
shp4.FillColor = &H404080
lbl4.Caption = "1%"
lbl11.Caption = "1%"

Case "Rojo"
shp4.FillColor = &HC0&
lbl4.Caption = "2%"
lbl11.Caption = "2%"

Case "Dorado"
shp4.FillColor = &HC0C0&
lbl4.Caption = "5%"
lbl11.Caption = "5%"

Case "Plata"
shp4.FillColor = &HC0C0C0
lbl4.Caption = "10%"
lbl11.Caption = "10%"

Case "Nada"
shp4.FillColor = 0
lbl4.Caption = "20%"
lbl11.Caption = "20%"
End Select

End Sub


Private Sub cbo5_Click()
Select Case cbo5.Text

Case "Negro"
shp5.FillColor = vbBlack
lbl5.Caption = "0"
N4 = N1 * 100

Case "Cafe"
shp5.FillColor = &H404080
lbl5.Caption = "1"
N4 = N1 * 100 + 1

Case "Rojo"
shp5.FillColor = &HC0&
lbl5.Caption = "2"
N4 = N1 * 100 + 2

Case "Naranja"
shp5.FillColor = &H80FF&
lbl5.Caption = "3"
N3 = N1 * 100 + 3

Case "Amarillo"
shp5.FillColor = vbYellow
lbl5.Caption = "4"
N3 = N1 * 100 + 4

Case "Verde"
shp5.FillColor = &HC000&
lbl5.Caption = "5"
N4 = N1 * 10 + 5

Case "Azul"
shp5.FillColor = &HFFFF80
lbl5.Caption = "6"
N4 = N1 * 10 + 6

Case "Violeta"
shp5.FillColor = &HFF80FF
lbl5.Caption = "7"
N4 = N1 * 100 + 7

Case "Gris"
shp5.FillColor = &HC0C0C0
lbl5.Caption = "8"
N4 = N1 * 100 + 8

Case "Blanco"
shp5.FillColor = vbWhite
lbl5.Caption = "9"
N4 = N1 * 100 + 9

End Select
End Sub

Private Sub cbo6_Click()
Select Case cbo6.Text

Case "Verde"
shp6.FillColor = &HC000&
lbl6.Caption = "20°"
lbltemp1.Caption = "20°"

Case "Cafe"
shp6.FillColor = &H404080
lbl6.Caption = "30°"
lbltemp1.Caption = "30°"

Case "Rojo"
shp6.FillColor = &HC0&
lbl6.Caption = "40°"
lbltemp1.Caption = "40°"

Case "Dorado"
shp6.FillColor = &HC0C0&
lbl6.Caption = "50°"
lbltemp1.Caption = "50°"

Case "Plata"
shp6.FillColor = &HC0C0C0
lbl6.Caption = "60°"
lbltemp1.Caption = "60°"

Case "Nada"
shp6.FillColor = 0
lbl6.Caption = "100°"
lbltemp1.Caption = "100°"
End Select
End Sub

Private Sub cbo7_Click()

Select Case cbo7.Text

Case "4 Bandas"
cbo5.Visible = False
cbo6.Visible = False
shp5.Visible = False
shp6.Visible = False
lbl5.Visible = False
lbl6.Visible = False
lbltemp.Visible = False
lbltemp1.Visible = False


Case "5 Bandas"
cbo5.Visible = True
cbo6.Visible = False
shp5.Visible = True
shp6.Visible = False
lbl5.Visible = True
lbl6.Visible = False
lbltemp.Visible = False
lbltemp1.Visible = False


Case "6 Bandas"
cbo5.Visible = True
cbo6.Visible = True
shp5.Visible = True
shp6.Visible = True
lbl5.Visible = True
lbl6.Visible = True
lbltemp.Visible = True
lbltemp1.Visible = True


End Select

End Sub

Private Sub cmdSalir_Click()
MsgBox "Hasta Pronto", vbOKOnly + vbCritical, "Adios"
End
End Sub

Private Sub Image1_Click()

End Sub

