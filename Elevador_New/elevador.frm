VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   3750
   ClientTop       =   1905
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   Picture         =   "elevador.frx":0000
   ScaleHeight     =   7635
   ScaleMode       =   0  'User
   ScaleWidth      =   7095
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Timer tmrPB 
      Interval        =   100
      Left            =   4560
      Top             =   7080
   End
   Begin VB.Timer tmrP3 
      Interval        =   100
      Left            =   4080
      Top             =   7080
   End
   Begin VB.Timer tmrP2 
      Interval        =   100
      Left            =   3600
      Top             =   7080
   End
   Begin VB.Timer tmrP1 
      Interval        =   100
      Left            =   3120
      Top             =   7080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   2160
      Top             =   3960
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "&Planta Baja"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H000080FF&
      Caption         =   "&Piso1"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "&Piso 2"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00FF0000&
      Caption         =   "&Piso 3"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox pct1 
      Height          =   7215
      Left            =   5160
      ScaleHeight     =   7155
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.VScrollBar vsb1 
      Height          =   4335
      LargeChange     =   3
      Left            =   240
      Max             =   15
      TabIndex        =   1
      Top             =   1200
      Value           =   15
      Width           =   375
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Shape shpp2 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape shpp1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shppb 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   3360
      Width           =   495
   End
   Begin VB.Shape shpp3 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shp6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   375
   End
   Begin VB.Shape shp7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   375
   End
   Begin VB.Shape shp8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   375
   End
   Begin VB.Shape shp9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   375
   End
   Begin VB.Shape shp11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape shp12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape shp13 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape shp14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape shp16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape shp17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape shp18 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   615
   End
   Begin VB.Shape shp26 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape shp27 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape shp28 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape shp29 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape shp30 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape shp31 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape shp32 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape shp33 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape shp34 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape shp35 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape shp19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape shp21 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape shp22 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape shp23 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape shp24 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape shp5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape shp10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   375
   End
   Begin VB.Shape shp15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape shp20 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   375
   End
   Begin VB.Shape shp25 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape shp1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape shp2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape shp3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   375
   End
   Begin VB.Shape shp4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   375
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoReset 
         Caption         =   "Reset"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoGuion 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivoSalir 
         Caption         =   "&Salir"
         Shortcut        =   +^{F4}
      End
   End
   Begin VB.Menu mnuCreditos 
      Caption         =   "&Creditos"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
MsgBox "Hasta Pronto", vbOKOnly + vbCritical, "Adios"
Beep
End
End Sub

Private Sub cmd2_Click()
txt1.Text = "3"
vsb1.Value = 0
tmrP3.Enabled = True
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0

shpp1.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H80FF&
shppb.FillColor = &H0&

pct1.Line (0, 0)-(40, 25), &HFF0000, B
pct1.Line (0, 26)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &H8000000F, B

End Sub

Private Sub cmd3_Click()
txt1.Text = "2"
vsb1.Value = 5
tmrP2.Enabled = True
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0


shpp1.FillColor = &H0&
shpp2.FillColor = &H80FF&
shpp3.FillColor = &H0&
shppb.FillColor = &H0&

pct1.Line (0, 0)-(40, 25), &H8000000F, B
pct1.Line (0, 25)-(40, 50), &HFF0000, B
pct1.Line (0, 51)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &H8000000F, B


End Sub

Private Sub cmd4_Click()
txt1.Text = "1"
tmrP1.Enabled = True
vsb1.Value = 10
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0


shpp1.FillColor = &H80FF&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&
shppb.FillColor = &H0&

pct1.Line (0, 0)-(40, 25), &H8000000F, B
pct1.Line (0, 25)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &HFF0000, B
pct1.Line (0, 76)-(40, 100), &H8000000F, B

End Sub

Private Sub cmd5_Click()
txt1.Text = "PB"
vsb1.Value = 15
tmrPB.Enabled = True
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0


shpp1.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&
shppb.FillColor = &H80FF&

pct1.Line (0, 0)-(40, 25), &H8000000F, B
pct1.Line (0, 25)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &HFF0000, B

End Sub

Private Sub Form_Load()
pct1.Scale (0, 0)-(40, 100)

End Sub

Private Sub mnuArchivoReset_Click()
txt1.Text = "PB"
vsb1.Value = 15
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0
tmrPB.Enabled = True

shpp1.FillColor = &H0&
shppb.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&

pct1.Line (0, 0)-(40, 25), &H8000000F, B
pct1.Line (0, 25)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &HFF0000, B

End Sub

Private Sub mnuArchivoSalir_Click()

MsgBox "Hasta Pronto", vbOKOnly + vbCritical, "Adios"
Beep
End
End Sub

Private Sub mnuCreditos_Click()
MsgBox "Hecho por Olguin Sanchez & Pacheco Altamirano 5IM7", vbOKOnly + 48, "Elaborado Por..."
End Sub

Private Sub pct1_Paint()

pct1.FillColor = &HFF0000
pct1.DrawWidth = 4

pct1.Line (0, 0)-(40, 25), &HFF0000, B

End Sub

Private Sub Timer1_Timer()
shp5.FillColor = &HFF&
shp10.FillColor = &HFF&
shp15.FillColor = &HFF&
shp20.FillColor = &HFF&
shp25.FillColor = &HFF&
shp30.FillColor = &HFF&
shp35.FillColor = &HFF&
shp2.FillColor = &HFF&
shp3.FillColor = &HFF&
shp4.FillColor = &HFF&
shp17.FillColor = &HFF&
shp18.FillColor = &HFF&
shp19.FillColor = &HFF&
shp32.FillColor = &HFF&
shp33.FillColor = &HFF&
shp34.FillColor = &HFF&
shp6.FillColor = &HFF&
shp11.FillColor = &HFF&
Timer1.Enabled = False

shppb.FillColor = &H80FF&
shpp1.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&

End Sub


Private Sub tmrP1_Timer()
tmrP1.Enabled = False
shp24.FillColor = &HFF&
shp27.FillColor = &HFF&
shp2.FillColor = &HFF&
shp28.FillColor = &HFF&
shp7.FillColor = &HFF&
shp12.FillColor = &HFF&
shp17.FillColor = &HFF&
shp22.FillColor = &HFF&
shp32.FillColor = &HFF&

shpp1.FillColor = &H80FF&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&
shppb.FillColor = &H0&
End Sub

Private Sub tmrP2_Timer()
tmrP2.Enabled = False
shp1.FillColor = &HFF&
shp2.FillColor = &HFF&
shp3.FillColor = &HFF&
shp4.FillColor = &HFF&
shp5.FillColor = &HFF&
shp10.FillColor = &HFF&
shp15.FillColor = &HFF&
shp20.FillColor = &HFF&
shp16.FillColor = &HFF&
shp17.FillColor = &HFF&
shp18.FillColor = &HFF&
shp19.FillColor = &HFF&
shp26.FillColor = &HFF&
shp21.FillColor = &HFF&
shp31.FillColor = &HFF&
shp32.FillColor = &HFF&
shp33.FillColor = &HFF&
shp34.FillColor = &HFF&
shp35.FillColor = &HFF&

shpp1.FillColor = &H0&
shpp2.FillColor = &H80FF&
shpp3.FillColor = &H0&
shppb.FillColor = &H0&
End Sub

Private Sub tmrP3_Timer()
tmrP3.Enabled = False
shp1.FillColor = &HFF&
shp6.FillColor = &HFF&
shp11.FillColor = &HFF&
shp16.FillColor = &HFF&
shp21.FillColor = &HFF&
shp26.FillColor = &HFF&
shp31.FillColor = &HFF&
shp2.FillColor = &HFF&
shp3.FillColor = &HFF&
shp4.FillColor = &HFF&
shp5.FillColor = &HFF&
shp17.FillColor = &HFF&
shp18.FillColor = &HFF&
shp19.FillColor = &HFF&
shp20.FillColor = &HFF&
shp32.FillColor = &HFF&
shp33.FillColor = &HFF&
shp34.FillColor = &HFF&
shp35.FillColor = &HFF&

shpp1.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H80FF&
shppb.FillColor = &H0&

End Sub

Private Sub tmrPB_Timer()

shp5.FillColor = &HFF&
shp10.FillColor = &HFF&
shp15.FillColor = &HFF&
shp20.FillColor = &HFF&
shp25.FillColor = &HFF&
shp30.FillColor = &HFF&
shp35.FillColor = &HFF&
shp17.FillColor = &HFF&
shp18.FillColor = &HFF&
shp19.FillColor = &HFF&
shp21.FillColor = &HFF&
shp26.FillColor = &HFF&
shp34.FillColor = &HFF&
shp33.FillColor = &HFF&
shp32.FillColor = &HFF&
tmrPB.Enabled = False
Timer1.Enabled = True

shpp1.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&
shppb.FillColor = &H80FF&

End Sub

Private Sub txt1_Change()
txt1.FontSize = 15
txt1.FontBold = True

End Sub

Private Sub vsb1_Change()
Select Case vsb1.Value

Case 0
txt1.Text = "3"
tmrP3.Enabled = True
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0

shpp1.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H80FF&
shppb.FillColor = &H0&

pct1.Line (0, 0)-(40, 25), &HFF0000, B
pct1.Line (0, 26)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &H8000000F, B

pct1.Line (0, 0)-(40, 25), &HFF0000, B
pct1.Line (0, 26)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &H8000000F, B

Case 5
txt1.Text = "2"
tmrP2.Enabled = True
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0


shpp1.FillColor = &H0&
shpp2.FillColor = &H80FF&
shpp3.FillColor = &H0&
shppb.FillColor = &H0&
pct1.Line (0, 0)-(40, 25), &H8000000F, B
pct1.Line (0, 25)-(40, 50), &HFF0000, B
pct1.Line (0, 51)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &H8000000F, B

Case 10
txt1.Text = "1"
tmrP1.Enabled = True
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0


shpp1.FillColor = &H80FF&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&
shppb.FillColor = &H0&
pct1.Line (0, 0)-(40, 25), &H8000000F, B
pct1.Line (0, 25)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &HFF0000, B
pct1.Line (0, 76)-(40, 100), &H8000000F, B

Case 15
txt1.Text = "PB"
tmrPB.Enabled = True
shp1.FillColor = &HE0E0E0
shp2.FillColor = &HE0E0E0
shp3.FillColor = &HE0E0E0
shp4.FillColor = &HE0E0E0
shp5.FillColor = &HE0E0E0
shp6.FillColor = &HE0E0E0
shp7.FillColor = &HE0E0E0
shp8.FillColor = &HE0E0E0
shp9.FillColor = &HE0E0E0
shp10.FillColor = &HE0E0E0
shp11.FillColor = &HE0E0E0
shp12.FillColor = &HE0E0E0
shp13.FillColor = &HE0E0E0
shp14.FillColor = &HE0E0E0
shp15.FillColor = &HE0E0E0
shp16.FillColor = &HE0E0E0
shp17.FillColor = &HE0E0E0
shp18.FillColor = &HE0E0E0
shp19.FillColor = &HE0E0E0
shp20.FillColor = &HE0E0E0
shp21.FillColor = &HE0E0E0
shp22.FillColor = &HE0E0E0
shp23.FillColor = &HE0E0E0
shp24.FillColor = &HE0E0E0
shp25.FillColor = &HE0E0E0
shp26.FillColor = &HE0E0E0
shp27.FillColor = &HE0E0E0
shp28.FillColor = &HE0E0E0
shp29.FillColor = &HE0E0E0
shp30.FillColor = &HE0E0E0
shp31.FillColor = &HE0E0E0
shp32.FillColor = &HE0E0E0
shp33.FillColor = &HE0E0E0
shp34.FillColor = &HE0E0E0
shp35.FillColor = &HE0E0E0


shpp1.FillColor = &H0&
shpp2.FillColor = &H0&
shpp3.FillColor = &H0&
shppb.FillColor = &H80FF&
pct1.Line (0, 0)-(40, 25), &H8000000F, B
pct1.Line (0, 25)-(40, 50), &H8000000F, B
pct1.Line (0, 50)-(40, 75), &H8000000F, B
pct1.Line (0, 75)-(40, 100), &HFF0000, B

End Select


End Sub
