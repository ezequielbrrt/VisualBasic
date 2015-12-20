VERSION 5.00
Begin VB.Form frmDisplay 
   Caption         =   "Display"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   1095
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   Picture         =   "frmDisplay.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   6810
   Begin VB.CommandButton cmdpuerto 
      Caption         =   "PUERTO"
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txthexa 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdhexa 
      Caption         =   "Valor"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4800
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdReiniciar 
      Caption         =   "&Reiniciar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3120
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame fraSegmentos 
      Caption         =   "Segmentos"
      Height          =   1575
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   4095
      Begin VB.CheckBox chkG 
         Caption         =   "G"
         Height          =   300
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   900
      End
      Begin VB.CheckBox chkF 
         Caption         =   "F"
         Height          =   300
         Left            =   2760
         TabIndex        =   9
         Top             =   360
         Width           =   900
      End
      Begin VB.CheckBox chkE 
         Caption         =   "E"
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   900
      End
      Begin VB.CheckBox chkD 
         Caption         =   "D"
         Height          =   300
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   900
      End
      Begin VB.CheckBox chkC 
         Caption         =   "C"
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   900
      End
      Begin VB.CheckBox chkB 
         Caption         =   "B"
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
      Begin VB.CheckBox chkA 
         Caption         =   "A"
         Height          =   300
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame fratipo 
      Caption         =   "Tipo"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
      Begin VB.OptionButton opt2 
         Caption         =   "Catodo Comun"
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Anodo Comun"
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Valor Hexadecimal"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Shape shpg 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   195
      Left            =   705
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   705
   End
   Begin VB.Shape shpf 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   900
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   195
   End
   Begin VB.Shape shpe 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   900
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   195
   End
   Begin VB.Shape shpd 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   195
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   900
   End
   Begin VB.Shape shpc 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   900
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   195
   End
   Begin VB.Shape shpb 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   900
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   195
   End
   Begin VB.Shape shpa 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   195
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public h, sa, sb, sc, sd, se, sf, sg, total As Integer
Private Sub chkA_Click()

If chkA.Value = Checked Then
    shpa.BackColor = vbRed
    h = h + 128
Else: chkA.Value = Unchecked
    shpa.BackColor = &HC0FFFF
    h = h - 128
End If

End Sub

Private Sub chkB_Click()

If chkB.Value = Checked Then
    shpb.BackColor = vbRed
    h = h + 64
Else: chkB.Value = Unchecked
    shpb.BackColor = &HC0FFFF
    h = h - 64
End If

End Sub

Private Sub chkC_Click()

If chkC.Value = Checked Then
    shpc.BackColor = vbRed
    h = h + 32
Else: chkC.Value = Unchecked
    shpc.BackColor = &HC0FFFF
    h = h - 32
End If

End Sub

Private Sub chkD_Click()

If chkD.Value = Checked Then
    shpd.BackColor = vbRed
    h = h + 16
Else: chkD.Value = Unchecked
    shpd.BackColor = &HC0FFFF
    h = h - 16
End If

End Sub

Private Sub chkE_Click()

If chkE.Value = Checked Then
    shpe.BackColor = vbRed
    h = h + 8
Else: chkE.Value = Unchecked
    shpe.BackColor = &HC0FFFF
    h = h - 8
End If

End Sub

Private Sub chkF_Click()

If chkF.Value = Checked Then
    shpf.BackColor = vbRed
    h = h + 4
Else: chkF.Value = Unchecked
    shpf.BackColor = &HC0FFFF
    h = h - 4
End If

End Sub

Private Sub chkG_Click()

If chkG.Value = Checked Then
    shpg.BackColor = vbRed
    h = h + 2
Else: chkG.Value = Unchecked
    shpg.BackColor = &HC0FFFF
    h = h - 2
End If

End Sub

Private Sub cmdhexa_Click()

If opt2.Value = True Then
txthexa.Text = Hex(h)
End If

If opt1.Value = True Then
txthexa.Text = Hex(h)
End If

End Sub

Private Sub cmdpuerto_Click()

If opt2.Value = True Then
Out &H378, h
End If

If opt1.Value = True Then
Out &H378, h
End If


End Sub

Private Sub cmdReiniciar_Click()

chkA.Value = Unchecked
chkB.Value = Unchecked
chkC.Value = Unchecked
chkD.Value = Unchecked
chkE.Value = Unchecked
chkF.Value = Unchecked
chkG.Value = Unchecked
txthexa.Text = " "

End Sub

Private Sub cmdSalir_Click()

End

End Sub

Private Sub opt1_Click()

If opt1.Value = True Then
    h = 0
End If

End Sub

Private Sub opt2_Click()

If opt2.Value = True Then
    h = 1
Else
    h = 0
End If

End Sub
