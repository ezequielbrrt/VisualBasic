VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   5310
   ClientTop       =   2865
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   Picture         =   "menus de.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   5085
   Begin VB.CommandButton cmdBlanco2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdNegro2 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdAmarillo2 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdazul2 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdverde2 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdRojo2 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdBlanco 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdNegro 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdamarillo 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdAzul 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdVerde 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   1440
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdRed 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdDer 
      Caption         =   "Derecha"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCentro 
      Caption         =   "Centro"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdIzq 
      Caption         =   "Izquierda"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "menus de.frx":BD622
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label lbl2 
      Caption         =   "Color de Fondo"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Caption         =   "Color de Fuente"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoSalir 
         Caption         =   "Salir"
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Formato"
      Begin VB.Menu mnuFormEstilo 
         Caption         =   "Estilo"
         Begin VB.Menu mnuFormEstiloNeg 
            Caption         =   "Negrita"
         End
         Begin VB.Menu mnuFormEstiloCurs 
            Caption         =   "Cursiva"
         End
         Begin VB.Menu mnuFormEstiloSub 
            Caption         =   "Subrayado"
         End
      End
      Begin VB.Menu mnuFuente 
         Caption         =   "Fuente"
         Begin VB.Menu mnuFuenteArial 
            Caption         =   "Arial"
         End
         Begin VB.Menu mnuFuenteTNR 
            Caption         =   "Times New Roman"
         End
         Begin VB.Menu mnuFuenteCourier 
            Caption         =   "Courier"
         End
      End
      Begin VB.Menu mnuTamaño 
         Caption         =   "Tamaño"
         Begin VB.Menu mnuTamañoNumero 
            Caption         =   "10"
            Index           =   0
         End
         Begin VB.Menu mnuTamañoNumero 
            Caption         =   "11"
            Index           =   1
         End
         Begin VB.Menu mnuTamañoNumero 
            Caption         =   "12"
            Index           =   2
         End
         Begin VB.Menu mnuTamañoNumero 
            Caption         =   "13"
            Index           =   3
         End
         Begin VB.Menu mnuTamañoNumero 
            Caption         =   "14"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdamarillo_Click()
txt1.ForeColor = &HFFFF&
End Sub

Private Sub cmdAmarillo2_Click()
txt1.BackColor = &HFFFF&
End Sub

Private Sub cmdAzul_Click()
txt1.ForeColor = &HFF0000
End Sub

Private Sub cmdazul2_Click()
txt1.BackColor = &HFF0000
End Sub

Private Sub cmdBlanco_Click()
txt1.ForeColor = &HFFFFFF
End Sub

Private Sub cmdBlanco2_Click()
txt1.BackColor = &HFFFFFF
End Sub

Private Sub cmdcentro_Click()
txt1.Alignment = 2

End Sub

Private Sub cmdDer_Click()
txt1.Alignment = 1

End Sub

Private Sub cmdIzq_Click()
txt1.Alignment = 0
End Sub

Private Sub cmdNegro_Click()
txt1.ForeColor = &H0&
End Sub

Private Sub cmdNegro2_Click()
txt1.BackColor = &H0&
End Sub

Private Sub cmdRed_Click()
txt1.ForeColor = vbRed

End Sub

Private Sub cmdRojo2_Click()
txt1.BackColor = &HFF&
End Sub

Private Sub cmdVerde_Click()
txt1.ForeColor = &HFF00&
End Sub

Private Sub cmdverde2_Click()
txt1.BackColor = &HFF00&
End Sub

Private Sub mnuArchivoSalir_Click()
End
End Sub

Private Sub mnuFormEstiloCurs_Click()
If txt1.FontItalic = False Then
txt1.FontItalic = True
Else
txt1.FontItalic = False
End If
If mnuFormEstiloCurs.Checked = False Then

mnuFormEstiloCurs.Checked = True
Else
mnuFormEstiloCurs.Checked = False

End If


End Sub

Private Sub mnuFormEstiloNeg_Click()

If txt1.FontBold = False Then
txt1.FontBold = True
Else
txt1.FontBold = False
End If
If mnuFormEstiloNeg.Checked = False Then

mnuFormEstiloNeg.Checked = True
Else
mnuFormEstiloNeg.Checked = False

End If

End Sub

Private Sub mnuFormEstiloSub_Click()
If txt1.FontUnderline = False Then
txt1.FontUnderline = True
Else
txt1.FontUnderline = False
End If
If mnuFormEstiloSub.Checked = False Then

mnuFormEstiloSub.Checked = True
Else
mnuFormEstiloSub.Checked = False

End If


End Sub

Private Sub mnuFuenteArial_Click()
txt1.Font = Arial



End Sub

Private Sub mnuFuenteCourier_Click()

txt1.Font = Courier

End Sub

Private Sub mnuFuenteTNR_Click()
txt1.Font = "Times New Roman"

End Sub

Private Sub mnuTamañoNumero_Click(Index As Integer)

Select Case Index

    Case 0
    txt1.FontSize = 10


    
    Case 1
    txt1.FontSize = 11
     
    Case 2
    txt1.FontSize = 12
      
    Case 3
    txt1.FontSize = 13
    
    Case 4
     txt1.FontSize = 14
      
End Select
     End Sub
