VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   Picture         =   "display.frx":0000
   ScaleHeight     =   6105
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Selecciona"
      Height          =   3255
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdpuer 
         Caption         =   "&Puerto"
         Height          =   495
         Left            =   1440
         TabIndex        =   34
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdhex 
         Caption         =   "&mostrar"
         Height          =   495
         Left            =   480
         TabIndex        =   33
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd20 
         Caption         =   "&LIMPIAR"
         Height          =   495
         Left            =   3000
         TabIndex        =   32
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txthexa 
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check1"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check1"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check1"
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check1"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H8000000B&
         Caption         =   "Check1"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lbl19 
         Height          =   375
         Left            =   960
         TabIndex        =   28
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl18 
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl17 
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl16 
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl15 
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl14 
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl1 
         Caption         =   "&G"
         Height          =   495
         Left            =   600
         TabIndex        =   21
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl2 
         Caption         =   "&F"
         Height          =   495
         Left            =   960
         TabIndex        =   20
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl3 
         Caption         =   "&E"
         Height          =   495
         Left            =   1320
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl4 
         Caption         =   "&D"
         Height          =   495
         Left            =   1680
         TabIndex        =   18
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl5 
         Caption         =   "&C"
         Height          =   495
         Left            =   2040
         TabIndex        =   17
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl6 
         Caption         =   "&B"
         Height          =   495
         Left            =   2400
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl7 
         Caption         =   "&A"
         Height          =   495
         Left            =   2760
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl9 
         Caption         =   "BINARIO"
         Height          =   495
         Left            =   4320
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl10 
         Caption         =   "HEXADECIMAL"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl12 
         BackColor       =   &H8000000B&
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   3480
      Width           =   2775
      Begin VB.OptionButton opt2 
         Caption         =   "CATODO COMUN"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton opt1 
         Caption         =   "ANODO COMUN"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label2"
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   1920
      Width           =   375
   End
   Begin VB.Shape shpd 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   720
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpc 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1200
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape shpg 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   720
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpe 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   600
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape shpb 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1200
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape shpa 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   720
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape shpf 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   600
      Top             =   1080
      Width           =   135
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "&Display"
      Begin VB.Menu mnuDisplayAnodo 
         Caption         =   "Anodo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDisplayCatodo 
         Caption         =   "Catodo"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuGuion 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisplaySalir 
         Caption         =   "Salir"
         Shortcut        =   +{F4}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public N1, N2, N3, H, H1, H2, H3, H4, H5, H6 As Double


Private Sub Check1_Click()

shpg.FillColor = &HFF&
shpg.BorderColor = &HFF&

If opt1.Value = True Then

    If Check1.Value = Checked Then
    lbl19.Caption = 0
    shpg.FillColor = &HFF&

    Else: Check1.Value = uncheked
    lbl19.Caption = 1
    shpg.FillColor = &H0&
    shpg.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check1.Value = Checked Then
    lbl19.Caption = 1
    shpg.FillColor = &HFF&

    Else: Check1.Value = uncheked
    lbl19.Caption = 0
    shpg.FillColor = &H0&
    shpg.BorderColor = &H0&
    End If

    If Check1.Value = Checked Then
    H6 = H5 + 1
Else: Check1.Value = Unchecked
  
    H6 = H - 1
End If
    
    End If


End Sub

Private Sub Check2_Click()
shpf.FillColor = &HFF&
shpf.BorderColor = &HFF&

If opt1.Value = True Then

    If Check2.Value = Checked Then
    lbl18.Caption = 0
    shpf.FillColor = &HFF&

    Else: Check2.Value = uncheked
    lbl18.Caption = 1
    shpf.FillColor = &H0&
    shpf.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check2.Value = Checked Then
    lbl18.Caption = 1
    shpf.FillColor = &HFF&

    Else: Check2.Value = uncheked
    lbl18.Caption = 0
    shpf.FillColor = &H0&
    shpf.BorderColor = &H0&
    End If
    If Check2.Value = Checked Then
    shpf.BackColor = vbRed
    H5 = H4 + 1
Else: Check2.Value = Unchecked
    shpf.BackColor = &HC0FFFF
    H5 = H4 - 1
End If

    End If


End Sub

Private Sub Check3_Click()

shpe.FillColor = &HFF&
shpe.BorderColor = &HFF&

If opt1.Value = True Then

    If Check3.Value = Checked Then
    lbl17.Caption = 0
    shpe.FillColor = &HFF&

    Else: Check3.Value = uncheked
    lbl17.Caption = 1
    shpe.FillColor = &H0&
    shpe.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check3.Value = Checked Then
    lbl17.Caption = 1
    shpe.FillColor = &HFF&

    Else: Check3.Value = uncheked
    lbl17.Caption = 0
    shpe.FillColor = &H0&
    shpe.BorderColor = &H0&
    End If
    
   If Check3.Value = Checked Then

    
    H4 = H3 + 1
Else: Check3.Value = Unchecked
    
    H4 = H3 - 1
End If
    End If
  

End Sub

Private Sub Check4_Click()

shpd.FillColor = &HFF&
shpd.BorderColor = &HFF&

If opt1.Value = True Then

    If Check4.Value = Checked Then
    lbl16.Caption = 0
    shpd.FillColor = &HFF&

    Else: Check4.Value = uncheked
    lbl16.Caption = 1
    shpd.FillColor = &H0&
    shpd.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check4.Value = Checked Then
    lbl16.Caption = 1
    shpd.FillColor = &HFF&

    Else: Check4.Value = uncheked
    lbl16.Caption = 0
    shpd.FillColor = &H0&
    shpd.BorderColor = &H0&
    End If
    
     If Check4.Value = Checked Then
   
    H3 = H2 + 1
Else: Check4.Value = Unchecked
    
    H3 = H2 - 1
End If

    End If

End Sub

Private Sub Check5_Click()

shpc.FillColor = &HFF&
shpc.BorderColor = &HFF&

If opt1.Value = True Then

    If Check5.Value = Checked Then
    lbl15.Caption = 0
    shpc.FillColor = &HFF&

    Else: Check5.Value = uncheked
    lbl15.Caption = 1
    shpc.FillColor = &H0&
    shpc.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check5.Value = Checked Then
    lbl15.Caption = 1
    shpc.FillColor = &HFF&

    Else: Check5.Value = uncheked
    lbl15.Caption = 0
    shpc.FillColor = &H0&
    shpc.BorderColor = &H0&
    End If
    
   If Check5.Value = Checked Then
   
    H2 = H1 + 1
Else: Check5.Value = Unchecked
    
    H2 = H1 - 1
End If
    End If
   
End Sub

Private Sub Check6_Click()

shpb.FillColor = &HFF&
shpb.BorderColor = &HFF&

If opt1.Value = True Then

    If Check6.Value = Checked Then
    lbl14.Caption = 0
    shpb.FillColor = &HFF&

    Else: Check6.Value = uncheked
    lbl14.Caption = 1
    shpb.FillColor = &H0&
    shpb.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check6.Value = Checked Then
    lbl14.Caption = 1
    shpb.FillColor = &HFF&

    Else: Check6.Value = uncheked
    lbl14.Caption = 0
    shpb.FillColor = &H0&
    shpb.BorderColor = &H0&
    End If
    
    
    If Check6.Value = Checked Then
   
    H1 = H + 1
Else: Check6.Value = Unchecked
   
    H1 = H - 1
End If
    End If
   


End Sub

Private Sub Check7_Click()

shpa.FillColor = &HFF&
shpa.BorderColor = &HFF&
If opt1.Value = True Then

    If Check7.Value = Checked Then
    lbl12.Caption = 0
    shpa.FillColor = &HFF&

    Else: Check7.Value = uncheked
    lbl12.Caption = 1
    shpa.FillColor = &H0&
    shpa.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check7.Value = Checked Then
    lbl12.Caption = 1
    shpa.FillColor = &HFF&

    Else: Check7.Value = uncheked
    lbl12.Caption = 0
    shpa.FillColor = &H0&
    shpa.BorderColor = &H0&
    End If
    
    
   If Check7.Value = Checked Then
    H = H + 1
Else: Check7.Value = Unchecked
  
    H = H - 1
End If

    End If

End Sub

Private Sub cmd1_Click()
End

End Sub

Private Sub Label2_Click()

End Sub

Private Sub cmd20_Click()
Check7.Value = Unchecked
Check6.Value = Unchecked
Check5.Value = Unchecked
Check4.Value = Unchecked
Check3.Value = Unchecked
Check2.Value = Unchecked
Check1.Value = Unchecked
txthexa.Text = " "

End Sub

Private Sub cmdhex_Click()

If opt2.Value = True Then
txthexa.Text = Hex(H)
End If

If opt1.Value = True Then
txthexa.Text = Hex(H)
End If
End Sub

Private Sub cmdpuer_Click()
If opt2.Value = True Then
Out &H378, H
End If
If opt2.Value = True Then
Out &H378, H1
End If
If opt2.Value = True Then
Out &H378, H2
End If
If opt2.Value = True Then
Out &H378, H3
End If
If opt2.Value = True Then
Out &H378, H4
End If
If opt2.Value = True Then
Out &H378, H5
End If
If opt2.Value = True Then
Out &H378, H6
End If


If opt1.Value = True Then
Out &H378, H
End If

If opt1.Value = True Then
Out &H378, H1
End If

If opt1.Value = True Then
Out &H378, H2
End If

If opt1.Value = True Then
Out &H378, H3
End If

If opt1.Value = True Then
Out &H378, H4
End If

If opt1.Value = True Then
Out &H378, H5
End If

If opt1.Value = True Then
Out &H378, H6
End If


End Sub

Private Sub mnuDisplaySalir_Click()
End
End Sub

Private Sub opt1_Click()
shpa.FillColor = &HFF&
shpa.BorderColor = &HFF&
If opt1.Value = True Then

    If Check7.Value = Checked Then
    lbl12.Caption = 0
    shpa.FillColor = &HFF&

    Else: Check7.Value = uncheked
    lbl12.Caption = 1
    shpa.FillColor = &H0&
    shpa.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check7.Value = Checked Then
    lbl12.Caption = 1
    shpa.FillColor = &HFF&

    Else: Check7.Value = uncheked
    lbl12.Caption = 0
    shpa.FillColor = &H0&
    shpa.BorderColor = &H0&

    End If
    End If
shpb.FillColor = &HFF&
shpb.BorderColor = &HFF&

If opt1.Value = True Then

    If Check6.Value = Checked Then
    lbl14.Caption = 0
    shpb.FillColor = &HFF&

    Else: Check6.Value = uncheked
    lbl14.Caption = 1
    shpb.FillColor = &H0&
    shpb.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check6.Value = Checked Then
    lbl14.Caption = 1
    shpb.FillColor = &HFF&

    Else: Check6.Value = uncheked
    lbl14.Caption = 0
    shpb.FillColor = &H0&
    shpb.BorderColor = &H0&
    End If
    End If
    shpc.FillColor = &HFF&
shpc.BorderColor = &HFF&

If opt1.Value = True Then

    If Check5.Value = Checked Then
    lbl15.Caption = 0
    shpc.FillColor = &HFF&

    Else: Check5.Value = uncheked
    lbl15.Caption = 1
    shpc.FillColor = &H0&
    shpc.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check5.Value = Checked Then
    lbl15.Caption = 1
    shpc.FillColor = &HFF&

    Else: Check5.Value = uncheked
    lbl15.Caption = 0
    shpc.FillColor = &H0&
    shpc.BorderColor = &H0&
    End If
    End If
    shpd.FillColor = &HFF&
shpd.BorderColor = &HFF&

If opt1.Value = True Then

    If Check4.Value = Checked Then
    lbl16.Caption = 0
    shpd.FillColor = &HFF&

    Else: Check4.Value = uncheked
    lbl16.Caption = 1
    shpd.FillColor = &H0&
    shpd.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check4.Value = Checked Then
    lbl16.Caption = 1
    shpd.FillColor = &HFF&

    Else: Check4.Value = uncheked
    lbl16.Caption = 0
    shpd.FillColor = &H0&
    shpd.BorderColor = &H0&
    End If
    End If
    shpe.FillColor = &HFF&
shpe.BorderColor = &HFF&

If opt1.Value = True Then

    If Check3.Value = Checked Then
    lbl17.Caption = 0
    shpe.FillColor = &HFF&

    Else: Check3.Value = uncheked
    lbl17.Caption = 1
    shpe.FillColor = &H0&
    shpe.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check3.Value = Checked Then
    lbl17.Caption = 1
    shpe.FillColor = &HFF&

    Else: Check3.Value = uncheked
    lbl17.Caption = 0
    shpe.FillColor = &H0&
    shpe.BorderColor = &H0&
    End If
    End If
shpf.FillColor = &HFF&
shpf.BorderColor = &HFF&

If opt1.Value = True Then

    If Check2.Value = Checked Then
    lbl18.Caption = 0
    shpf.FillColor = &HFF&

    Else: Check2.Value = uncheked
    lbl18.Caption = 1
    shpf.FillColor = &H0&
    shpf.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check2.Value = Checked Then
    lbl18.Caption = 1
    shpf.FillColor = &HFF&

    Else: Check2.Value = uncheked
    lbl18.Caption = 0
    shpf.FillColor = &H0&
    shpf.BorderColor = &H0&
    End If
    End If
    shpg.FillColor = &HFF&
shpg.BorderColor = &HFF&

If opt1.Value = True Then

    If Check1.Value = Checked Then
    lbl19.Caption = 0
    shpg.FillColor = &HFF&

    Else: Check1.Value = uncheked
    lbl19.Caption = 1
    shpg.FillColor = &H0&
    shpg.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check1.Value = Checked Then
    lbl19.Caption = 1
    shpg.FillColor = &HFF&

    Else: Check1.Value = uncheked
    lbl19.Caption = 0
    shpg.FillColor = &H0&
    shpg.BorderColor = &H0&
    
    End If
    If opt1.Value = True Then
    H = 0
Else
    H = 1
End If
 
    End If
End Sub

Private Sub opt2_Click()
shpa.FillColor = &HFF&
shpa.BorderColor = &HFF&
If opt1.Value = True Then

    If Check7.Value = Checked Then
    lbl12.Caption = 0
    shpa.FillColor = &HFF&

    Else: Check7.Value = uncheked
    lbl12.Caption = 1
    shpa.FillColor = &H0&
    shpa.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check7.Value = Checked Then
    lbl12.Caption = 1
    shpa.FillColor = &HFF&

    Else: Check7.Value = uncheked
    lbl12.Caption = 0
    shpa.FillColor = &H0&
    shpa.BorderColor = &H0&

    End If
    End If
shpb.FillColor = &HFF&
shpb.BorderColor = &HFF&

If opt1.Value = True Then

    If Check6.Value = Checked Then
    lbl14.Caption = 0
    shpb.FillColor = &HFF&

    Else: Check6.Value = uncheked
    lbl14.Caption = 1
    shpb.FillColor = &H0&
    shpb.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check6.Value = Checked Then
    lbl14.Caption = 1
    shpb.FillColor = &HFF&

    Else: Check6.Value = uncheked
    lbl14.Caption = 0
    shpb.FillColor = &H0&
    shpb.BorderColor = &H0&
    End If
    End If
    shpc.FillColor = &HFF&
shpc.BorderColor = &HFF&

If opt1.Value = True Then

    If Check5.Value = Checked Then
    lbl15.Caption = 0
    shpc.FillColor = &HFF&

    Else: Check5.Value = uncheked
    lbl15.Caption = 1
    shpc.FillColor = &H0&
    shpc.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check5.Value = Checked Then
    lbl15.Caption = 1
    shpc.FillColor = &HFF&

    Else: Check5.Value = uncheked
    lbl15.Caption = 0
    shpc.FillColor = &H0&
    shpc.BorderColor = &H0&
    End If
    End If
    shpd.FillColor = &HFF&
shpd.BorderColor = &HFF&

If opt1.Value = True Then

    If Check4.Value = Checked Then
    lbl16.Caption = 0
    shpd.FillColor = &HFF&

    Else: Check4.Value = uncheked
    lbl16.Caption = 1
    shpd.FillColor = &H0&
    shpd.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check4.Value = Checked Then
    lbl16.Caption = 1
    shpd.FillColor = &HFF&

    Else: Check4.Value = uncheked
    lbl16.Caption = 0
    shpd.FillColor = &H0&
    shpd.BorderColor = &H0&
    End If
    End If
    shpe.FillColor = &HFF&
shpe.BorderColor = &HFF&

If opt1.Value = True Then

    If Check3.Value = Checked Then
    lbl17.Caption = 0
    shpe.FillColor = &HFF&

    Else: Check3.Value = uncheked
    lbl17.Caption = 1
    shpe.FillColor = &H0&
    shpe.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check3.Value = Checked Then
    lbl17.Caption = 1
    shpe.FillColor = &HFF&

    Else: Check3.Value = uncheked
    lbl17.Caption = 0
    shpe.FillColor = &H0&
    shpe.BorderColor = &H0&
    End If
    End If
shpf.FillColor = &HFF&
shpf.BorderColor = &HFF&

If opt1.Value = True Then

    If Check2.Value = Checked Then
    lbl18.Caption = 0
    shpf.FillColor = &HFF&

    Else: Check2.Value = uncheked
    lbl18.Caption = 1
    shpf.FillColor = &H0&
    shpf.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check2.Value = Checked Then
    lbl18.Caption = 1
    shpf.FillColor = &HFF&

    Else: Check2.Value = uncheked
    lbl18.Caption = 0
    shpf.FillColor = &H0&
    shpf.BorderColor = &H0&
    End If
    End If
    shpg.FillColor = &HFF&
shpg.BorderColor = &HFF&

If opt1.Value = True Then

    If Check1.Value = Checked Then
    lbl19.Caption = 0
    shpg.FillColor = &HFF&

    Else: Check1.Value = uncheked
    lbl19.Caption = 1
    shpg.FillColor = &H0&
    shpg.BorderColor = &H0&
    
    End If
  
ElseIf opt2.Value = True Then

    If Check1.Value = Checked Then
    lbl19.Caption = 1
    shpg.FillColor = &HFF&

    Else: Check1.Value = uncheked
    lbl19.Caption = 0
    shpg.FillColor = &H0&
    shpg.BorderColor = &H0&
    End If
     If opt2.Value = True Then
    H = 1
Else
    H = 0
End If
 
    End If
End Sub

