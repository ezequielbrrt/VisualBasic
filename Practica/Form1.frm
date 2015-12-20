VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   4875
   ClientTop       =   2490
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8670
   Begin VB.CommandButton cmdbo 
      Caption         =   "&RESET"
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6600
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5400
      Top             =   4080
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2&Segundos"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1 &Segundo"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   7440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblseg2 
      Caption         =   "2 &Segundos"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblseg 
      Alignment       =   2  'Center
      Caption         =   "1& Segundo"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Bienvenido a este programa"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub cmd1_Click()

Timer2.Enabled = True





End Sub

Private Sub cmd2_Click()

Timer3.Enabled = True


End Sub

Private Sub cmdbo_Click()
lblseg.Visible = False
lblseg2.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False









End Sub

Private Sub cmdSalir_Click()
MsgBox "hasta pronto", vbOKOnly + 48, "salida"
End
End Sub

Private Sub lblseg2_Click()
lblseg.Visible = False


End Sub

Private Sub Timer1_Timer()
lbl1.Visible = True
Timer1.Enabled = True


End Sub

Private Sub Timer2_Timer()
lblseg.Visible = True




End Sub

Private Sub Timer3_Timer()
lblseg2.Visible = True

End Sub
