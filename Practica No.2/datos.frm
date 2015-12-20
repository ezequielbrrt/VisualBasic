VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   4845
   ClientTop       =   4350
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdTodo 
      Caption         =   "VER TODO"
      Height          =   855
      Left            =   3600
      TabIndex        =   11
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "BORRAR"
      Height          =   855
      Left            =   1080
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdVer3 
      Caption         =   "VER"
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdVer2 
      Caption         =   "VER"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdVer1 
      Caption         =   "VER"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtEscuela 
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtTelefono 
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtNombre 
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblEscuela 
      AutoSize        =   -1  'True
      Caption         =   "ESCUELA"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lbltelefono 
      AutoSize        =   -1  'True
      Caption         =   "TELEFONO"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "DATOS PERSONALES"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBorrar_Click()
txtNombre.Text = ""
txtNombre.BackColor = vbYellow
txtNombre.ForeColor = vbBlack

txtTelefono.Text = ""
txtTelefono.BackColor = vbYellow
txtTelefono.ForeColor = vbBlack

txtEscuela.Text = ""
txtEscuela.BackColor = vbYellow
txtEscuela.ForeColor = vbBlack


End Sub

Private Sub cmdTodo_Click()
txtNombre.Visible = True
txtTelefono.Visible = True
txtEscuela.Visible = True

txtNombre.Text = "CESAR MORZA"
txtNombre.BackColor = vbYellow
txtNombre.ForeColor = vbBlack

txtTelefono.Text = "5556938620"
txtTelefono.BackColor = vbYellow
txtTelefono.ForeColor = vbBlack

txtEscuela.Text = "CECYT 1"
txtEscuela.BackColor = vbYellow
txtEscuela.ForeColor = vbBlack



End Sub

Private Sub cmdVer1_Click()
txtNombre.Text = "CESAR MORZA"
txtNombre.BackColor = vbYellow
txtNombre.ForeColor = vbBlack


End Sub

Private Sub cmdVer2_Click()
txtTelefono.Text = "5556938620"
txtTelefono.BackColor = vbYellow
txtTelefono.ForeColor = vbBlack



End Sub

Private Sub cmdVer3_Click()
txtEscuela.Text = "CECYT 1"
txtEscuela.BackColor = vbYellow
txtEscuela.ForeColor = vbBlack



End Sub
