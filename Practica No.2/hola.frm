VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MouseIcon       =   "hola.frx":0000
   ScaleHeight     =   4620
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd3 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "ver &Adios"
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ver &Hola"
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ADIOS"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl 
      BackColor       =   &H000000FF&
      Caption         =   "HOLA"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub cmd_Click()
lbl.Visible = True
lbl.ForeColor = vbBlue
cmd.Enabled = False


End Sub

Private Sub cmd2_Click()
lbl2.Visible = True
lbl.Visible = False
cmd2.Enabled = False
End Sub

Private Sub cmd3_Click()
End

End Sub

Private Sub Form_Load()

End Sub
