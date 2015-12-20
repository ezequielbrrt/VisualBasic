VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saul Olguin Aguirre 05 Abril 2005"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir de esta aplicacion"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then Me.WindowState = vbMaximized
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
