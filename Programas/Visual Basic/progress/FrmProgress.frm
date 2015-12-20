VERSION 5.00
Object = "{AF8CD3F4-666F-11D1-940D-000021A73813}#5.0#0"; "OSPROGRESS.OCX"
Begin VB.Form FrmProgress 
   Caption         =   "Formulario de prueba del control osProgress"
   ClientHeight    =   4620
   ClientLeft      =   1950
   ClientTop       =   2085
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5100
   Begin Progress.osProgress osProgress1 
      Height          =   1005
      Left            =   1058
      TabIndex        =   3
      Top             =   998
      Width           =   2685
      _ExtentX        =   6932
      _ExtentY        =   1111
      BorderWidth     =   98
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   115
      DelayTime       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visualizar fuentes del sistema"
      Height          =   615
      Left            =   1102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   795
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   735
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
   End
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim I As Integer
Command1.Enabled = False
With osProgress1
   .Min = 0
   .Value = 0
   .Max = Screen.FontCount - 1
   For I = 0 To .Max
      Label1 = I & " de " & Screen.FontCount - 1
      Label2 = Screen.Fonts(I)
      .Value = I
      DoEvents
   Next
End With
Command1.Enabled = True
End Sub


Private Sub Form_Load()
Label1 = "0 de " & Screen.FontCount - 1
End Sub

Private Sub osProgress1_Finished()
MsgBox "Terminado"
End Sub
