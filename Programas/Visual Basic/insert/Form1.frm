VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   1800
   ClientTop       =   1770
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5235
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
gInsert = True
MuestraInsert
End Sub

Private Sub MuestraInsert()
'para mostrar si estamos o no insertando
'si se va a usar en el programa es mejor incluirla en la funcion ControlaInsert
If gInsert Then
    label1 = "Insertar"
Else
    label1 = "Sobreescribir"
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
ControlaInsert text1, KeyCode, Shift, gInsert
MuestraInsert
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
ControlaInsert text1, KeyAscii, -1, gInsert
End Sub

