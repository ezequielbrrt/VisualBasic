VERSION 2.00
Begin Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
   Height          =   6225
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Top             =   1140
   Width           =   7485
   Begin CommandButton Command1 
      Caption         =   "Convertir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   2835
   End
   Begin Label Label1 
      Height          =   3915
      Left            =   240
      TabIndex        =   2
      Top             =   1380
      Width           =   6435
   End
End
Option Explicit

Sub Command1_Click ()
label1 = Numlet(CCur(text1))
End Sub

