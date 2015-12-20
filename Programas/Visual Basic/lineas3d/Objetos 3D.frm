VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saul Olguin Aguirre 05 Abril 2005"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   1575
      Left            =   5880
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   1995
      Left            =   5730
      TabIndex        =   1
      Top             =   2340
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Este es un frame real. Mientras que el resto son otros objetos"
      Height          =   1515
      Left            =   360
      TabIndex        =   0
      Top             =   285
      Width           =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   900
      X2              =   5175
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   1215
      Index           =   0
      Left            =   465
      Top             =   2130
      Width           =   4605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   1215
      Index           =   3
      Left            =   495
      Shape           =   2  'Oval
      Top             =   3795
      Width           =   4605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Height          =   1215
      Index           =   1
      Left            =   480
      Top             =   2145
      Width           =   4605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Height          =   1215
      Index           =   2
      Left            =   510
      Shape           =   2  'Oval
      Top             =   3810
      Width           =   4605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   1
      X1              =   915
      X2              =   5190
      Y1              =   5595
      Y2              =   5595
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

Private Sub Command2_Click()
End
End Sub
