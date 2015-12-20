VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   4920
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3135
      Left            =   360
      Max             =   255
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   3120
      Y1              =   1560
      Y2              =   2640
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   1560
      Y2              =   2280
   End
   Begin VB.Shape Shape3 
      Height          =   1095
      Left            =   2520
      Shape           =   2  'Oval
      Top             =   480
      Width           =   855
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   2520
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VScroll1_Change()

End Sub
