VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Uso del ScrollBar"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   765
      Left            =   3840
      TabIndex        =   5
      Top             =   2355
      Width           =   1080
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1575
      Left            =   2925
      Max             =   100
      TabIndex        =   4
      Top             =   1005
      Width           =   285
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      Left            =   420
      Max             =   100
      TabIndex        =   3
      Top             =   2625
      Width           =   2505
   End
   Begin VB.Label lblVScrollBar 
      Caption         =   "Para el VScrollBar: 0"
      Height          =   255
      Left            =   465
      TabIndex        =   2
      Top             =   1725
      Width           =   2115
   End
   Begin VB.Label lblHScrollBar 
      Caption         =   "Para el HScrollBar: 0"
      Height          =   255
      Left            =   465
      TabIndex        =   1
      Top             =   1275
      Width           =   2115
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      Caption         =   $"frmScrollBar2.frx":0000
      Height          =   660
      Left            =   825
      TabIndex        =   0
      Top             =   180
      Width           =   3810
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
'Cierra la ventana
Unload Me
End Sub

Private Sub VScroll1_Change()
'Cuando el VScrollBar cambia de posición
'se muestra su nuevo valor en su Label.
lblVScrollBar.Caption = "Para el VScrollBar: " & VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
'Cuando el VScrollBar se está deslizando
'se muestra su nuevo valor en su Label.
lblVScrollBar.Caption = "Para el VScrollBar: " & VScroll1.Value
End Sub

Private Sub HScroll1_Change()
'Cuando el HScrollBar cambia de posición
'se muestra su nuevo valor en su Label.
lblHScrollBar.Caption = "Para el HScrollBar: " & HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
'Cuando el HScrollBar se está deslizando
'se muestra su nuevo valor en su Label.
lblHScrollBar.Caption = "Para el HScrollBar: " & HScroll1.Value
End Sub

