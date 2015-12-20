VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de Visual Basic"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   3195
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H80000008&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   1260
      TabIndex        =   0
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(c) 2005"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   2115
   End
   Begin VB.Image imgUser 
      Height          =   630
      Left            =   1620
      MouseIcon       =   "frmAbout.frx":095E
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":0AB0
      Top             =   1740
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "e-mail:solguin@nextsystem.com.mx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   840
      MouseIcon       =   "frmAbout.frx":162A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Saúl Olguin Aguirre"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Desarrollo:"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   300
      Width           =   1035
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub imgUser_Click()
    Static cantclicks As Integer
    
    If cantclicks > 2 Then cantclicks = 0
    imgUser.Picture = LoadResPicture(100 + cantclicks, vbResBitmap)
    
    cantclicks = cantclicks + 1
End Sub

'Esto es para llamar al programa de e-mail
Private Sub Label1_Click(Index As Integer)
    Dim result As Long
    
    If Index = 2 Then
        Screen.MousePointer = vbHourglass
        result = ShellExecute(Me.hWnd, vbNullString, "mailto:solguin@nextsystem.com.mx", vbNullString, "c:\", 1)
        Screen.MousePointer = vbNormal
    End If
    
End Sub
