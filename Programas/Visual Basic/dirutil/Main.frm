VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saúl Olguín Aguirre  05 Abril 2005"
   ClientHeight    =   5520
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7335
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
   ScaleHeight     =   5520
   ScaleWidth      =   7335
   Begin VB.PictureBox outDirInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4545
      ScaleWidth      =   7305
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label PTamaño 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Número"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5880
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label PFicheros 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   " Serie"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2940
      TabIndex        =   5
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tamaño ficheros :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4260
      TabIndex        =   4
      Top             =   4800
      Width           =   1545
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nº ficheros :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1740
      TabIndex        =   3
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblBusqueda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuscar_Click()
Dim x As Integer, NumFich As Long, Tamaño As Long

   NumFich = 0: Tamaño = 0
   Me.MousePointer = 11
   x = DirUtil(Left(CurDir$, 2), 1, NumFich, Tamaño)

   ' expandir todas las ramas
   For x = 0 To outDirInfo.ListCount - 1
      outDirInfo.Expand(x) = True
   Next
   lblBusqueda = "Finalizado"
   PFicheros = Format$(NumFich, "###,###,##0")
   PTamaño = Format$(Tamaño, "###,###,###,##0") + " Kb"
   Me.MousePointer = 0
End Sub

