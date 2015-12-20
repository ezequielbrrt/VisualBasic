VERSION 5.00
Begin VB.Form FrmFondo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar el fondo de Windows 95/NT"
   ClientHeight    =   5055
   ClientLeft      =   1410
   ClientTop       =   1920
   ClientWidth     =   5175
   Icon            =   "FondoAPI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5055
   ScaleWidth      =   5175
   Begin VB.CommandButton WinDir 
      Caption         =   "Directorio Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   2775
      Begin VB.OptionButton Mosaico 
         Caption         =   "Mosaico"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Mosaico 
         Caption         =   "Centrado"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ajustar imagen al cuadro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   0
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Vista Previa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   2520
      Pattern         =   "*.bmp"
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label LblWP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Tag             =   "Fondo actual: "
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   480
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image ImgPantalla 
      Height          =   1395
      Left            =   195
      Picture         =   "FondoAPI.frx":0442
      Stretch         =   -1  'True
      Top             =   570
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esc - Salir del programa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para cambiar el tapiz, hacer doble click en el nombre del fichero"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "FrmFondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Ajustado As Boolean
Dim AltoImg As Integer, AnchoImg As Integer, CadTemp$

Private Sub Cambiar(Optional Fichero)
If (IsMissing(Fichero)) Then
Fichero = Dir1.Path + "\" + File1.List(File1.ListIndex) '.filename
End If
N = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, Fichero, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
If N = 0 Then
   MsgBox "Se ha producido un error con el fichero " & Fichero
Else
   Me.Tag = Fichero
   Mosaico(0).Enabled = True
   Mosaico(1).Enabled = True
   LblWP = LblWP.Tag & Fichero
End If
End Sub

Private Sub Check1_Click()
With Image1
If Check1.Value = vbChecked Then
    .Height = AltoImg
    .Width = AnchoImg
    .Stretch = True
Else
    Image1.Stretch = False
End If
End With
End Sub

Private Sub Check2_Click()
Image1.Visible = IIf(Check2.Value = vbChecked, True, False)
Check1.Enabled = Image1.Visible
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Label1.Visible = File1.ListCount > 0
If Label1.Visible Then
File1.ListIndex = 0
End If
End Sub

Private Sub Drive1_Change()
On Error GoTo NoDrive
Dir1.Path = Drive1.Drive
Exit Sub
NoDrive:
Drive1.Drive = Left(Dir1.Path, 2)
End Sub

Private Sub File1_Click()
Dim Fil As String
Fil = Dir1.Path + "\" + File1.List(File1.ListIndex)
Image1.Picture = LoadPicture(Fil)
Label1.Visible = True
Mosaico(0).Enabled = (Fil = Me.Tag)
Mosaico(1).Enabled = (Fil = Me.Tag)
End Sub

Private Sub File1_DblClick()
Cambiar
End Sub

Private Sub Form_Activate()
Dim D As String, B As Long, WP As String
WP = String$(255, " ")
B = GetProfileString("desktop", "Wallpaper", "", WP, Len(WP))
WP = Left$(WP, B)
CDir WP
D = String$(255, " ")
B = GetProfileString("desktop", "TileWallpaper", "", D, Len(D))
Mosaico(Val(Left$(D, B))).Value = True
File1_DblClick
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub

Private Sub Form_Load()
Pantalla.Alto = 600
Pantalla.Ancho = 800
Ajustar Me
AltoImg = Image1.Height 'Para conservar el tamaño en tiempo de diseño
AnchoImg = Image1.Width
'Para poder observar el fondo se pone me.left=0
Me.Left = 0 '(Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Check1.Value = vbChecked
Check2.Value = vbChecked
CadTemp = String$(255, " ")
End Sub

Private Sub Mosaico_Click(Index As Integer)
Dim B As Long, CadReg As String, F As String
CadReg = "REGEDIT4" & vbCrLf & vbCrLf
CadReg = CadReg & "[HKEY_CURRENT_USER\Control Panel\desktop]" & vbCrLf
CadReg = CadReg & Chr$(34) & "TileWallpaper" & Chr$(34)
CadReg = CadReg & "=" & Chr$(34) & Trim(Str$(Index)) & Chr$(34)
F = App.Path & "\" & App.EXEName & ".reg"
B = FreeFile
Open F For Output As #B
Print #B, CadReg
Close #B
B = Shell("regedit /s " & F, vbNormalFocus)
CadTemp = String$(255, " ")
B = GetProfileString("desktop", "Wallpaper", "", CadTemp, Len(CadTemp))
CadTemp = Left$(CadTemp, B)
Cambiar CadTemp
Kill F
End Sub

Private Sub WinDir_Click()
Dim WD As String, B As Long
WD = String$(255, " ")
B = GetWindowsDirectory(WD, 255)
WD = Left$(WD, B) & "\."
CDir WD
End Sub
