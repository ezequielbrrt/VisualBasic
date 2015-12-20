VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Saúl Olguín Aguirre  05 Abril 2005"
   ClientHeight    =   2265
   ClientLeft      =   1170
   ClientTop       =   1500
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2265
   ScaleWidth      =   5115
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Número de Serie"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Introduzca la letra de la Unidad a explorar:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVolumeInformation Lib _
"kernel32" Alias "GetVolumeInformationA" _
(ByVal lpRootPathName As String, ByVal _
lpVolumeNameBuffer As String, ByVal nVolumeNameSize _
As Long, lpVolumeSerialNumber As Long, _
lpMaximumComponentLength As Long, lpFileSystemFlags _
As Long, ByVal lpFileSystemNameBuffer As String, _
ByVal nFileSystemNameSize As Long) As Long


Private Sub Command1_Click()
Dim A As String
Dim VolumeNameBuffer As String
Dim tamaVolume As Long
Dim SerieNum As Long
Dim Maximo As Long
Dim Band As Long
Dim SystemaBuf As Long
Dim TamNameSize As Long
Dim V As Long
A = Text1 & ":\"
V = GetVolumeInformation(A, VolumeNameBuffer, _
    tamaVolume, SerieNum, Maximo, Band, SystemaBuf, _
    TamNameSize)
If SerieNum <> 0 Then
    Label1 = "Número de serie: " & Left(Hex(SerieNum), 4) _
    & "-" & Right(Hex(SerieNum), 4)
Else
    Label1 = "No se para que sirve esta aplicación"
End If
End Sub

Private Sub Form_Load()
Move (Screen.Height - Height) \ 2, (Screen.Width - Width) \ 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
