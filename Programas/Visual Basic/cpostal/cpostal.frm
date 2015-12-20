VERSION 5.00
Begin VB.Form CPfrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saul Olguin Aguirre"
   ClientHeight    =   3000
   ClientLeft      =   1680
   ClientTop       =   2460
   ClientWidth     =   5430
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
   ScaleHeight     =   3000
   ScaleWidth      =   5430
   Begin VB.CommandButton CmdPobDir 
      Appearance      =   0  'Flat
      Caption         =   "Cambiar Pob/Dir"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2580
      Width           =   5175
   End
   Begin VB.ListBox ListCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   1395
      Left            =   4140
      TabIndex        =   6
      Tag             =   "ol"
      Top             =   1020
      Width           =   1155
   End
   Begin VB.ListBox ListPoblacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   1395
      Left            =   120
      TabIndex        =   5
      Tag             =   "ol"
      Top             =   1020
      Width           =   3855
   End
   Begin VB.TextBox TxtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4140
      TabIndex        =   4
      Tag             =   "ol"
      Text            =   "Código"
      Top             =   600
      Width           =   1155
   End
   Begin VB.TextBox TxtPoblacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Tag             =   "ol"
      Text            =   "Población"
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton CmdCerrar 
      Appearance      =   0  'Flat
      Caption         =   "Cerrar"
      Height          =   285
      Left            =   4140
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Tag             =   "ol"
      Text            =   "Combo1"
      Top             =   120
      Width           =   2715
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Provincia"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "CPfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
' gsCPostal Form para los códigos postales
' (c)Guillermo Som Cerezo, 1993-96
'--------------------------------------------------
Option Explicit
Option Compare Text

Dim iniH As Integer
Dim iniW As Integer
Dim lngTmp As Long
Dim Provincias() As tipoPobDAT
Dim numProv As Integer
Dim EsFormLoad As Integer
Dim CPostal As Variant
Dim miniCP As String
Dim Poblacion As String
Dim tCP As tipoPobDAT
Dim nomFicCP As String
Dim nomDirCP As String
Dim numFic As Integer
Dim numRegCP As Integer
Dim lenPOB As Integer
Dim PobODir As Integer
Dim Pob_Dir As String
Dim ClickEnList As Integer

Private Sub AbrirProvincia()
Dim i           As Integer
Dim Combo1Index As Integer
Dim sFile       As String
Dim lTmp        As Long

'Cambiar el cursor del ratón...
MousePointer = 11
DoEvents
'Abrir el fichero de datos de la provincia activa...
sFile = nomDirCP & Provincias(numProv).miniCP & Pob_Dir
ClickEnList = True
If Len(Dir$(sFile)) Then
    numFic = FreeFile
    Open sFile For Random As numFic Len = Len(tCP)
    numRegCP = LOF(numFic) \ Len(tCP)
    ListPoblacion.Clear
    ListCodigo.Clear
    TxtCodigo = ""
    TxtPoblacion = ""
    For i = 1 To numRegCP
        Get numFic, i, tCP
        lenPOB = Asc(tCP.PobLen)
        Poblacion = Left$(tCP.Poblacion, lenPOB)
        CPostal = tCP.CPostal
        ListPoblacion.AddItem Poblacion
        ListCodigo.AddItem CPostal
        ListPoblacion.ItemData(ListPoblacion.NewIndex) = i ' ListCodigo.NewIndex
        ListCodigo.ItemData(ListCodigo.NewIndex) = i ' ListPoblacion.NewIndex
    Next i
    Close numFic
End If
ClickEnList = False
If ListCodigo.ListCount > 0 Then
    ListCodigo.ListIndex = 0
End If
Caption = "Códigos Postales -" & Provincias(numProv).Prov
Provincia = Provincias(numProv).Prov
If PobODir Then
    CmdPobDir.Caption = "Mostrar las Direcciones de " & Provincia
Else
    CmdPobDir.Caption = "Mostrar las Poblaciones de " & Provincia
End If
MousePointer = 0
End Sub

Private Function BuscarEnList(oTextB As TextBox, oListB As ListBox) As Integer
'Actualizar el ListBox, según lo escrito en el TextBox
'Ver ficha 103 de Notas Guille
Dim i As Integer, j As Integer
Dim iHallado As Integer
Dim sTmp As String

If ClickEnList Then
    Exit Function
End If
If oListB.ListCount Then
    ClickEnList = True
    sTmp = oTextB.Text
    j = Len(sTmp)
    iHallado = 0
    For i = 0 To (oListB.ListCount - 1)
        If StrComp(sTmp, Left$(oListB.List(i), j)) = 0 Then
            iHallado = i + 1
            Exit For
        End If
    Next i
    If iHallado Then
        oListB.TopIndex = iHallado - 1
        oListB.ListIndex = iHallado - 1
    End If
    BuscarEnList = iHallado
    ClickEnList = False
End If
End Function

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdPobDir_Click()
'------------------------------------------------------
'Poder seleccionar ente Poblaciones o Direcciones
'                                           (25/May/94)
'------------------------------------------------------
PobODir = Not PobODir
If PobODir Then
    Pob_Dir = "POB.DAT"
Else
    Pob_Dir = "DIR.DAT"
End If
AbrirProvincia
End Sub

Private Sub Combo1_Change()
If Not EsFormLoad Then
    numProv = Combo1.ListIndex + 1
End If
End Sub

Private Sub Combo1_Click()
'Seleccionar la provincia
If Not EsFormLoad Then
    numProv = Combo1.ListIndex + 1
    AbrirProvincia
End If
End Sub

Private Function ConectarLists(ByVal queItem As Integer, oListB As ListBox) As Integer
Dim i As Integer
Dim iHallado As Integer

iHallado = 0
For i = 0 To oListB.ListCount - 1
    If oListB.ItemData(i) = queItem Then
        iHallado = i
        Exit For
    End If
Next
ConectarLists = iHallado
End Function

Private Sub Form_Load()
'Abrir el fichero de provincias
Dim i As Integer
Dim ProvinciaDat As String

'Sólo una copia cada vez
If App.PrevInstance Then
    End
End If
iniH = Height
iniW = Width
'Posicionar la ventana
Centrar Me

If Provincia = "" Then
    Provincia = "MALAGA"
End If
PobODir = True
Pob_Dir = "POB.DAT"
'
'Buscar el fichero PROV.DAT en este orden:
'primero: en el directorio C:\CP
'segundo: en el de la aplicación
'tercero: en el directorio actual
'
On Error Resume Next

nomFicCP = "C:\CP\PROV.DAT"
nomDirCP = "C:\CP\"
i = 0
Do
    If Len(Dir$(nomFicCP)) = 0 Then
        If i = 0 Then
            i = i + 1
            nomFicCP = App.Path & "\PROV.DAT"
            nomDirCP = App.Path & "\"
        ElseIf i = 1 Then
            i = 2
            nomFicCP = "PROV.DAT"
            nomDirCP = ""
            Else
                Beep
                MsgBox "No he hallado el fichero de PROVINCIAS (PROV.DAT)," & Chr$(13) & Chr$(13) & "Este archivo debe estar en una de estas localizaciones:" & Chr$(13) & "   El directorio C:\CP," & Chr$(13) & "   El directorio del programa " & App.Path & "," & Chr$(13) & "   El directorio actual." & Chr$(13) & Chr$(13) & "Programa Terminado.", MB_ICONSTOP, "Código Postal"
                End
            End If

    Else
        Exit Do
    End If
Loop
numProv = 0
numFic = FreeFile
Open nomFicCP For Input As numFic
Do While Not EOF(numFic)
    Input #numFic, ProvinciaDat
    Input #numFic, miniCP
    ProvinciaDat = Trim$(ProvinciaDat)
    numProv = numProv + 1
    'Reservar memoria...
    ReDim Preserve Provincias(numProv) As tipoProvDAT
    Provincias(numProv).Prov = ProvinciaDat
    Provincias(numProv).miniCP = Left$(Trim$(miniCP), 2)
    Combo1.AddItem ProvinciaDat
    Combo1.ItemData(Combo1.NewIndex) = Val(miniCP)
Loop
Close numFic
EsFormLoad = True
Combo1.ListIndex = 0
If Provincia <> "" Then
    For i = 1 To numProv
        If Provincia = Provincias(i).Prov Then
            Combo1.ListIndex = i - 1
            Exit For
        End If
    Next
End If
numProv = Combo1.ListIndex + 1
AbrirProvincia
HayCP = True
EsFormLoad = False
End Sub

Private Sub Form_Paint()
WOutLines Me
OutLines Me
End Sub

Private Sub Form_Resize()
 If WindowState <> 1 Then
    Height = iniH
    Width = iniW
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SaveSetting "CPostal.ini", "Posiciones", "Top", Top
'SaveSetting "CPostal.ini", "Posiciones", "Left", Left
'SaveSetting "CPostal.ini", "General", "Provincia", Provincia
Set CPfrm = Nothing
End Sub

Private Sub ListCodigo_Click()
'Conectar los ListBoxs
ListPoblacion.ListIndex = ConectarLists(ListCodigo.ItemData(ListCodigo.ListIndex), ListPoblacion)
ListPoblacion.TopIndex = ListPoblacion.ListIndex
ListCodigo.TopIndex = ListCodigo.ListIndex
If ClickEnList Then Exit Sub
ClickEnList = True
TxtPoblacion.Text = ListPoblacion.Text
TxtCodigo.Text = ListCodigo.Text
ClickEnList = False
End Sub

Private Sub ListPoblacion_Click()
'Conectar los ListBoxs
ListCodigo.ListIndex = ConectarLists(ListPoblacion.ItemData(ListPoblacion.ListIndex), ListCodigo)
ListPoblacion.TopIndex = ListPoblacion.ListIndex
ListCodigo.TopIndex = ListCodigo.ListIndex
If ClickEnList Then Exit Sub
ClickEnList = True
TxtPoblacion.Text = ListPoblacion.Text
TxtCodigo.Text = ListCodigo.Text
ClickEnList = False
End Sub

Private Sub TxtCodigo_Change()
lngTmp = BuscarEnList(TxtCodigo, ListCodigo)
End Sub

Private Sub TxtCodigo_GotFocus()
TxtCodigo.SelStart = 0
TxtCodigo.SelLength = Len(TxtCodigo.Text)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BuscarEnList(TxtCodigo, ListCodigo) Then
        TxtCodigo = ListCodigo.Text
        TxtPoblacion = ListPoblacion.Text
    End If
    KeyAscii = 0
    TxtCodigo_GotFocus
End If
End Sub

Private Sub TxtPoblacion_Change()
lngTmp = BuscarEnList(TxtPoblacion, ListPoblacion)
End Sub

Private Sub TxtPoblacion_GotFocus()
TxtPoblacion.SelStart = 0
TxtPoblacion.SelLength = Len(TxtPoblacion.Text)
End Sub

Private Sub TxtPoblacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If BuscarEnList(TxtPoblacion, ListPoblacion) Then
        TxtCodigo = ListCodigo.Text
        TxtPoblacion = ListPoblacion.Text
    End If
    KeyAscii = 0
    TxtPoblacion_GotFocus
End If
End Sub

