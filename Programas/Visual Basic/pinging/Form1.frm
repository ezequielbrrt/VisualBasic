VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pinging"
   ClientHeight    =   7695
   ClientLeft      =   2055
   ClientTop       =   1545
   ClientWidth     =   9690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox C 
      Height          =   495
      Left            =   9120
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox B 
      Height          =   495
      Left            =   9120
      Picture         =   "Form1.frx":0884
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox A 
      Height          =   495
      Left            =   9120
      Picture         =   "Form1.frx":0CC6
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Text            =   "2"
      Top             =   360
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9240
      Top             =   720
   End
   Begin VB.TextBox NomIp 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton Borrar 
      Caption         =   "&Borrar Dirección"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.ListBox Result 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   9495
   End
   Begin VB.TextBox DirIp 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Cargar 
      Caption         =   "&Cargar Nueva Dirección"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label11 
      Caption         =   "Dirección IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Archivo As String
Dim Matriz() As RegPing


Private Sub Borrar_Click()
Dim i As Integer
Timer1.Enabled = Not Timer1.Enabled
If MsgBox("Esta seguro de Eliminar este Direccion IP", vbYesNo, "Confirmar") = vbYes Then
    Open Archivo For Output As #1
    For i = 1 To UBound(Matriz)
        If Trim(Matriz(i).Direccion) <> Trim(DirIp.Text) Then
            Print #1, Matriz(i).Direccion & "," & Matriz(i).Nombre
        End If
    Next
    Close #1
End If
CargarMatriz
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Cargar_Click()
Open Archivo For Append As #1
Print #1, DirIp.Text & "," & NomIp.Text
Close #1
CargarMatriz
End Sub


Private Sub Command1_Click()
End
End Sub

Private Sub NomIp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cargar.SetFocus
End If
End Sub


Private Sub DirIp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    NomIp.SetFocus
End If
End Sub

Private Sub Form_Load()
DirIp.MaxLength = C_LongDir
NomIp.MaxLength = C_LongNom
Archivo = App.Path & "\PING.TXT"
CargarMatriz
End Sub

Sub CargarMatriz()
Dim linea As String
Dim LongMat As Integer

ReDim Matriz(1)
Matriz(1).Direccion = "127.127.127.127"
Matriz(1).Nombre = "Mi Pc"
If Dir(Archivo) = "PING.TXT" Then
    Open Archivo For Input As #1
    While Not EOF(1)
        Line Input #1, linea
        LongMat = UBound(Matriz) + 1
        ReDim Preserve Matriz(LongMat)
        Matriz(LongMat).Direccion = ExtDir(linea)
        Matriz(LongMat).Nombre = ExtNom(linea)
    Wend
    Close #1
End If
pingear
End Sub


Sub pingear()
Dim Ping As cPing
Dim i As Integer
Dim Resultado As String

Set Ping = New cPing
Result.Clear
Resultado = Lpad("NOMBRE PC", C_LongNom, " ") & Chr(9) & Lpad("DIRECCION IP", C_LongDir, " ") & Chr(9)
Resultado = Resultado & Lpad("STATUS", 5, " ") & Chr(9) & Lpad("ERROR", 5, " ") & Chr(9) & Lpad("TIEMPO MLS", 10, " ")
Result.AddItem Resultado
Resultado = Lpad("-", 80, "-")
Result.AddItem Resultado
For i = 1 To UBound(Matriz)

    Ping.IPDestino = Trim(Matriz(i).Direccion)
    Ping.LongitudDatos = 200
    Ping.Ping
    
    Resultado = Matriz(i).Nombre & Chr(9) & Matriz(i).Direccion & Chr(9)
    Resultado = Resultado & Ping.Estado & Chr(9) & Ping.Descripcion & Chr(9) & Ping.Tiempo
    
    Result.AddItem Resultado
Next
End Sub


Private Sub Text1_Change()
Timer1.Interval = Val(Text1.Text) * 1000
End Sub

Private Sub Timer1_Timer()
If A.Visible Then
    A.Visible = False
    B.Visible = True
ElseIf B.Visible Then
    B.Visible = False
    C.Visible = True
ElseIf C.Visible Then
    C.Visible = False
    A.Visible = True
End If
pingear
End Sub
