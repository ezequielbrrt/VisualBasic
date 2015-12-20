VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1:  DBGrid - Tutorial 2"
   ClientHeight    =   3675
   ClientLeft      =   2430
   ClientTop       =   2070
   ClientWidth     =   4890
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
   ScaleHeight     =   3675
   ScaleWidth      =   4890
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "TUTOR2.frx":0000
      Top             =   2760
      Width           =   4740
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Execute SQL"
      Height          =   390
      Left            =   60
      TabIndex        =   1
      Top             =   2325
      Width           =   1305
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ".\dbgtutor\TDBGDEMO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   75
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CUSTOMER"
      Top             =   75
      Width           =   4755
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR2.frx":001C
      Height          =   1815
      Left            =   75
      OleObjectBlob   =   "TUTOR2.frx":002C
      TabIndex        =   2
      Top             =   450
      Width           =   4740
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ' Execute the SQL statement in Text1, and trigger an error
    ' message if something goes wrong.
    On Error GoTo SQLErr
    Data1.RecordSource = Text1.Text
    Data1.Refresh
    DBGrid1.SetFocus
    Exit Sub

SQLErr:
    MsgBox "Error Executing SQL Statement"
    Exit Sub
End Sub


