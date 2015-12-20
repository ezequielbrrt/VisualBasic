VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1:  DBGrid - Tutorial 1"
   ClientHeight    =   3030
   ClientLeft      =   1710
   ClientTop       =   2160
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3030
   ScaleWidth      =   5490
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ".\dbgtutor\TDBGDEMO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   75
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Composer"
      Top             =   75
      Width           =   5340
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR1.frx":0000
      Height          =   2340
      Left            =   75
      OleObjectBlob   =   "TUTOR1.frx":0010
      TabIndex        =   0
      Top             =   525
      Width           =   5340
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
