VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1:  DBGrid - Tutorial 10"
   ClientHeight    =   3135
   ClientLeft      =   1950
   ClientTop       =   2085
   ClientWidth     =   7365
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
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   7365
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ".\dbgtutor\TDBGDEMO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Composer"
      Top             =   75
      Width           =   7065
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR10.frx":0000
      Height          =   2490
      Left            =   150
      OleObjectBlob   =   "TUTOR10.frx":0010
      TabIndex        =   0
      Top             =   525
      Width           =   7065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Before modifying the grid's properties, make sure the
    ' grid is initialized by refreshing the Data control.
    Data1.Refresh
      
    ' Create an additional splits:
    DBGrid1.Splits.Add 0 ' Create an additional split

    ' Hide all columns in the leftmost split, Splits(0),
    ' except for columns 0 and 1
    Dim Cols As Columns, C As Column
    Set Cols = DBGrid1.Splits(0).Columns
    For Each C In Cols
        C.Visible = False
    Next C
    Cols(0).Visible = True
    Cols(1).Visible = True

    ' Configure Splits(0) to display exactly two columns,
    ' and disable resizing
    With DBGrid1.Splits(0)
        .SizeMode = dbgNumberOfColumns
        .Size = 2
        .AllowSizing = False
      End With

    ' Usually, if you fix columns 0 and 1 from scrolling
    ' in a split, you will want to make them invisible in
    ' other splits:
    Set Cols = DBGrid1.Splits(1).Columns
    Cols(0).Visible = False
    Cols(1).Visible = False
    
    ' Turn off the record selectors in Split 1:
    DBGrid1.Splits(1).RecordSelectors = False
End Sub
