VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1:  DBGrid - Tutorial 3"
   ClientHeight    =   4590
   ClientLeft      =   2205
   ClientTop       =   1980
   ClientWidth     =   5130
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
   ScaleHeight     =   4590
   ScaleWidth      =   5130
   Begin VB.Data Data2 
      Appearance      =   0  'Flat
      Caption         =   "Their Compositions"
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
      RecordSource    =   "OPUS"
      Top             =   2400
      Width           =   4965
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Composers"
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
      RecordSource    =   "COMPOSER"
      Top             =   75
      Width           =   4965
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "TUTOR3.frx":0000
      Height          =   1740
      Left            =   75
      OleObjectBlob   =   "TUTOR3.frx":0010
      TabIndex        =   1
      Top             =   2775
      Width           =   4965
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR3.frx":09C2
      Height          =   1815
      Left            =   75
      OleObjectBlob   =   "TUTOR3.frx":09D2
      TabIndex        =   0
      Top             =   450
      Width           =   4965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' A query is performed by taking the LAST name field from
    ' the Data1 control and building an SQL query on the LAST
    ' name field in the Data2 (compositions) file.

    ' The second grid will respond automatically when the Data
    ' Control causes the change.  We put up an hourglass so that
    ' there's a bit of feedback if Access is slow at finishing
    ' the query.

    ' To reliably compare bookmarks, you must first convert them
    ' into strings.  LastRow is the bookmark of the row before
    ' the RowColChange takes place.

    Dim bk1 As String, bk2 As String, lastname As String
    bk1 = LastRow
    bk2 = DBGrid1.Bookmark

    If bk1 <> bk2 Then
        Screen.MousePointer = vbHourglass
    
        ' If the row has changed in DBGrid1, make sure that
        ' DBGrid2 displays only compositions by the composer at
        ' the new row of DBGrid1.

        lastname = Data1.Recordset("Last")
        Data2.RecordSource = "SELECT * FROM OPUS WHERE LAST = " _
                             & Chr$(34) & lastname & Chr$(34)
        Data2.Refresh
        Screen.MousePointer = vbDefault
    End If
End Sub
