VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1:  DBGrid - Tutorial 6"
   ClientHeight    =   3360
   ClientLeft      =   2025
   ClientTop       =   2310
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   5430
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   390
      Left            =   1350
      TabIndex        =   1
      Top             =   2850
      Width           =   1140
   End
   Begin VB.Data Data1 
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
      RecordSource    =   "select * from composer, opus, composer inner join opus on composer.last = opus.last"
      Top             =   75
      Width           =   5265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   390
      Left            =   75
      TabIndex        =   0
      Top             =   2850
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR6.frx":0000
      Height          =   2340
      Left            =   75
      OleObjectBlob   =   "TUTOR6.frx":0010
      TabIndex        =   2
      Top             =   450
      Width           =   5265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ' This routine loops through the Recordset to find and
    ' highlight all records with Country = "Germany"
    
    ' We shall use a clone so that we do not move the actual
    ' record position of the Data control
    Dim dclone As Recordset
    Set dclone = Data1.Recordset.Clone()

    ' In case there is a large Recordset to search through
    Screen.MousePointer = vbHourglass
    
    ' For each matching record, add the bookmark to the
    ' SelBookmarks collection of the grid.  The grid will
    ' highlight the corresponding rows.  Note that the bookmarks
    ' of a clone are compatible with the original set.
    ' This is ONLY true of clones.
    Dim SelBks As SelBookmarks
    Set SelBks = DBGrid1.SelBookmarks

    Dim Criteria$
    Criteria$ = "Country = " & Chr$(34) & "Germany" & Chr$(34)
    dclone.FindFirst Criteria$
    While Not dclone.NoMatch
        SelBks.Add dclone.Bookmark
        dclone.FindNext Criteria$
    Wend
    
    ' Restore regular mouse pointer
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    ' Clear all selected rows by removing the selected records from
    ' the SelBookmarks collection.

    Dim SelBks As SelBookmarks
    Set SelBks = DBGrid1.SelBookmarks

    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
End Sub

