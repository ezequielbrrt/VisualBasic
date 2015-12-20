VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1:  DBGrid - Tutorial 5"
   ClientHeight    =   5400
   ClientLeft      =   2190
   ClientTop       =   2295
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   4980
   Begin VB.CommandButton Command1 
      Caption         =   "Configure"
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   4800
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
      RecordSource    =   "Customer"
      Top             =   150
      Width           =   4815
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "TUTOR5.frx":0000
      Height          =   2040
      Left            =   75
      OleObjectBlob   =   "TUTOR5.frx":0010
      TabIndex        =   2
      Top             =   2700
      Width           =   4815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR5.frx":09C2
      Height          =   2040
      Left            =   75
      OleObjectBlob   =   "TUTOR5.frx":09D2
      TabIndex        =   1
      Top             =   525
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ' To configure TDBGrid2 to look the same as TDBGrid1.
    
    ' To configure DBGrid2 to look the same as DBGrid1.

    ' Set the AllowAddNew and AllowDelete properties of DBGrid2.
    DBGrid2.AllowAddNew = True
    DBGrid2.AllowDelete = True
    
    ' Instead of just adding 1 more column, we shall delete all
    ' existing columns and insert 3 new columns using the Remove
    ' and Add methods of the Columns Collection.  (We do this
    ' for the sake of illustrating a programming technique.)
    While DBGrid2.Columns.Count <> 0
        DBGrid2.Columns.Remove 0
    Wend
    DBGrid2.Columns.Add 0
    DBGrid2.Columns.Add 0
    DBGrid2.Columns.Add 0

    ' At run time, newly created columns are invisible.
    ' Make them visible by setting the Visible property.
    DBGrid2.Columns(0).Visible = True
    DBGrid2.Columns(1).Visible = True
    DBGrid2.Columns(2).Visible = True
   
    ' Set other column properties to make DBGrid2 look the same as
    ' DBGrid1.
    DBGrid2.Columns(0).DataField = "FirstName"
    DBGrid2.Columns(0).Caption = "First Name"
    DBGrid2.Columns(0).Alignment = dbgCenter
    DBGrid2.Columns(0).Width = DBGrid1.Columns(0).Width
    
    DBGrid2.Columns(1).DataField = "LastName"
    DBGrid2.Columns(1).Caption = "Last Name"
    DBGrid2.Columns(1).Alignment = dbgCenter
    DBGrid2.Columns(1).Width = DBGrid1.Columns(1).Width
    
    DBGrid2.Columns(2).DataField = "Contacted"
    DBGrid2.Columns(2).Caption = "Contact Date"
    DBGrid2.Columns(2).Alignment = dbgLeft
    DBGrid2.Columns(2).Width = DBGrid1.Columns(2).Width
    DBGrid2.Columns(2).NumberFormat = "Medium Date"
    
    ' Set the row height
    DBGrid2.RowHeight = DBGrid1.RowHeight
    
    ' Set the font
    DBGrid2.Font = DBGrid1.Font
End Sub

