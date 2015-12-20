VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1:  DBGrid - Tutorial 4"
   ClientHeight    =   5325
   ClientLeft      =   1845
   ClientTop       =   2265
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5325
   ScaleWidth      =   6240
   Begin VB.TextBox Text3 
      DataField       =   "Company"
      DataSource      =   "Data1"
      Height          =   390
      Left            =   2100
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4800
      Width           =   2715
   End
   Begin VB.TextBox Text2 
      DataField       =   "LastName"
      DataSource      =   "Data1"
      Height          =   390
      Left            =   2100
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   3975
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      Height          =   390
      Left            =   2100
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3150
      Width           =   2715
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ".\dbgtutor\TDBGDEMO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   75
      Width           =   4740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4950
      TabIndex        =   6
      Top             =   525
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4950
      TabIndex        =   5
      Top             =   1125
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4950
      TabIndex        =   4
      Top             =   1725
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4950
      TabIndex        =   3
      Top             =   2925
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4950
      TabIndex        =   2
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4950
      TabIndex        =   1
      Top             =   4125
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4950
      TabIndex        =   0
      Top             =   4725
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR4.frx":0000
      Height          =   2190
      Left            =   75
      OleObjectBlob   =   "TUTOR4.frx":0010
      TabIndex        =   15
      Top             =   525
      Width           =   4740
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "TUTOR4.frx":09C2
      DataField       =   "LastName"
      DataSource      =   "Data1"
      Height          =   2010
      Left            =   75
      TabIndex        =   10
      Top             =   3150
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   3545
      _Version        =   327680
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ListField       =   "LastName"
      BoundColumn     =   "LastName"
   End
   Begin VB.Label Label4 
      Caption         =   "Company:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2100
      TabIndex        =   14
      Top             =   4500
      Width           =   2040
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2100
      TabIndex        =   13
      Top             =   3675
      Width           =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2100
      TabIndex        =   12
      Top             =   2850
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name List:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   11
      Top             =   2850
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ' True DBGrid will automatically respond to the
    ' update and will clear the "modified indicator"
    ' (the pencil icon) on the record selector column
    ' to indicate that the modified data has been written
    ' to the database.

    Data1.Recordset.Edit
    Data1.Recordset.Update
    DBGrid1.SetFocus

End Sub

Private Sub Command2_Click()
    ' When the current record is deleted, Jet Engine leaves
    ' the record pointer at the deleted record.  Use MoveNext
    ' to move the current record to the row after the deleted
    ' record.
    
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
    
    ' If the last record was deleted, move to the previous
    ' record (which is now the last record).
    
    If Data1.Recordset.EOF = True Then
        Data1.Recordset.MovePrevious
    End If
    DBGrid1.SetFocus
 End Sub

Private Sub Command3_Click()
    ' This "Add New" button moves the cursor to the
    ' "new (blank) row" at the end so that user can start
    ' adding data to the new record.
           
    ' Move to the last record so that the grid's AddNew row
    ' becomes visible.
    Data1.Recordset.MoveLast
    
    ' Move the cursor to the AddNew row.
    DBGrid1.Row = DBGrid1.Row + 1
    DBGrid1.SetFocus
End Sub

Private Sub Command4_Click()
    ' True DBGrid will follow the record movement.
    Data1.Recordset.MoveFirst
    DBGrid1.SetFocus
End Sub

Private Sub Command5_Click()
    ' True DBGrid will follow the record movement.

    Data1.Recordset.MoveNext
    
    ' Keep the current record away from EOF, which is not a
    ' valid position.
    If Data1.Recordset.EOF = True Then
        Data1.Recordset.MovePrevious
    End If
    
    DBGrid1.SetFocus
End Sub

Private Sub Command6_Click()
    ' True DBGrid will follow the record movement.

    Data1.Recordset.MovePrevious
    
    ' Keep the current record away from BOF, which is not a
    ' valid position.
    If Data1.Recordset.BOF = True Then
        Data1.Recordset.MoveNext
    End If
    
    DBGrid1.SetFocus
End Sub

Private Sub Command7_Click()
    ' True DBGrid will follow the record movement.

    Data1.Recordset.MoveLast
    DBGrid1.SetFocus
End Sub

