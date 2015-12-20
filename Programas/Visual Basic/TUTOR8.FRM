VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1:  DBGrid - Tutorial 8"
   ClientHeight    =   3135
   ClientLeft      =   1500
   ClientTop       =   2085
   ClientWidth     =   7050
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
   ScaleWidth      =   7050
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   5325
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Customers"
      Connect         =   "Access"
      DatabaseName    =   ".\dbgtutor\TDBGDEMO.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   75
      Width           =   6795
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TUTOR8.frx":0000
      Height          =   2565
      Left            =   150
      OleObjectBlob   =   "TUTOR8.frx":0010
      TabIndex        =   1
      Top             =   450
      Width           =   6765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    ' BeforeColEdit is called before the grid enters into
    ' edit mode.  You can decide what happens and whether
    ' standard editing proceeds.  This allows you to
    ' substitute different kinds of editing for the current
    ' cell, as is done here.

    If DBGrid1.Columns(ColIndex).DataField = "CustType" Then
        ' Let the user edit by entering a key.
        If KeyAscii <> 0 Then Exit Sub
      
        ' Otherwise, cancel built-in editing and call the
        ' ButtonClick event to drop down List1.
        Cancel = True
        DBGrid1_ButtonClick (ColIndex)
    End If
End Sub

Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
    ' Assign the Column object to Co because it will be used
    ' more than once.
    Dim Co As Column
    Set Co = DBGrid1.Columns(ColIndex)
    
    ' Position and drop down List1 at the right edge of the
    ' current cell.
    List1.Left = DBGrid1.Left + Co.Left + Co.Width
    List1.Top = DBGrid1.Top + DBGrid1.RowTop(DBGrid1.Row)
    List1.Visible = True
    List1.ZOrder 0
    List1.SetFocus
End Sub

Private Sub DBGrid1_Scroll(Cancel As Integer)
    ' Hide the list if we scroll.
    List1.Visible = False
End Sub

Private Sub Form_Load()
    ' Add customer types to List1
    List1.AddItem "Prospective"
    List1.AddItem "Normal"
    List1.AddItem "Buyer"
    List1.AddItem "Distributor"
    List1.AddItem "Other"

    ' Place a button in the CustType column
    DBGrid1.Columns("CustType").Button = True
End Sub

Private Sub List1_DblClick()
    ' When an item is selected in List1, copy its index to the
    ' proper column in DBGrid1, then make List1 invisible.
    DBGrid1.Columns("CustType").Text = List1.ListIndex + 1
    List1.Visible = False
End Sub

Private Sub List1_LostFocus()
    ' Hide the list if it loses focus.
    List1.Visible = False
End Sub
