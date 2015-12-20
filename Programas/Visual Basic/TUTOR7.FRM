VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1:  DBGrid - Tutorial 7"
   ClientHeight    =   2700
   ClientLeft      =   2985
   ClientTop       =   2115
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2700
   ScaleWidth      =   4065
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   2415
      Left            =   150
      OleObjectBlob   =   "TUTOR7.frx":0000
      TabIndex        =   0
      Top             =   150
      Width           =   3765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MaxCol = 2
Const MaxRow = 100
Dim MyData(MaxCol, MaxRow) As Variant

Private Sub DBGrid1_UnboundGetRelativeBookmark(StartLocation As Variant, ByVal OffSet As Long, NewLocation As Variant, ApproximatePosition As Long)
    ' DBGrid1 calls this routine each time it
    ' needs to position a row.

    ' StartLocation is a bookmark supplied by the grid to
    ' indicate which records the grid wants to positioning from.
    ' If StartLocation is NULL, then data is requested either
    ' from the first row or the last row.
    Dim curRow%
    If IsNull(StartLocation) Then
        If OffSet < 0 Then
            curRow% = OffSet + MaxRow ' Reading data from the last row
        Else
            curRow% = OffSet - 1      ' Reading data from the first row
        End If
    Else
        ' Positioning from StartLocation plus the OffSet
        curRow% = Val(StartLocation) + OffSet
    End If

    If curRow% < 0 Or curRow% >= MaxRow Then
        NewLocation = Null
    Else
       NewLocation = curRow%
       ApproximatePosition = curRow% + 1  ' One based
    End If
End Sub

Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
    ' DBGrid1 calls this routine each time it needs data.

    ' Keep track of how many rows have been provided by the
    ' programmer to the grid.
    Dim rowsFetched%
    rowsFetched% = 0
    
    ' See if the grid is requesting data forward or backward:
    '   If ReadPriorRows = True, then the grid is requesting
    '   rows prior to the StartLocation (backward).
    '   If ReadPriorRows = False, then the grid is requesting
    '   rows after the StartLocation (forward).
    Dim incr%
    If ReadPriorRows Then
        incr% = -1      ' Reading data backward
    Else
        incr% = 1       ' Reading data forward
    End If

    ' StartLocation is a bookmark supplied by the grid to
    ' indicate which records the grid wants to read.
    ' If StartLocation is NULL, then data is requested either
    ' from the first row or the last row.
    Dim curRow%
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
           curRow% = MaxRow - 1  ' Reading data from the last row
        Else
            curRow% = 0           ' Reading data from the first row
        End If
    Else
        ' Reading data starting from the row next (previous) to
        ' StartLocation
        curRow% = Val(StartLocation) + incr%
    End If

    Dim i%
    i% = 0

    While i% < RowBuf.RowCount
        If curRow% < 0 Or curRow% >= MaxRow Then
            ' No more data to read, return RowCount and exit.  Note
            ' that the grid does not need to know the total number of
            ' rows ahead of time.  This is an important concept because
            ' this enables the unbound grid to operate under a virtual
            ' mode (i.e., it can be used to display unlimited rows of
            ' data) for any data you wish to display.
            RowBuf.RowCount = rowsFetched%
            Exit Sub
        End If

        ' Return the row data to the grid through the Row Buffer
        ' (RowBuf) object.
        RowBuf.Value(i%, 0) = MyData(0, curRow%)
        RowBuf.Value(i%, 1) = MyData(1, curRow%)

        ' Use the array row index as a bookmark and return it to
        ' the grid.
        RowBuf.Bookmark(i%) = curRow%

        curRow% = curRow% + incr%
        rowsFetched% = rowsFetched% + 1
        i% = i% + 1
    Wend

    ' Return the number of rows of data fetched
    RowBuf.RowCount = rowsFetched%
End Sub

Private Sub Form_Load()
    ' Set approximate count value:
    DBGrid1.ApproxCount = MaxRow

   ' Initialize MyData
    Dim i%, j%
    For i% = 0 To MaxCol - 1
        For j% = 0 To MaxRow - 1
            MyData(i%, j%) = "Col" & Str$(i%) & ", Row" & Str$(j%)
        Next j%
    Next i%
End Sub
