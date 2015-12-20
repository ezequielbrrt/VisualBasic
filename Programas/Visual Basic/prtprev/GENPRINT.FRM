VERSION 2.00
Begin Form frmPreview 
   Caption         =   "Print or Preview"
   ClientHeight    =   7104
   ClientLeft      =   528
   ClientTop       =   1872
   ClientWidth     =   7440
   Height          =   7524
   Left            =   480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7104
   ScaleWidth      =   7440
   Top             =   1500
   Width           =   7536
   Begin CheckBox Check1 
      Caption         =   "Pre&view"
      Height          =   252
      Left            =   5640
      TabIndex        =   3
      Top             =   3000
      Width           =   1452
   End
   Begin CommandButton Command1 
      Caption         =   "&Print"
      Height          =   492
      Left            =   5640
      TabIndex        =   2
      Top             =   3480
      Width           =   1332
   End
   Begin PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   4068
      Left            =   5760
      Picture         =   GENPRINT.FRX:0000
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   444
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   5352
   End
   Begin PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6612
      Left            =   120
      Picture         =   GENPRINT.FRX:0E94
      ScaleHeight     =   6588
      ScaleWidth      =   5388
      TabIndex        =   0
      Top             =   240
      Width           =   5412
   End
End

Sub Command1_Click ()
   
   
   'Setup (Could be done at design time or in form load)
   'Make printing stick
   'Picture1.AutoRedraw = True
   'Add a palette for 256 colors
   'Picture1.Picture = LoadPicture("C:\VB\PASTEL.DIB")
   'Setup hidden picture
   'Picture2.AutoRedraw = False
   'Picture2.ScaleMode = 3 'Pixels
   'Picture2.Visible = False
   'Picture2.AutoSize = True
   'Picture2.Picture = LoadPicture("C:\VB\METAFILE\BUSINESS\PRINTER.WMF")
   


   'This print job can go to the printer or the picture box
   If Check1.Value = 0 Then PrinterFlag = True
   PrintStartDoc Picture1, PrinterFlag, 8.5, 11
   
   'All the subs use inches
   PrintBox 1, 1, 6.5, 9
   PrintLine 1.1, 2, 7.4, 2
   PrintPicture Picture2, 1.1, 1.1, .8, .8
   PrintFilledBox 2.1, 1.2, 5.2, .7, RGB(200, 200, 200)
   PrintFontName "Arial"
   PrintCurrentX 2.3
   PrintCurrentY 1.3
   PrintFontSize 35
   PrintPrint "Visual Basic Printing"
   For x = 3 To 5.5 Step .2
      PrintCircle x, 3.5, .75
   Next
   PrintFontName "Courier New"
   PrintFontSize 30
   PrintCurrentX 1.5
   PrintCurrentY 5
   PrintPrint "It is possible to do"
   PrintFontSize 24
   PrintCurrentX 1.5
   PrintCurrentY 6.5
   PrintPrint "It is possible to do print"
   PrintFontSize 18
   PrintCurrentX 1.5
   PrintCurrentY 8
   PrintPrint "It is possible to do print preview"
   PrintFontSize 12
   PrintCurrentX 1.5
   PrintCurrentY 9.5
   PrintPrint "It is possible to do print preview with good results."
   PrintEndDoc

End Sub

