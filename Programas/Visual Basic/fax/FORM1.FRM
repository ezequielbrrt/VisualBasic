VERSION 2.00
Begin Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   930
   ClientLeft      =   1560
   ClientTop       =   1500
   ClientWidth     =   8175
   Height          =   1335
   Left            =   1500
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   8175
   Top             =   1155
   Width           =   8295
   Begin CommandButton Command4 
      Caption         =   "Mandar un fax"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   1695
   End
   Begin CommandButton Command3 
      Caption         =   "Grabar Direccion"
      Height          =   495
      Left            =   6180
      TabIndex        =   3
      Top             =   180
      Width           =   1695
   End
   Begin CommandButton Command2 
      Caption         =   "Grabar Primer Fax"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   180
      Width           =   1695
   End
   Begin CommandButton Command1 
      Caption         =   "Ver Primer Fax"
      Height          =   495
      Left            =   2220
      TabIndex        =   0
      Top             =   180
      Width           =   1695
   End
   Begin Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      LinkTimeout     =   -1
      TabIndex        =   1
      Top             =   1020
      Visible         =   0   'False
      Width           =   1875
   End
End
Option Explicit

Sub Command1_Click ()
Dim i As Integer
On Error GoTo Error_DDE
Repetimos:
label1.LinkMode = 0
label1.LinkTimeout = -1
DoEvents
label1.LinkTopic = "DATAFAX|TOPIC"
label1.LinkMode = 2
'label1.LinkExecute "DF_ADDRESS(nombre de prueba, empresa, 5102303, ,,,,,,SAVE)"
label1.LinkExecute "DF_SHOWFAX(1,1)"
label1.LinkMode = 0
Exit Sub

Error_DDE:
label1.LinkMode = 0
If Err = 282 Then 'lanzar el captive
    i = Shell("d:\aplicwin\datafax\captive.exe", 4)
    DoEvents
    Resume Repetimos
Else
    Beep
    MsgBox "Error en DDE : ", Err
    End
End If


End Sub

Sub Command2_Click ()
Dim i As Integer
On Error GoTo Error_DDE2
Repetimos2:
label1.LinkMode = 0
label1.LinkTimeout = 50
DoEvents
label1.LinkTopic = "DATAFAX|TOPIC"
label1.LinkMode = 2
'label1.LinkExecute "DF_ADDRESS(nombre de prueba, empresa, 5102303, ,,,,,,SAVE)"
label1.LinkExecute "DF_SAVEFAX(1,1,0, , ,c:\prueba.bmp)"
label1.LinkMode = 0
i = Shell("pbrush c:\prueba.bmp", 1)
Exit Sub

Error_DDE2:
label1.LinkMode = 0
If Err = 282 Then 'lanzar el captive
    i = Shell("d:\aplicwin\datafax\captive.exe", 4)
    DoEvents
    Resume Repetimos2
Else
    Beep
    MsgBox "Error en DDE : ", Err
    End
End If



End Sub

Sub Command3_Click ()
Dim i As Integer
On Error GoTo Error_DDE3
Repetimos3:
label1.LinkMode = 0
label1.LinkTimeout = 50
DoEvents
label1.LinkTopic = "DATAFAX|TOPIC"
label1.LinkMode = 2
label1.LinkExecute "DF_ADDRESS(nombre de prueba,empresa,55,22,34,34,direccion,ciudad,0,SAVE)"
label1.LinkMode = 0
Exit Sub

Error_DDE3:
label1.LinkMode = 0
If Err = 282 Then 'lanzar el captive
    i = Shell("d:\aplicwin\datafax\captive.exe", 4)
    DoEvents
    Resume Repetimos3
Else
    Beep
    MsgBox "Error en DDE : ", Err
    End
End If



End Sub

Sub Command4_Click ()
Dim i As Integer
Dim NumErr As Integer
On Error GoTo Error_DDE4
Repetimos4:
label1.LinkMode = 0
label1.LinkTimeout = 50
DoEvents
label1.LinkTopic = "DATAFAX|TOPIC"
label1.LinkMode = 2
'poner datos destinatario
label1.LinkExecute "DF_ADDRESS(nombre de prueba, empresa, 5102303,,34,34,,,,SAVE)"
'elegir portada
label1.LinkExecute "DF_COVER(1,este es el mensaje de la portada)"
'informar de otros datos, SÓLO PORTADA
label1.LinkExecute "DF_CREATEINFO(1,,,,1)"
label1.LinkMode = 0
'escribir el resto del fax
On Error Resume Next
printer.Print
printer.EndDoc
Exit Sub

Error_DDE4:
NumErr = Err
label1.LinkMode = 0
If Err = 282 Then 'lanzar el captive
    i = Shell("d:\aplicwin\datafax\captive.exe", 4)
    DoEvents
    Resume Repetimos4
Else
    Beep
    MsgBox "Error en DDE : " & NumErr
    End
End If




End Sub

Sub Form_Load ()
Dim i As Integer
If (FindWindow("", "Servidor de DataFAX") = 0) Then
    i = Shell("C:\aplicwin\datafax\winserve.exe", 6)
End If
End Sub

Sub Label1_LinkClose ()
MsgBox "Cerrado enlace"
End Sub

Sub Label1_LinkError (LinkErr As Integer)
MsgBox "Error en dde : " & LinkErr, 0, "LinkError"
End Sub

Sub Label1_LinkNotify ()
label1.LinkRequest
End Sub

Sub Label1_LinkOpen (Cancel As Integer)
MsgBox "Establecido enlace"
End Sub

