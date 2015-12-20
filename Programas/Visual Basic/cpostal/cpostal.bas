'
'gsCPostal.bas Declaración de tipos y variables globales
'en VB4 no es necesario usar este módulo ya que se pueden
'declarar los tipos de datos de forma privada en los forms
'
Option Explicit

Type tipoProvDAT                'Para el fichero Prov.DAT
    Prov As String
    miniCP As String
End Type
'----------------------------------------------------------
'NOTA: los ficheros xxPOB.dat (poblaciones)     (19/Abr/94)
'                 y xxDIR.dat (direcciones de la capital),
'tienen la mísma longitud de datos y las mísmas posiciones.
'Por ejemplo el campo "Poblacion":
'   en población hace referencia a la población,
'   en dirección hace referencia a la calle.
'----------------------------------------------------------
Type tipoPobDAT                 'Longitud 83 caracteres
    PrvLen As String * 1
    Provincia As String * 17
    PobLen As String * 1
    Poblacion As String * 58
    CPLen As String * 1
    CPostal As String * 5
End Type

Global Provincia As String
Global HayCP As Integer

' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Sub AbrirCPfrm ()
    'Abrir y mostrar la ventana de Códigos Postales
    'Comprobar si ya está en memoria...
    If HayCP Then
        If CPfrm.WindowState = 0 Then
            CPfrm.WindowState = 1
        Else
            CPfrm.WindowState = 0
        End If
    Else
        CPfrm.Show
    End If

End Sub

Sub Centrar (frm As Form)
frm.Move (screen.Width - frm.Width) / 2, (screen.Height - frm.Height) / 2
End Sub

Sub OutLines (formname As Form)
    Dim drkgray As Long, fullwhite As Long
    Dim i As Integer
    Dim ctop As Integer, cleft As Integer, cright As Integer, cbottom As Integer

    ' Outline a form's controls for 3D look unless control's TAG
    ' property is set to "skip".

    Dim cname As Control
    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    For i = 0 To (formname.Controls.Count - 1)
        Set cname = formname.Controls(i)
        If TypeOf cname Is Menu Then
            'Debug.Print "menu item"
        'ElseIf (UCase(cname.Tag) = "OL") Then
        ElseIf (InStr(1, cname.Tag, "OL", 1) > 0) Then
                ctop = cname.Top - screen.TwipsPerPixelY
                cleft = cname.Left - screen.TwipsPerPixelX
                cright = cname.Left + cname.Width
                cbottom = cname.Top + cname.Height
                formname.Line (cleft, ctop)-(cright, ctop), drkgray
                formname.Line (cleft, ctop)-(cleft, cbottom), drkgray
                formname.Line (cleft, cbottom)-(cright, cbottom), fullwhite
                formname.Line (cright, ctop)-(cright, cbottom), fullwhite
        End If
    Next i
End Sub

Sub WOutLines (formname As Form)
    Dim drkgray As Long, fullwhite As Long

    drkgray = RGB(128, 128, 128)
    fullwhite = RGB(255, 255, 255)

    formname.DrawWidth = 2
    formname.Line (0, 0)-(formname.Width, 0), fullwhite
    formname.Line (0, 0)-(0, formname.Height), fullwhite
    formname.Line (0, formname.Height - 320)-(formname.Width - 50, formname.Height - 320), drkgray
    formname.Line (formname.Width - 30, 0)-(formname.Width - 30, formname.Height - 30), drkgray
    formname.DrawWidth = 1
End Sub

