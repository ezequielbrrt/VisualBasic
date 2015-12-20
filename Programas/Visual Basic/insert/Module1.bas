Option Explicit
Global gInsert As Integer

Sub ControlaInsert (texto As Control, tecla As Integer, shift As Integer, insertar As Integer)
If shift > -1 Then  'evento keydown
    If tecla = 45 Then insertar = Not insertar  'tecla insert
    Exit Sub
End If
If insertar Or texto.SelLength > 0 Then Exit Sub 'emplear el comportamiento por defecto
texto.SelLength = 1 'marco el siguiente caracter al cursor
texto.SelText = Chr$(tecla) 'lo sustituyo por el que teclean
texto.SelLength = 0 'no selecciono ningun caracter
tecla = 0   'como procese la tecla la quito para que no se escriba de nuevo
End Sub

