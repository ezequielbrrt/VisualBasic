Option Explicit

Function CalculaCCC (entidad As String, oficina As String, cuenta As String) As String
Dim i As Integer
Dim j As Currency
Dim dig1 As String
Dim dig2 As String
If entidad = "" Then entidad = "0"
If oficina = "" Then oficina = "0"
If cuenta = "" Then cuenta = "0"
i = CInt(entidad)
entidad = Format$(i, "0000")
i = CInt(oficina)
oficina = Format$(i, "0000")
j = CCur(cuenta)
cuenta = Format$(j, "0000000000")
dig1 = chkpesos(entidad + oficina)
dig2 = chkpesos(cuenta)
CalculaCCC = dig1 + dig2
End Function

Function CalculaLetraDNI (dni As Variant) As String
Const TablaLetras = "TRWAGMYFPDXBNJZSQVHLCKE"
Dim numero As Long, numero2 As Long

On Error Resume Next
CalculaLetraDNI = ""
If Trim$(dni) = "" Then Exit Function
If Len(TablaLetras) <> 23 Then
    Beep
    MsgBox "Tabla de letras erronea"
    Exit Function
End If
If Not IsNumeric(dni) Then
    Beep
    MsgBox "Dni no numérico"
    Exit Function
End If

numero = CLng("0" & dni)
If numero < 23 Then Exit Function
numero2 = Int(numero / 23)
numero = numero - (numero2 * 23) + 1
CalculaLetraDNI = Mid$(TablaLetras, numero, 1)
End Function

Function chkpesos (cadena As String) As String
ReDim pesos(10) As Integer
Dim i As Integer
Dim j As Integer
Dim total As Currency

pesos(1) = 1: pesos(2) = 2: pesos(3) = 4: pesos(4) = 8
pesos(5) = 5: pesos(6) = 10: pesos(7) = 9: pesos(8) = 7
pesos(9) = 3: pesos(10) = 6

total = 0
j = 10
For i = Len(cadena) To 1 Step -1
    total = total + (CInt(Mid$(cadena, i, 1)) * pesos(j))
    j = j - 1
Next i
i = 11 - (total Mod 11)
If i = 10 Then
    i = 1
Else
    If i = 11 Then
        i = 0
    End If
End If
chkpesos = CStr(i)
End Function

