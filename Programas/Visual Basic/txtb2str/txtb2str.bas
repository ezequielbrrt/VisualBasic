Sub TextBoxAStrings (ByVal TextoTB As String, Linea() As String, NumCar As Integer, NumLin As Integer, Ocupadas As Integer)
'Parametros:
    'TextoTB : el texto del textbox
    'Linea() : matriz de string en la que quedara el texto
    'NumCar : nº de caracteres por linea que queremos
    'NumLin : nº maximo de lineas a rellenar
    'Ocupadas : devuelve el nº de lineas realmente utilizadas
Dim i As Integer, j As Integer, k As Integer, L As Integer
For i = 1 To NumLin
    Linea(i) = Space$(NumCar + 1)
Next i
j = 1
i = 1
k = 1
While i <= Len(TextoTB) And j <= NumLin
    If Asc(Mid$(TextoTB, i, 1)) = 13 Then
    i = i + 2
    j = j + 1
    k = 1
    Else
    Mid$(Linea(j), k, 1) = Mid$(TextoTB, i, 1)
    k = k + 1
    i = i + 1
   End If
   If k > NumCar + 1 Then
    If Mid$(Linea(j), k - 1, 1) = " " Then
    Else
        L = k - 1
        While Mid$(Linea(j), L, 1) <> " " And L > 0
        L = L - 1
        Wend
        If L > 0 Then
        i = i - (k - L) + 1
        Linea(j) = Left$(Linea(j), L)
        End If
    End If
    j = j + 1
    k = 1
    End If
Wend
If j > NumLin Then
    Ocupadas = NumLin
Else
    Ocupadas = j
End If
End Sub
