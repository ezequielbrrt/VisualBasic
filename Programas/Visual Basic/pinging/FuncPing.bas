Attribute VB_Name = "Module1"
Option Base 1

Global Const C_LongDir = 20
Global Const C_LongNom = 20

Type RegPing
    Direccion As String * C_LongDir
    Nombre As String * C_LongNom
End Type


'**************************************************
'Esta funcion rellena una expresion de tipo string
'Texto  = Valor a rellenar
'Largo  = Largo del string a devolver
'char   = Caracter de relleno
'**************************************************
Function Lpad(Texto As String, largo As Byte, char As String) As String
Texto = Trim(Texto)
leng = Len(Texto)
Lpad = Texto
For i = 1 To largo - leng
    Lpad = Lpad & char
Next
End Function


Function At(Texto As String, Caracter As String)
At = 0
For i = 1 To Len(Texto)
    If Mid(Texto, i, 1) = Caracter Then
        At = i
        Exit For
    End If
Next
End Function


Function ExtDir(linea As String)
ExtDir = Mid(linea, 1, At(linea, ",") - 1)
End Function

Function ExtNom(linea As String)
ExtNom = Mid(linea, At(linea, ",") + 1)
End Function

