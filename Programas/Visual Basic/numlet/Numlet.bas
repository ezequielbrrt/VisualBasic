Option Explicit
Dim Unidades$(9), Decenas$(9), Oncenas$(9)
Dim Veintes$(9), Centenas$(9)

Function Descifrar$ (numero%)
Static SAL$(4)
Dim I%, CT As Double, DC As Double, DU As Double, UD  As Double
Dim VARIABLE$

For I% = 1 To 4: SAL$(I%) = " ": Next I%
VARIABLE$ = String$(3 - Len(Trim$(Str$(numero%))), "0") + Trim$(Str$(numero%))
CT = Val(Mid$(VARIABLE$, 1, 1)): '*** CENTENA
DC = Val(Mid$(VARIABLE$, 2, 1)): '*** DECENA
DU = Val(Mid$(VARIABLE$, 2, 2)): '*** DECENA + UNIDAD
UD = Val(Mid$(VARIABLE$, 3, 1)): '*** UNIDAD
If numero% = 100 Then
        SAL$(1) = "CIEN "
Else
        If CT <> 0 Then SAL$(1) = Centenas$(CT)
        If DC <> 0 Then
                If DU <> 10 And DU <> 20 Then
                        If DC = 1 Then SAL$(2) = Oncenas$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)): Exit Function
                        If DC = 2 Then SAL$(2) = Veintes$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)): Exit Function
                End If
                SAL$(2) = " " + Decenas$(DC)
                If UD <> 0 Then SAL$(3) = "Y "
        End If
        If UD <> 0 Then SAL$(4) = Unidades$(UD)
End If
Descifrar = Trim$(SAL$(1) + SAL$(2) + SAL$(3) + SAL$(4))
End Function

Function Numlet$ (NUM#)
Dim DEC$, MILM$, MILL$, MILE$, UNID$
ReDim SALI$(11)
Dim var$, I%, AUX$
'NUM# = Round(NUM#, 2)
var$ = Trim$(Str$(NUM#))
If InStr(var$, ".") = 0 Then
        var$ = var$ + ".00"
End If
If InStr(var$, ".") = Len(var$) - 1 Then
        var$ = var$ + "0"
End If
var$ = String$(15 - Len(LTrim$(var$)), "0") + LTrim$(var$)
DEC$ = Mid$(var$, 14, 2)
MILM$ = Mid$(var$, 1, 3)
MILL$ = Mid$(var$, 4, 3)
MILE$ = Mid$(var$, 7, 3)
UNID$ = Mid$(var$, 10, 3)
For I% = 1 To 11: SALI$(I%) = " ": Next I%
I% = 0
Unidades$(1) = "UNA    "
Unidades$(2) = "DOS    "
Unidades$(3) = "TRES   "
Unidades$(4) = "CUATRO "
Unidades$(5) = "CINCO  "
Unidades$(6) = "SEIS   "
Unidades$(7) = "SIETE  "
Unidades$(8) = "OCHO   "
Unidades$(9) = "NUEVE  "

Decenas$(1) = "DIEZ      "
Decenas$(2) = "VEINTE    "
Decenas$(3) = "TREINTA "
Decenas$(4) = "CUARENTA "
Decenas$(5) = "CINCUENTA "
Decenas$(6) = "SESENTA "
Decenas$(7) = "SETENTA "
Decenas$(8) = "OCHENTA "
Decenas$(9) = "NOVENTA "

Oncenas$(1) = "ONCE       "
Oncenas$(2) = "DOCE       "
Oncenas$(3) = "TRECE      "
Oncenas$(4) = "CATORCE    "
Oncenas$(5) = "QUINCE     "
Oncenas$(6) = "DIECISEIS  "
Oncenas$(7) = "DIECISIETE "
Oncenas$(8) = "DIECIOCHO  "
Oncenas$(9) = "DIECINUEVE "

Veintes$(1) = "VEINTIUNA    "
Veintes$(2) = "VEINTIDOS    "
Veintes$(3) = "VEINTITRES   "
Veintes$(4) = "VEINTICUATRO "
Veintes$(5) = "VEINTICINCO  "
Veintes$(6) = "VEINTISEIS   "
Veintes$(7) = "VEINTISIETE  "
Veintes$(8) = "VEINTIOCHO   "
Veintes$(9) = "VEINTINUEVE  "

Centenas$(1) = "       CIENTO "
Centenas$(2) = "   DOSCIENTOS "
Centenas$(3) = "  TRESCIENTOS "
Centenas$(4) = "CUATROCIENTOS "
Centenas$(5) = "   QUINIENTOS "
Centenas$(6) = "  SEISCIENTOS "
Centenas$(7) = "  SETECIENTOS "
Centenas$(8) = "  OCHOCIENTOS "
Centenas$(9) = "  NOVECIENTOS "

If NUM# > 999999999999.99 Then Numlet$ = " ": Exit Function
If Val(MILM$) >= 1 Then
        SALI$(2) = " MIL ": '** MILES DE MILLONES
        SALI$(4) = " MILLONES "
        If Val(MILM$) <> 1 Then
                Unidades$(1) = "UN     "
                Veintes$(1) = "VEINTIUN     "
                SALI$(1) = Descifrar$(Val(MILM$))
        End If
End If
If Val(MILL$) >= 1 Then
        If Val(MILL$) < 2 Then
                SALI$(3) = "UN ": '*** UN MILLON
                If Trim$(SALI$(4)) <> "MILLONES" Then
                        SALI$(4) = " MILLON "
                End If
        Else
                SALI$(4) = " MILLONES ": '*** VARIOS MILLONES
                Unidades$(1) = "UN     "
                Veintes$(1) = "VEINTIUN     "
                SALI$(3) = Descifrar$(Val(MILL$))
        End If
End If
For I% = 2 To 9
        Centenas$(I%) = Mid$(Centenas(I%), 1, 11) + "AS"
Next I%
If Val(MILE$) > 0 Then
        SALI$(6) = " MIL ": '*** MILES
        If Val(MILE$) <> 1 Then
                SALI$(5) = Descifrar$(Val(MILE$))
        End If
End If
Unidades$(1) = "UNA    "
Veintes$(1) = "VEINTIUNA"
If Val(UNID$) >= 1 Then
        SALI$(7) = Descifrar$(Val(UNID$)):  '*** CIENTOS
        If Val(DEC$) >= 10 Then
            SALI$(8) = " CON ": '*** DECIMALES
            SALI$(10) = Descifrar$(Val(DEC$))
        End If
End If
If Val(MILM$) = 0 And Val(MILL$) = 0 And Val(MILE$) = 0 And Val(UNID$) = 0 Then SALI$(7) = " CERO "
AUX$ = ""
For I% = 1 To 11
        AUX$ = AUX$ + SALI$(I%)
Next I%
Numlet$ = Trim$(AUX$)
End Function

