Attribute VB_Name = "ModuloResize"
Option Explicit

Type Monitor
   Alto As Integer
   Ancho As Integer
End Type
Public Pantalla As Monitor

Public FactorX As Single, FactorY As Single

'Módulo escrito por Oscar Manuel Gómez Senovilla (oscarsen@cros.es)
'Módulo Resize.bas escrito en Visual Basic 4, en principio válido para cualquier versión de VB
'El objetivo de este módulo es no tener que preocuparse por la resolución en la cual
'ha sido diseñado un formulario, así como la si la resolución de la pantalla de
'otro ordenador donde se vaya a ejecutar es distinta a la de la la pantalla en que fue diseñado
'
'Consejos y recomendaciones:
'   - Usar fuentes escalables, como la Arial, en vez de la fuente por defecto MS Sans Serif. Si existe
'     la seguridad de que la resolución de origen y destino es la misma, no habrá problemas, ya que la rutina,
'     para ganar tiempo, comprueba si coinciden las resoluciones, en cuyo caso simplemente no hará ningún reajuste.
'     Si hay un reajuste y se usa la fuente Sans Serif, se producirán resultados indeseados, y que se deben
'     a que dicha fuente no soporta la escalabilidad necesaria.
'
'   - Trabajar siempre con el módulo, y no con una copia (por llamarlo de alguna manera), desde el cual, al crear un
'     proyecto, se añada, y si se le encuentra algún fallo con algún control que no ha sido probado,
'     se hacen las modificaciones y pruebas oportunas (con la precaución correspondiente),
'     y dichas modificaciones estarán disponibles para los proyectos pasados y futuros
'
'
'Para usarlo:
'  1) Añadir este módulo al proyecto
'
'  2) Hay que inicializar las variables. en el formulario de inicio, se pueden inicializar las variables.
'     Pantalla.Ancho (normalmente = 800) y Pantalla.Alto (normalmente = 600).
'     Cuando se llegue a un formulario que ha sido diseñado en una resolución distinta, habrá que cambiar los
'     valores de las variables por los de la resolución correspondiente, sin olvidar reponerlos después,
'     siendo recomendable, si hay varios cambios de resolución en los formularios, que cada uno tenga en su Form_Load
'     sus propios valores. Si sólo hay uno o dos formularios con otra resolución distinta, pueden reponerse los valores
'     inmediatamente después de llamar a la rutina Ajustar (se explica en el paso 4).
'     Si todos los formularios han sido diseñados en la misma resolución, sólo es necesario inicializar
'     una vez las variables Pantalla.Ancho y Pantalla.Alto, permaneciendo dicho valor permanentemente
'     durante toda la aplicación.
'
'  3) En cada formulario, añadir después del Option Explicit la siguiente línea:
'     Public Ajustado as Boolean
'
'     El objetivo de esta variable es que, por error, se pueda ejecutar más de una vez a la rutina,
'     lo cual produciría un doble redimensionamiento de los controles.
'
'  4) Si hay algún SSTab, hay que llamar a la rutina en el Form_Activate.
'
'     Ajustar Me
'
'     Si se quiere centrar el form, llamar directamente a
'
'     Centrar Me
'
'     la cual llama a la rutina del módulo que se encarga de redimensionar los controles.
'
'
'
'     Y eso es todo. Dudas, comentarios, sugerencias, etc, a la dirección
'
'     oscarsen@cros.es
'
'
'
'

Public Sub Ajustar(F As Form)
FactorX = Pantalla.Ancho * Screen.TwipsPerPixelX / Screen.Width
FactorY = Pantalla.Alto * Screen.TwipsPerPixelY / Screen.Height
If (FactorX = 1 And FactorY = 1) Or F.Ajustado Then Exit Sub
F.Visible = False
Dim C As Object
If F.WindowState = vbNormal Then
   AjusteNormal F
End If
For Each C In F.Controls
   Select Case LCase(TypeName(C.Container))
   Case LCase(F.Name)
      Select Case LCase(TypeName(C))
      Case "label"
         AjusteNormal C
         C.AutoSize = C.AutoSize
      Case "line"
         C.X1 = C.X1 / FactorX
         C.X2 = C.X2 / FactorX
         C.Y1 = C.Y1 / FactorY
         C.Y2 = C.Y2 / FactorY
      Case "picturebox"
         AjusteNormal C
         'C.Align = C.Align
      Case "shape"
         AjusteNormal C
      'No se ha detectado nada
      Case "textbox"
         AjusteNormal C
      'No se ha detectado nada excepto la escalabilidad de la fuente
      Case Else
         'Shape
         AjusteNormal C
      End Select
   Case "sstab"
      Dim T As Integer
      T = C.Container.Tab
      C.Container.Tab = 0
      Do
         If Left$(Str(C.Left), 1) = "-" Then
            C.Container.Tab = C.Container.Tab + 1
         Else
            Exit Do
         End If
      Loop
      AjusteNormal C
      C.Container.Tab = T
   Case Else
      AjusteNormal C
   End Select
Next
F.Ajustado = True
F.Visible = True
End Sub

Private Sub AjusteNormal(C2 As Object)
On Error Resume Next
C2.Font.Size = C2.FontSize / FactorX
C2.Height = C2.Height / FactorY
C2.Width = C2.Width / FactorX
C2.Left = C2.Left / FactorX
C2.Top = C2.Top / FactorY
End Sub

Public Sub Centrar(Optional F As Form)
If IsMissing(F) Then Set F = Screen.ActiveForm
Ajustar F
F.Move (Screen.Width - F.Width) / 2, (Screen.Height - F.Height) / 2
End Sub

Public Sub VolcarAImpresora(Fich As String, Optional Puerto As String)
On Error GoTo SalirImpri
If IsMissing(Puerto) Or Len(Puerto) = 0 Then Puerto = "LPT1"
'Dim L As Long
'If FileLen(Fich) < 10000 Then
'L = FileLen(Fich)
'Else
'L = 100
'End If
Dim Cadena As String * 100, X As Integer, Y%
X = FreeFile
Open Fich For Binary Access Read As #X
Y = FreeFile
Open Puerto For Binary Access Write As #Y
Do While Not EOF(X)
Get #X, , Cadena
DoEvents
Put #Y, , Cadena
DoEvents
Loop
SalirImpri:
Close #X, #Y
If Err.Number = 0 Then Exit Sub
MsgBox "Ha ocurrido un error con el fichero " & vbCrLf & Fich
End Sub
