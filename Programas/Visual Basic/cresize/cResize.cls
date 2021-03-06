VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Ajus As Boolean
Private FactorX As Single, FactorY As Single

'Clase escrita por Oscar Manuel G�mez Senovilla
'Clase cResize.cls escrita en Visual Basic 5, siendo el codigo
'en principio v�lido para cualquier versi�n de VB
'El objetivo de este m�dulo es no tener que preocuparse por la resoluci�n en la cual
'ha sido dise�ado un formulario, as� como la si la resoluci�n de la pantalla de
'otro ordenador donde se vaya a ejecutar es distinta a la
'de la la pantalla en que fue dise�ado
'
'Consejos y recomendaciones:
'   - Usar fuentes escalables, como la Arial, en vez de
'     la fuente por defecto MS Sans Serif. Si existe
'     la seguridad de que la resoluci�n de origen y
'     destino es la misma, no habr� problemas, ya que la rutina,
'     para ganar tiempo, comprueba si coinciden las resoluciones,
'     en cuyo caso simplemente no har� ning�n reajuste.
'     Si hay un reajuste y se usa la fuente Sans Serif,
'     se producir�n resultados indeseados, y que se deben
'     a que dicha fuente no soporta la escalabilidad necesaria.
'
'   - Tambi�n tiene un m�todo Centrar
'     para centrar el formulario en la pantalla
'
'Para usarla:
'
'  1) A�adir esta clase al proyecto
'
'  2) Crear una variable tipo cResize a nivel de formulario
'
'     Dim Rsz as New cResize
'
'  3) Llamar al m�todo Ajustar con el siguiente formato:
'
'     Rsz.Ajustar [Form [, AnchoX]]
'
'     Donde Form es el formulario activo (se recomienda usar Me),
'     y AnchoX es el ancho en pixels de la pantalla en modo dise�o.
'     Si se omite AnchoX, toma por defecto 800 pixels, por lo
'     que si el formulario ha sido dise�ado en otra resoluci�n
'     es necesario incluir el par�metro en la llamada
'     para no obtener resultados err�neos.
'
'     Si se llama a la rutina en el Form_Load, el par�metro
'     del form NO ES OPCIONAL, en cuyo caso (al menos en VB5)
'     generar� el error 91.
'
'     Ejemplo:
'
'     Private Sub Form_Load()
'     Rsz.Ajustar Me
'     ...
'
'     End Sub
'     Si se quiere centrar el form, llamar directamente a
'
'     Rsz.Centrar Me
'
'     la cual llama a su vez al m�todo Ajustar
'
'     Para que el par�metro Form sea opcional, es necesario que
'     Screen.ActiveForm ya est� establecido. Para ello, se recomienda
'     llamar al m�todo que se desee en el evento Activate del Form.
'
'     AVISO: Si hay un SSTab, hay que llamar al m�todo en
'     el evento Activate, ya que si no, genera un error.
'     Y eso es todo. Dudas, comentarios, sugerencias, etc, a la direcci�n
'
'     oscarsen@cros.es

Public Property Get Ajustado() As Boolean
Ajustado = Ajus
End Property

Public Sub Ajustar(Optional F As Object, Optional XRes)
Attribute Ajustar.VB_Description = "M�todo para ajustar los controles del formulario"
Attribute Ajustar.VB_UserMemId = 0
Dim YRes As Integer
If IsMissing(XRes) Then XRes = 800
If XRes = 800 Then
   YRes = 600
ElseIf XRes = 640 Then
   YRes = 480
ElseIf XRes = 1024 Then
   YRes = 768
ElseIf XRes = 1280 Then
   YRes = 1024
Else
   XRes = 800
   YRes = 600
End If
FactorX = XRes * Screen.TwipsPerPixelX / Screen.Width
FactorY = YRes * Screen.TwipsPerPixelY / Screen.Height
If (FactorX = 1 And FactorY = 1) Or Ajus Then
   Ajus = True
   Exit Sub
End If
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
      Loop 'While C.Container.Tab <= C.Container.Tabs
      AjusteNormal C
      C.Container.Tab = T
   Case Else
      AjusteNormal C
   End Select
Next
Ajus = True
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

Public Sub Centrar(Optional F As Form, Optional XRes)
If IsMissing(XRes) Then XRes = 800
If TypeName(F) = "Nothing" Then Set F = Screen.ActiveForm
Ajustar F, XRes
F.Move (Screen.Width - F.Width) / 2, _
   (Screen.Height - F.Height) / 2
End Sub
