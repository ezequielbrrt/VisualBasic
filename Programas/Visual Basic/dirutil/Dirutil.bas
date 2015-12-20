Option Explicit

Const ATTR_DIRECTORY = 16

Function DirUtil (DirOrigen As String, NivelIndent As Integer, TotalFicheros As Long, TotalTamaño As Long) As Integer

Dim DirOK As Integer, i As Integer
Dim DirReturn As String
ReDim d(100) As String
Dim NumDir As Integer
Dim FicheroActual As String
Dim DirActual As String
Dim NumFicheros As Long, Tamaño As Long


'si ejecutas esto con VB 3 en W95 hay problemas con algunos nombres
'de fichero que tengan acentos y otros caracteres no permitidos en
'ms-dos 6.2 y anteriores. Para evitarlos es necesario ignorar los errores
On Error Resume Next

   frmMain!lblBusqueda = "Buscando en " & DirOrigen

   DirActual$ = CurDir$
   
   DirOrigen = UCase$(DirOrigen)

   ' Inicializar las vbles
   DirReturn = Dir(DirOrigen & "\*.*", ATTR_DIRECTORY)
   
   ' Buscar todos los subdirectorios
   Do While DirReturn <> ""
      ' No tratar los directorios  "." y ".."
      If DirReturn <> "." And DirReturn <> ".." Then
         If GetAttr(DirOrigen & "\" & DirReturn) = ATTR_DIRECTORY Then
         ' Si es un directorio añadirlo a la lista
            NumDir = NumDir + 1
            d(NumDir) = DirOrigen & "\" & DirReturn
         End If
      End If
      DirReturn = Dir
      DoEvents
   Loop
   
   ' Coger ahora todos los ficheros que no sean directorios
   DirReturn = Dir(DirOrigen & "\*.*", 0)

   DoEvents

   Do While DirReturn <> ""
      ' Comprobar que no sean directorios
      If Not (GetAttr(DirOrigen & "\" & DirReturn) = ATTR_DIRECTORY) Then
         ' Es un fichero
         NumFicheros = NumFicheros + 1
         Tamaño = Tamaño + FileLen(DirOrigen & "\" & DirReturn)
      End If
      DirReturn = Dir
      DoEvents
   Loop

   ' Añadir la información al outline
   ' Buscar la ultima "\"
   For i% = Len(DirOrigen) To 1 Step -1
      If Mid$(DirOrigen, i%, 1) = "\" Then Exit For
   Next
   DirOrigen = Right$(DirOrigen, Len(DirOrigen) - i%)

   frmMain!outDirInfo.AddItem DirOrigen & ":" & Chr$(9) & Format$(Tamaño / 1024, "###,###,##0") & " Kb"
   frmMain!outDirInfo.PictureType(frmMain!outDirInfo.ListCount - 1) = 0
   frmMain!outDirInfo.Indent(frmMain!outDirInfo.ListCount - 1) = NivelIndent
   DoEvents

   ' Actualizar la vbles. de retorno
   TotalFicheros = TotalFicheros + NumFicheros
   TotalTamaño = TotalTamaño + Val(Format$(Tamaño / 1024, "000000000"))

   ' Recorrer todos los subdirectorios que encontramos antes
   For i = 1 To NumDir
      DirOK = DirUtil(d(i), NivelIndent + 1, TotalFicheros, TotalTamaño)
   Next

   DoEvents
   DirUtil = True

ExitFunc:

   ChDir DirActual$

   Exit Function

DirErr:

   MsgBox "Error: " & Error$(Err)
   
   DirUtil = False
   Resume ExitFunc
   
End Function

