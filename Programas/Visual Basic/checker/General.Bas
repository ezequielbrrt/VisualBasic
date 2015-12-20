Attribute VB_Name = "General"
'Systray
Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uID As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

Global Const NIM_ADD = &H0& 'constants & flags for NotifyIcons
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200
Global NI As NOTIFYICONDATA

'File Dialog Open
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
 
Type OPENFILENAME
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

'Constantes y flags para OpenFile Dialog
Global Const OFNFileMustExist = &H1000
Global Const OFNHideReadOnly = &H4
Global Const OFNPathMustExist = &H800
Global Const OFNHelpButton = &H10
Global Const OFNExplorer = &H80000

'Constantes para la comunicacion
Global Const POP3Port = 110

'Opciones de configuracion del programa
Global pop3Host As String
Global pop3User As String
Global pop3Passwd As String
Global Interval As String
Global EmailProgram As String
Global Arguments As String
Global Timeout As String

'Para poder hacer e-mail desde el programa
Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal _
lpDirectory As String, ByVal nShowCmd As Long) As Long

'Para poner el form "always on top"
Declare Function SetWindowPos& Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long)

'Para poder esperar un período de tiempo
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

