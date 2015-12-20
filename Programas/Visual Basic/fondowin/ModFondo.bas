Attribute VB_Name = "ModFondo"
Option Explicit

Declare Function WriteProfileString& Lib "kernel32" _
   Alias "WriteProfileStringA" (ByVal lpszSection As String, _
   ByVal lpszKeyName As String, ByVal lpszString As String)
Declare Function GetProfileString& Lib "kernel32" _
   Alias "GetProfileStringA" (ByVal lpAppName As String, _
   ByVal lpKeyName As String, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Long)
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetLastError& Lib "kernel32" ()
Declare Function PostMessageByString Lib "user32" _
   Alias "PostMessageA" (ByVal hwnd As Long, _
   ByVal wMsg As Long, ByVal wParam As Long, _
   ByVal lParam As String) As Long
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_WININICHANGE = &H1A
Public Const WM_SETTINGCHANGE = &H1A&

#If Win32 Then
Declare Function GetActiveWindow& Lib "user32" ()
Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, _
    ByVal lpvParam As String, ByVal fuWinIni As Long) As Long

Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDESKPATTERN = 21
Public Const SPIF_SENDWININICHANGE = &H2
Public N As Long


#Else
Declare Function GetActiveWindow% Lib "user" ()
Declare Function SystemParametersInfo Lib "user" _
    (ByVal uAction As Integer, ByVal uParam As Integer, _
    ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer

Public Const SPIF_UPDATEINIFILE = 1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDESKPATTERN = 21
Public Const SPIF_SENDWININICHANGE = &H2
Public N As Integer

Declare Function GetWindowsDirectory% Lib _
    "kernel" (ByVal lpBuffer As String, ByVal nSize As Integer)
#End If 'WIN32


Public Sub CDir(WP As String)
Dim I As Integer
FrmFondo!Drive1.Drive = Left$(WP, 2)
I = Len(WP)
Do While Not Mid$(WP, I, 1) = "\"
I = I - 1
Loop
FrmFondo!Dir1.Path = Left$(WP, I - 1)
WP = Mid$(WP, I + 1)
With FrmFondo!File1
For I = 1 To .ListCount
If UCase(.List(I)) = UCase(WP) Then
   .ListIndex = I
   Exit For
End If
Next
End With
End Sub
