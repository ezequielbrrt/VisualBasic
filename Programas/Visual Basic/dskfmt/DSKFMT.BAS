Attribute VB_Name = "Module1"
=======
I am unsure how to do this in vb4/win95 however this is the Correct way to do it in VB3.0/Win3.1
Below is code on how to format a floppy disk from Visual Basic for Windows. This code will not work with Windows NT or Windows 95.
Paste Code into Module and from Form that you want to initialize Format add....

HTH
Curt

Example:
Call FMFormat(form1)


Type Rect
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type


Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As Rect)
Declare Function IsWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function WinExec Lib "Kernel" (ByVal lpCmdLine As String, ByVal nCmdShow As Integer) As Integer
Declare Function SetActiveWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetActiveWindow Lib "User" () As Integer
Declare Function LockWindowUpdate Lib "User" (ByVal hwndLock As Integer) As Integer
Declare Function GetDesktopWindow Lib "User" () As Integer
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Integer
Const WM_COMMAND = &H111
Const WM_CLOSE = &H10
Dim wFlag%
Dim lpDlgRect As Rect
Dim lpDskRect As Rect
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4

Sub CenterDialog(WinText As String)
   Do
    If FindWindow(0&, WinText) Then Exit Do
    x% = DoEvents()
   Loop

 wnd% = GetActiveWindow()
 Call GetWindowRect(wnd%, lpDlgRect)
 wdth% = lpDlgRect.Right - lpDlgRect.Left
 hght% = lpDlgRect.Bottom - lpDlgRect.Top
 Call GetWindowRect(GetDesktopWindow(), lpDskRect)
 Scrwdth% = lpDskRect.Right - lpDskRect.Left
 Scrhght% = lpDskRect.Bottom - lpDskRect.Top
 x% = (Scrwdth% - wdth%) / 2
 Y% = (Scrhght% - hght%) / 2
 Call SetWindowPos(wnd%, 0, x%, Y%, 0, 0, SWP_NOZORDER Or SWP_NOSIZE)
End Sub

Sub FMFormat(F As Form)

    FMhWnd = FindWindow("WFS_Frame", 0&)

    If FMhWnd = 0 Then
    I% = WinExec("Winfile", 0)
    FMhWnd = FindWindow("WFS_Frame", 0&)
       If FMhWnd = 0 Then
           MsgBox "FileMan ain't home"
           Exit Sub
       End If
    wFlag = 1
    End If

    I% = LockWindowUpdate(GetDesktopWindow())

    I% = PostMessage(FMhWnd, WM_COMMAND, &HCB, 0)

    Call CenterDialog("Format Disk")

    I% = LockWindowUpdate(0)

    wnd% = GetActiveWindow()

    While IsWindow(wnd%)
    x = DoEvents()
    Wend

    x = DoEvents()

    If wFlag Then
    wFlag = 0
    I% = PostMessage(FMhWnd, WM_CLOSE, 0, 0)
    End If

    I% = SetActiveWindow(F.hWnd)

End Sub








royfran@statcan.ca (Francois Roy) wrote:
>Bjorn Larsen (blarsen@login.eunet.no) wrote:
>: 1. Anyone knows how to format a diskette from within VB4.0 windows 95 ??
>
>: 2. Anyone knows how to change label on a diskette or drive ?
>
>Dunno about Win95, but with VB3/Win3.1, the only solution I could find
>was to use Shell to call the Dos command.

