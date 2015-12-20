Sub VisualFileCopy (SourceFileName As String, TargetFileName As String, ProgressGuage As Control)

   'This routine will copy any size and type of file giving
   'a visual progress indication to the user.  Simply pass
   'the Source Filename, Target Filename, and name of the
   'control to use as a progress guage.  The code below
   'uses a standard Panel3D1 control from THREED.VBX but
   'any control that gives the desired effect may be used.
   'The progress range is 1 to 100 but can be any range.



   Dim I As Integer
   Dim SourceFileNo As Integer
   Dim TargetFileNo As Integer
   Dim SourceFileSize As Long
   Dim CopyBuffer As String

   On Error GoTo FileCopyErrorRoutine
   SourceFileSize = FileLen(SourceFileName)
   CopyBuffer = Space$(25000)             'AS LARGE AS POSSIBLE UNDER 65,000

'--KILL THE CURRENT TARGET FILE IF IT EXISTS
   If Len(Dir$(TargetFileName)) Then
      Kill TargetFileName
   End If

'--OPEN FILES
   SourceFileNo = FreeFile
   Open SourceFileName For Binary Access Read As SourceFileNo
   TargetFileNo = FreeFile
   Open TargetFileName For Binary Access Write As TargetFileNo

'--COPY SOURCE FILE TO TARGET FILE
   For I = 1 To SourceFileSize \ Len(CopyBuffer)
      Get #SourceFileNo, , CopyBuffer
      ProgressGuage.FloodPercent = I * Len(CopyBuffer) / SourceFileSize * 100  'UPDATE
PROGRESS GUAGE
      Put #TargetFileNo, , CopyBuffer
      DoEvents
   Next I

'--COPY ANY ODD PORTION OF THE SOURCE FILE REMAINING
   CopyBuffer = Space$(SourceFileSize - Loc(TargetFileNo))
   If Len(CopyBuffer) Then
      Get #SourceFileNo, , CopyBuffer
      Put #TargetFileNo, , CopyBuffer
   End If
   Close SourceFileNo
   Close TargetFileNo

Exit Sub

FileCopyErrorRoutine:
   MsgBox Error$
   Exit Sub
End Sub
