'se supone que esto sirve para imprimir textbox multilinea
Sub parseandprint (printstring As String, Lmarg As Integer, Rmarg As Integer)
Lmarg = (printer.Width / 80) * Lmarg
Rmarg = (printer.Width / 80) * Rmarg


nl = Chr$(13) & Chr$(10)
'To keep track of where we are and last were when searching through a string
Dim currentpos As Integer
Dim lastpos As Integer
'Storage for the component substrings of a string with embedded CRs
Dim substring() As String
'Storage for current printable string
Dim tempstring As String
'The ubiqitous loop counter
Dim counter As Integer
'The maximum printable line width; based on page width less the L and R margins
Dim linelength As Single

linelength = printer.Width - Lmarg - Rmarg

'First, if there are no CRs and the line is short enough, print it
If InStr(printstring, nl) = 0 And printer.TextWidth(printstring) <= linelength Then


  'Print it and skip the rest of the checks
  printer.CurrentX = Lmarg
  printer.Print printstring
  Exit Sub
End If

'Now, The string has CR and/or is too long. First, check for CRs and if they exist,
'split the text into substrings at the CRs.
If InStr(printstring, nl) <> 0 Then
  'Set the array to one element and assign the first substring to that element
  ReDim substring(1)
  currentpos = InStr(printstring, nl)
  substring(1) = Left$(printstring, currentpos - 1)
  'Now work through the remainder of the string and assign the remaining substrings to the array
  Do
  lastpos = currentpos + 2
  currentpos = InStr(currentpos + 2, printstring, nl)
  'If no more CRs found, we have the last substring to extract; then we must exit the loop
  If currentpos = 0 Then


    ReDim Preserve substring(UBound(substring) + 1)
    substring(UBound(substring)) = Right$(printstring, Len(printstring) - lastpos + 1)
    Exit Do
  End If
  ReDim Preserve substring(UBound(substring) + 1)
  substring(UBound(substring)) = Mid$(printstring, lastpos, currentpos - lastpos)
  Loop
Else
'There are no CRs, but the line is too long. Put it in substring(1) and process it to length
  ReDim substring(1)
  substring(1) = printstring
End If

'The string is now broken down into substrings without CRs. Work through the substrings and break
'them down into printable size bits.
For counter = 1 To UBound(substring)
  If printer.TextWidth(substring(counter)) <= linelength Then


    'Print it - it fits
    printer.CurrentX = Lmarg
    printer.Print substring(counter)
  Else
  'Break it into page-width chunks and print them
    Do
      currentpos = 1
      lastpos = 1

      Do
      'Now build a temporary string and test its textwidth until it is too big. Then back up one
      'and print it. We will build the string by searching for spaces and adding the text previous
      'to the space to the temp string (instant word wrap!)
      'For second or subsequent pass, if it fits, print it and exit the loop to test next element
        If printer.TextWidth(substring(counter)) <= linelength Then
          printer.CurrentX = Lmarg
          printer.Print substring(counter)


          'Debug.Print substring(counter)
          Exit Do
        End If
        lastpos = currentpos
        currentpos = InStr(currentpos + 1, substring(counter), " ")
        If currentpos = 0 Then
        'We have reached the end of the string. If the substring is still longer than a printable
        'length, grab the front part up to the last found space and print it.
          If printer.TextWidth(substring(counter)) > linelength Then
            tempstring = Left$(substring(counter), lastpos - 1)
            printer.CurrentX = Lmarg
            printer.Print tempstring
            substring(counter) = Right$(substring(counter), Len(substring(counter)) - lastpos)
            'We need to empty tempstring here so we know how to handle a zero value in currentpos
            'when we exiit this loop.
            tempstring = ""
            Exit Do


          Else
            'It is a printable size, so print it and exit the loop
            printer.CurrentX = Lmarg
            printer.Print substring(counter)
            tempstring = ""
            Exit Do
          End If
        End If
        tempstring = Left$(substring(counter), currentpos - 1)
      Loop Until printer.TextWidth(tempstring) >= linelength
      'We may now have a zero value in currentpos. If so, tempstring will have a zero length (I hope)
      'Remove that string from the substring and retest
      If Len(tempstring) <> 0 Then
        tempstring = Left$(substring(counter), lastpos - 1)
        printer.CurrentX = Lmarg
        printer.Print tempstring
        substring(counter) = Right$(substring(counter), Len(substring(counter)) - Len(tempstring) - 1)
      End If


      'Now there are three possibilities
      'It is empty; exit and check next element
      If printer.TextWidth(substring(counter)) = 0 Then Exit Do
      'It is less than or = to the line length, so print it, and exit to check Next element
      If printer.TextWidth(substring(counter)) <= linelength Then
        printer.CurrentX = Lmarg
        printer.Print substring(counter)
        Exit Do
      End If
      'Otherwise it is greater than line length, so loop and extract the next printable bit
    Loop
  End If
Next counter
End Sub
