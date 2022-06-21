'Removes a single large span of spaces. Useful to undo format where the original author used spaces instead of new lines to format the output of formulas. 
'Using spaces to format leads to odd formatting as the sheet is zoomed in or out.


Sub concat()
Dim match As Integer
Dim nomore As Boolean
Dim length As Integer
Dim space_length As Integer
Dim min_length As Integer
Dim x As Integer
For Each cell In Selection

nomore = False
length = Len(cell.Formula)
min_length = 2

For space_length = 1 To length 'increment all possible length of spaces.  String needs to be at least a character, otherwise no need to run.

x = Len(cell.Formula) + min_length - space_length 'invert increment counting from longest string of spaces to shortest.


'Generally we only want to remove one block of spaces
'Instead of checking for every length of spaces and calling replace each time, find the first match, which will be the biggest
On Error GoTo catch
match = InStr(1, cell.Formula, space(x))
If nomore = True Then
GoTo catch
End If

If nomore = False Then

If match > 0 Then


    nomore = True
    'Once match is true jump to replace and remove the space
    cell.Formula = Replace(cell.Formula, space(x), "" + Chr(10), 1, 1, vbTextCompare)

    End If
    
End If

Next space_length

catch:
Next cell
nomore = False
