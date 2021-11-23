Sub LaskeSaldo()
    Dim x As Integer
    Dim a As Integer
    Dim b As Integer
    NumRows = Range("E1:E2145").Rows.Count

'    This is a simple VBA script subroutine to perform multiple cell counts.
'    Replace values in "Range" function call with your own cell range 
'    Replace the value of x variable in the for loop to a value where your first cell values appear that you want to sum or subtract.
'    Sum of values is performed depending on the condition. Then the result is passed into a root cell. In this example the root cell is "G", replace it with your own cell.
    
    For x = 6 To NumRows
            
            If Range("E" + CStr(x)).Value < Range("F" + CStr(x)).Value Then
                a = Range("E" + CStr(x)).Value
                b = Range("F" + CStr(x)).Value
                Range("G" + CStr(x)).Value = (b - a)
           
            
            ElseIf Range("E" + CStr(x)).Value > Range("F" + CStr(x)).Value Then
                a = Range("E" + CStr(x)).Value
                b = Range("F" + CStr(x)).Value
                Range("G" + CStr(x)).Value = Excel.WorksheetFunction.Sum(a, -b) * -1
                
                
                
            End If
        
        ActiveCell.Offset(1, 0).Select
    Next x

End Sub