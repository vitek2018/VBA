Attribute VB_Name = "Data Range Editor"

Public Sub Main()

' This VBA program copyright Vitek2018
' This program takes a range of cells contained in two columns and checks to see if they are within a predetermined allowable range.
' If the cell value is outside of the allowable range, the appropriate subroutine is called.

 Dim CellRange As Range
   
 Worksheets("pHData").Activate
  
 Set CellRange = Range("B23:C2266")
 
 For Each cell In CellRange
 
    If (cell.Value <= 6.66) Then
      Call IncreaseCellValue(cell)
    ElseIf (cell.Value >= 13.00) Then
      Call DecreaseCellValue(cell)
    End If
 
 Next cell
 
 MsgBox "Done"
 
End Sub

Private Sub IncreaseCellValue(cell)

' This subroutine keeps increasing the cell value until it is within allowable range

Do
 cell.Value = cell.Value + 1
 Loop While cell.Value <= 6.66

End Sub

Private Sub DecreaseCellValue(cell)

' This subroutine keeps decreasing the cell value until it is within allowable range

Do
 cell.Value = cell.Value - 1
Loop While cell.Value >= 13.00

End Sub
