# HighlightSelectedRowColumn_VBA
While working, highlight row and column of selected cell

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'**********************************General comments about this procedure****************************************

'1 - This procedure should be copied in worksheet, not in module.
'2 - The excel application event is being used to run the procedure when a new cell selected.
'3 - This procedure will only work for sheet that has its data distributed up to 15000 rows/columns or less _
     because the number of elements in each array is 15000 for now.
'4 - This procedure will keep all the original interior colorindex of cells.
'5 - To stop this procedure, select a cell outside of used range or delete this entire procedure.

'****************************************************************************************************************
    'Declare all the variables here.
    Dim wSheet As Worksheet

    Static maxRows As Long
    Static maxColumns As Long
    Dim cellColorIndexInRow As Long
    Dim c As Long
    Dim r As Long
    Static previousScanRow As Long
    Static previousScanColumn As Long

    Static ArrayColorIndexInRow(1 To 15000) As Long
    Static ArrayColorIndexInColumn(1 To 15000) As Long

'*****************************************************************************************************************
    'Set all the objects here.
    Set wSheet = ThisWorkbook.ActiveSheet

'*****************************************************************************************************************
    'If more than one cell is selected, the procedure will stop.
    If Target.Cells.Count > 1 Then Exit Sub

    Application.ScreenUpdating = False

'*****************************************************************************************************************
    'This chunk of code assigns all the previous colors back to the row and column of previously selected cell.
    If previousScanRow <> 0 Then
        For c = 1 To maxColumns

            With wSheet.Cells(previousScanRow, c).Interior
                If ArrayColorIndexInRow(c) <> 0 Then
                    .ColorIndex = ArrayColorIndexInRow(c)
                End If
            End With

        Next c
    End If

    If previousScanColumn <> 0 Then
        For r = 1 To maxRows

            With wSheet.Cells(r, previousScanColumn).Interior
                If ArrayColorIndexInColumn(r) <> 0 And r <> previousScanRow Then
                    .ColorIndex = ArrayColorIndexInColumn(r)
                End If
            End With

        Next r
    End If

'*******************************************************************************************************************
    'This code determines the maximum number of rows and columns to use.
    'This section is also used for some general conditioning code.
    
    With wSheet.UsedRange
        maxRows = .Rows.Count
        maxColumns = .Columns.Count
    End With

    If maxRows = 0 And maxColumns = 0 Then
        MsgBox "This worksheet is empty."
        Exit Sub
    End If

    If Target.Column > maxColumns Or Target.Row > maxRows Then
        MsgBox "Selected cell is out of used range."
        Exit Sub
    End If

'******************************************************************************************************************
    'This code takes all the original colorIndex values of selected row/column and assigns to arrays for storage.
    'Then predefined colorIndex values are assigned to entire selected row and column.
    For c = 1 To maxColumns

        With wSheet.Cells(Target.Row, c).Interior
            If .ColorIndex <> xlColorIndexNone Then
                ArrayColorIndexInRow(c) = .ColorIndex
                .ColorIndex = 42
            Else
                ArrayColorIndexInRow(c) = xlColorIndexNone
                .ColorIndex = 8
            End If
        End With

    Next c

    For r = 1 To maxRows

        With wSheet.Cells(r, Target.Column).Interior
            If .ColorIndex <> xlColorIndexNone Then
                If r <> Target.Row Then
                    ArrayColorIndexInColumn(r) = .ColorIndex
                    .ColorIndex = 42
                End If
            Else
                If r <> Target.Row Then
                    ArrayColorIndexInColumn(r) = xlColorIndexNone
                    .ColorIndex = 8
                End If
            End If
        End With

    Next r
    
'*******************************************************************************************************************
    'The indexes of row and column of selected cell are assigned to variables to be used in next scan.
    previousScanRow = Target.Row
    previousScanColumn = Target.Column

    Application.ScreenUpdating = True

End Sub
