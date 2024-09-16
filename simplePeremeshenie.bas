Sub RearrangeColumnsACBD()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Укажи свой лист, если название другое
    
    ' Переместим третий столбец (C) на второе место
    ws.Columns("C:C").Cut
    ws.Columns("B:B").Insert Shift:=xlToRight
End Sub
