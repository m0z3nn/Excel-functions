Attribute VB_Name = "Module4"
 Sub Очистка()
    Dim i As Long
    Dim lastRow As Long
    
    ' Определить последнюю строку с данными
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Перебрать строки снизу вверх
    For i = lastRow To 2 Step -1
        ' Если в ячейке A стоит тип оплаты, то очистить строку от A до G
        If Cells(i, "A").Value = "Наличные" Or Cells(i, "A").Value = "Карта" Or Cells(i, "A").Value = "Терминал" Or Cells(i, "A").Value = "Оплачено" Then
            Range(Cells(i, "A"), Cells(i, "G")).ClearContents
        End If
    Next i
    MsgBox ("Очистка выполнена успешно")
End Sub

