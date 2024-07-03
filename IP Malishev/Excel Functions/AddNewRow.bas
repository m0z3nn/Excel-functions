Attribute VB_Name = "Module3"
Sub ДобавитьНовуюСтроку()
    Dim newRow As Long
    Dim currentRow As Long
    
    ' Определить текущую строку
    currentRow = ActiveCell.Row
    
    ' Вставить новую строку перед текущей
    Rows(currentRow).Insert
    
    ' Копировать форматы из строки выше
    Rows(currentRow - 1).Copy
    Rows(currentRow).PasteSpecial xlPasteFormats
    
    ' Копировать выпадающие списки из строки выше
    Rows(currentRow - 1).Columns("A").Copy
    Rows(currentRow).Columns("A").PasteSpecial xlPasteValidation
    
    Rows(currentRow - 1).Columns("E").Copy
    Rows(currentRow).Columns("E").PasteSpecial xlPasteValidation
    
    ' Очистить значения в новой строке
    Rows(currentRow).ClearContents
    
     ' Установить функцию в столбце H
    Cells(currentRow, "H").Formula = "=РасчётПроцОтЧека(D" & currentRow & ",F" & currentRow & ",G" & currentRow & ")"
    
    ' Перейти к новой строке
    Application.Goto Cells(currentRow, 1)
    
    ' Очистить буфер обмена
    Application.CutCopyMode = False
    MsgBox ("Строка создана!")
    End Sub
