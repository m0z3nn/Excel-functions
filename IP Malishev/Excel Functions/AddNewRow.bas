Attribute VB_Name = "Module3"
Sub �������������������()
    Dim newRow As Long
    Dim currentRow As Long
    
    ' ���������� ������� ������
    currentRow = ActiveCell.Row
    
    ' �������� ����� ������ ����� �������
    Rows(currentRow).Insert
    
    ' ���������� ������� �� ������ ����
    Rows(currentRow - 1).Copy
    Rows(currentRow).PasteSpecial xlPasteFormats
    
    ' ���������� ���������� ������ �� ������ ����
    Rows(currentRow - 1).Columns("A").Copy
    Rows(currentRow).Columns("A").PasteSpecial xlPasteValidation
    
    Rows(currentRow - 1).Columns("E").Copy
    Rows(currentRow).Columns("E").PasteSpecial xlPasteValidation
    
    ' �������� �������� � ����� ������
    Rows(currentRow).ClearContents
    
     ' ���������� ������� � ������� H
    Cells(currentRow, "H").Formula = "=����������������(D" & currentRow & ",F" & currentRow & ",G" & currentRow & ")"
    
    ' ������� � ����� ������
    Application.Goto Cells(currentRow, 1)
    
    ' �������� ����� ������
    Application.CutCopyMode = False
    MsgBox ("������ �������!")
    End Sub
