Attribute VB_Name = "Module4"
 Sub �������()
    Dim i As Long
    Dim lastRow As Long
    
    ' ���������� ��������� ������ � �������
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' ��������� ������ ����� �����
    For i = lastRow To 2 Step -1
        ' ���� � ������ A ����� ��� ������, �� �������� ������ �� A �� G
        If Cells(i, "A").Value = "��������" Or Cells(i, "A").Value = "�����" Or Cells(i, "A").Value = "��������" Or Cells(i, "A").Value = "��������" Then
            Range(Cells(i, "A"), Cells(i, "G")).ClearContents
        End If
    Next i
    MsgBox ("������� ��������� �������")
End Sub

