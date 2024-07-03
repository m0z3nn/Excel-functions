Attribute VB_Name = "Module2"
Sub ОбновлениеДаты()
    Range("M1").Value = Date
    Range("M1").NumberFormat = "dd.mm.yyyy"
    MsgBox ("Дата успешно изменена")
End Sub
