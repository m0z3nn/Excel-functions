Attribute VB_Name = "Module1"
Function РасчётПроцОтЧека(Percent As String, Count As Integer, Sum As Double) As Double
    If Percent = "*" Then
        РасчётПроцОтЧека = 1000 * Count
    Else
        РасчётПроцОтЧека = (Count * Sum) * Percent
    End If
End Function
   
   
