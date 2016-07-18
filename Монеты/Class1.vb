Public Class Class1
    ' Класс для передачи значений между формами
    Private Shared CatNum As String = ""

    Public Shared Function GetCat()
        Return CatNum
    End Function

    Public Shared Sub PutCat(e As String)
        CatNum = e
    End Sub
End Class
