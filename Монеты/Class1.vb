Public Class Class1
    ' Класс для передачи значений между формами
    Private Shared CatNum As String = ""
    Private Shared dateSince As Date
    Private Shared dateTo As Date
    Private Shared FileToOpenInWebBro As String = ""

    Public Shared Function GetCat()
        Return CatNum
    End Function

    Public Shared Sub PutCat(e As String)
        CatNum = e
    End Sub

    Public Shared Function getDate(type As Int16)
        Select Case type
            Case 1
                Return dateSince
            Case Else
                Return dateTo
        End Select
    End Function

    Public Shared Sub setDate(dates As Date, datet As Date)
        dateSince = dates
        dateTo = datet
    End Sub

    Public Shared Function getFileWeb()
        Return FileToOpenInWebBro
    End Function

    Public Shared Sub setFileWeb(str As String)
        FileToOpenInWebBro = str
    End Sub

End Class
