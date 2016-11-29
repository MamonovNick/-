Public Class AppSettings
    Public ConnStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:/Монеты-Access/Монеты.mdb"
    Public PicPath As String = "D:\Монеты-Access\Фотографии монет\"
    Public FileDBPath As String = "D:\Монеты-Access\Монеты.mdb"
    Public LastVisit As Date
    Public Name As String
    Public Phone As String
End Class

Public Class MainSettings
    Public Shared AppS As New AppSettings
End Class