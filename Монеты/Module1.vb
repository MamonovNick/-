Module Module1
    'настройка папок и других прараметров перед запуском программы
    Sub Start_Setup()
        If Not IO.Directory.Exists(Application.StartupPath + "\TempReports") Then
            IO.Directory.CreateDirectory(Application.StartupPath + "\TempReports")
        End If
    End Sub

    Sub Del_Tmp()
        If IO.Directory.Exists(Application.StartupPath + "\TempReports") Then
            IO.Directory.Delete(Application.StartupPath + "\TempReports", True)
        End If
    End Sub
End Module
