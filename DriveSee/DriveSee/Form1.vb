Public Class Form1
    Dim FSO, Drive, Drive1
    Dim DriveAr(100)
    Dim i, j As Integer
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FSO = CreateObject("Scripting.FileSystemObject")
        For Each Drive In FSO.Drives
            ListBox1.Items.Add(Drive.DriveLetter.ToString)
            TreeView1.Nodes.Add(Drive.DriveLetter.ToString)
        Next

        On Error Resume Next
        For i = 0 To TreeView1.Nodes.Count - 1
            Drive = FSO.GetDrive(ListBox1.Items.Item(i))
            If Drive.IsReady <> False Then
                TreeView1.Nodes(i).Nodes.Add("Параметры диска")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Полный объём диска: " & Drive.TotalSize \ 1000000 & " Мб.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("Занятое место: " & (Drive.TotalSize - Drive.FreeSpace) \ 1000000 & " Мб.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("Свободное место: " & Drive.FreeSpace \ 1000000 & " Мб.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("Доступное место: " & Drive.AvailableSpace \ 1000000 & " Мб.")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Системные папки")
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("Путь к диску: " & Drive.Path)
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("Корневая папка: " & Drive.RootFolder.Path)

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Системные имена")
                If Drive.VolumeName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Локальное имя тома: " & Drive.VolumeName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Локальное имя тома: " & "отсутствует")
                End If
                If Drive.ShareName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Сетевое имя тома: " & Drive.ShareName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Сетевое имя тома: " & "отсутствует")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Системные параметры")
                Select Case Drive.DriveType.ToString
                    Case 1
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "устройство со сменным носителем")
                    Case 2
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "жёсткий диск")
                    Case 3
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "сетевой диск")
                    Case 4
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "CD-ROM/DVD-ROM/BLURAY/HD-DVD")
                    Case 5
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "RAM-диск")
                    Case Else
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "абсолютно неизвестное устройство")
                End Select

                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Файловая система накопителя: " & Drive.FileSystem.ToString)
                If Drive.IsReady <> False Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Готовность: " & "готов")
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Готовность: " & "не готов")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Буква диска: " & Drive.DriveLetter)
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Серийный номер диска: " & Drive.SerialNumber)
            Else
                TreeView1.Nodes(i).Nodes.Add("Диск не готов, отсутствует носитель, недостаточно прав. Обратитесь к администратору сети.")
            End If
        Next

ErrorResolve:
        Exit Sub

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TreeView1.Nodes.Clear()

        FSO = CreateObject("Scripting.FileSystemObject")
        For Each Drive In FSO.Drives
            ListBox1.Items.Add(Drive.DriveLetter.ToString)
            TreeView1.Nodes.Add(Drive.DriveLetter.ToString)
        Next

        On Error Resume Next
        For i = 0 To TreeView1.Nodes.Count - 1
            Drive = FSO.GetDrive(ListBox1.Items.Item(i))
            If Drive.IsReady <> False Then
                TreeView1.Nodes(i).Nodes.Add("Параметры диска")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Полный объём диска: " & Drive.TotalSize \ 1000000 & " Мб.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("Занятое место: " & (Drive.TotalSize - Drive.FreeSpace) \ 1000000 & " Мб.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("Свободное место: " & Drive.FreeSpace \ 1000000 & " Мб.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("Доступное место: " & Drive.AvailableSpace \ 1000000 & " Мб.")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Системные папки")
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("Путь к диску: " & Drive.Path)
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("Корневая папка: " & Drive.RootFolder.Path)

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Системные имена")
                If Drive.VolumeName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Локальное имя тома: " & Drive.VolumeName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Локальное имя тома: " & "отсутствует")
                End If
                If Drive.ShareName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Сетевое имя тома: " & Drive.ShareName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("Сетевое имя тома: " & "отсутствует")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes.Add("Системные параметры")
                Select Case Drive.DriveType.ToString
                    Case 1
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "устройство со сменным носителем")
                    Case 2
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "жёсткий диск")
                    Case 3
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "сетевой диск")
                    Case 4
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "CD-ROM/DVD-ROM/BLURAY/HD-DVD")
                    Case 5
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "RAM-диск")
                    Case Else
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Тип накопителя: " & "абсолютно неизвестное устройство")
                End Select

                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Файловая система накопителя: " & Drive.FileSystem.ToString)
                If Drive.IsReady <> False Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Готовность: " & "готов")
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Готовность: " & "не готов")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Буква диска: " & Drive.DriveLetter)
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" Серийный номер диска: " & Drive.SerialNumber)
            Else
                TreeView1.Nodes(i).Nodes.Add("Диск не готов, отсутствует носитель, недостаточно прав. Обратитесь к администратору сети.")
            End If
        Next

ErrorResolve:
        Exit Sub
    End Sub

   
End Class
