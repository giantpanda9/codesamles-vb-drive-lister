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
                TreeView1.Nodes(i).Nodes.Add("��������� �����")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("������ ����� �����: " & Drive.TotalSize \ 1000000 & " ��.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("������� �����: " & (Drive.TotalSize - Drive.FreeSpace) \ 1000000 & " ��.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("��������� �����: " & Drive.FreeSpace \ 1000000 & " ��.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("��������� �����: " & Drive.AvailableSpace \ 1000000 & " ��.")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("��������� �����")
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("���� � �����: " & Drive.Path)
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("�������� �����: " & Drive.RootFolder.Path)

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("��������� �����")
                If Drive.VolumeName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("��������� ��� ����: " & Drive.VolumeName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("��������� ��� ����: " & "�����������")
                End If
                If Drive.ShareName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("������� ��� ����: " & Drive.ShareName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("������� ��� ����: " & "�����������")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes.Add("��������� ���������")
                Select Case Drive.DriveType.ToString
                    Case 1
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "���������� �� ������� ���������")
                    Case 2
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "������ ����")
                    Case 3
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "������� ����")
                    Case 4
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "CD-ROM/DVD-ROM/BLURAY/HD-DVD")
                    Case 5
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "RAM-����")
                    Case Else
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "��������� ����������� ����������")
                End Select

                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" �������� ������� ����������: " & Drive.FileSystem.ToString)
                If Drive.IsReady <> False Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ����������: " & "�����")
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ����������: " & "�� �����")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ����� �����: " & Drive.DriveLetter)
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" �������� ����� �����: " & Drive.SerialNumber)
            Else
                TreeView1.Nodes(i).Nodes.Add("���� �� �����, ����������� ��������, ������������ ����. ���������� � �������������� ����.")
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
                TreeView1.Nodes(i).Nodes.Add("��������� �����")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("������ ����� �����: " & Drive.TotalSize \ 1000000 & " ��.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("������� �����: " & (Drive.TotalSize - Drive.FreeSpace) \ 1000000 & " ��.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("��������� �����: " & Drive.FreeSpace \ 1000000 & " ��.")
                TreeView1.Nodes(i).Nodes(0).Nodes(0).Nodes.Add("��������� �����: " & Drive.AvailableSpace \ 1000000 & " ��.")

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("��������� �����")
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("���� � �����: " & Drive.Path)
                TreeView1.Nodes(i).Nodes(0).Nodes(1).Nodes.Add("�������� �����: " & Drive.RootFolder.Path)

                TreeView1.Nodes(i).Nodes(0).Nodes.Add("��������� �����")
                If Drive.VolumeName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("��������� ��� ����: " & Drive.VolumeName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("��������� ��� ����: " & "�����������")
                End If
                If Drive.ShareName <> "" Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("������� ��� ����: " & Drive.ShareName)
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(2).Nodes.Add("������� ��� ����: " & "�����������")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes.Add("��������� ���������")
                Select Case Drive.DriveType.ToString
                    Case 1
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "���������� �� ������� ���������")
                    Case 2
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "������ ����")
                    Case 3
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "������� ����")
                    Case 4
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "CD-ROM/DVD-ROM/BLURAY/HD-DVD")
                    Case 5
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "RAM-����")
                    Case Else
                        TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ��� ����������: " & "��������� ����������� ����������")
                End Select

                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" �������� ������� ����������: " & Drive.FileSystem.ToString)
                If Drive.IsReady <> False Then
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ����������: " & "�����")
                Else
                    TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ����������: " & "�� �����")
                End If
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" ����� �����: " & Drive.DriveLetter)
                TreeView1.Nodes(i).Nodes(0).Nodes(3).Nodes.Add(" �������� ����� �����: " & Drive.SerialNumber)
            Else
                TreeView1.Nodes(i).Nodes.Add("���� �� �����, ����������� ��������, ������������ ����. ���������� � �������������� ����.")
            End If
        Next

ErrorResolve:
        Exit Sub
    End Sub

   
End Class
