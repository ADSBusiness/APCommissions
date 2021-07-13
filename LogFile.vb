Imports System.IO
Imports System.Text
Module LogFile

    Sub CreateFile()


        Dim path As String = "c:\temp\MyTest.txt"

        ' Create or overwrite the file.
        Dim fs As FileStream = File.Create(path)

        ' Add text to the file.
        Dim info As Byte() = New UTF8Encoding(True).GetBytes("This is some text in the file.")
        fs.Write(info, 0, info.Length)
        fs.Close()





    End Sub

    Sub OpenFile()

        Dim foundFile As String = ""
        foundFile = System.DateTime.Now & "  -  " & vbCrLf
        My.Computer.FileSystem.WriteAllText("c:\temp\MyTest.txt", foundFile, True)

        foundFile = System.DateTime.Now & "  -  " & vbCrLf
        My.Computer.FileSystem.WriteAllText("c:\temp\MyTest.txt", foundFile, True)

        foundFile = System.DateTime.Now & "  -  " & vbCrLf
        My.Computer.FileSystem.WriteAllText("c:\temp\MyTest.txt", foundFile, True)

        foundFile = System.DateTime.Now & "  -  " & vbCrLf
        My.Computer.FileSystem.WriteAllText("c:\temp\MyTest.txt", foundFile, True)

        foundFile = System.DateTime.Now & "  -  " & vbCrLf
        My.Computer.FileSystem.WriteAllText("c:\temp\MyTest.txt", foundFile, True)



        Dim path As String = "c:\temp\MyTest.txt"

        ' Create or overwrite the file.
        Dim fs As FileStream = File.OpenWrite(path)

        Dim info As Byte() = New UTF8Encoding(True).GetBytes("This is some text in the file.")
        fs.Write(info, 0, info.Length)
        fs.Close()



    End Sub


End Module
