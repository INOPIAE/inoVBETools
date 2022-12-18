Imports System.IO

Public Class GitHandling

    Public Sub AppendToGitIgnoreFile(strPath As String, strIgnore As String)
        My.Computer.FileSystem.WriteAllText(Path.Combine(strPath, ".gitignore"), strIgnore, True)
    End Sub
End Class
