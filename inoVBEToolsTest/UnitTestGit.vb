Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports inoVBETools
Imports inoVBETools.GitHandling
Imports Xunit

Namespace inoVBEToolsTest
    Public Class UnitTestGit
        Public ClsGit As New inoVBETools.GitHandling("C:\Program Files\Git\cmd\git.exe")

        Private TestFolder As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Replace("bin\Debug\net6.0", "TestData")

        <Fact>
        Sub TestFindGitStatusEntryByName()

            ClsGit.GitStatusEntries.Clear()

            ClsGit.GitStatusEntries.Add(New GitStatusEntry("firstfile", "Stashed", "new file:"))
            ClsGit.GitStatusEntries.Add(New GitStatusEntry("secondfile", "Stashed", "new file:"))

            Dim g As New GitStatusEntry
            g = ClsGit.FindGitStatusEntryByName("firstfile")
            Assert.Equal("firstfile", g.FileName)
            Assert.Equal("Stashed", "Stashed")
            Assert.Equal("new file:", g.Type)

            g = ClsGit.FindGitStatusEntryByName("secondfile")
            Assert.Equal("secondfile", g.FileName)
            Assert.Equal("Stashed", "Stashed")
            Assert.Equal("new file:", g.Type)

            g = ClsGit.FindGitStatusEntryByName("somethingelse")
            Assert.Equal(Nothing, g)

        End Sub

        <Fact>
        Sub TestInitRepo()
            Dim strPath As String = Path.Combine(TestFolder, "GitTest")
            If Directory.Exists(strPath) Then
                Directory.Delete(strPath, True)
            End If
            Directory.CreateDirectory(strPath)

            Assert.False(ClsGit.IsDirectoryRepo(strPath))

            ClsGit.InitializeRepo(strPath)

            Assert.True(ClsGit.IsDirectoryRepo(strPath))

            Directory.Delete(strPath, True)
        End Sub

    End Class
End Namespace
