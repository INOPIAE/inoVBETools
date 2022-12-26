Imports System.ComponentModel
Imports inoVBETools
Imports inoVBETools.GitHandling
Imports Xunit

Namespace inoVBEToolsTest
    Public Class UnitTestGit
        Public ClsGit As New inoVBETools.GitHandling

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
    End Class
End Namespace
