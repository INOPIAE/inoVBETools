Imports System.IO
Imports System.Reflection
Imports inoVBETools
Imports inoVBETools.VBEHandling
Imports Xunit

Namespace inoVBEToolsTest
    Public Class UnitTestVBEHandling

        Private ClsVBEH As New VBEHandling

        Private TestFolder As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Replace("bin\Debug\net6.0", "TestData")

        <Fact>
        Sub TestAddProjectEntry()
            ClsVBEH.ProjectEntries.Clear()

            Assert.Equal(0, ClsVBEH.ProjectEntries.Count)

            ClsVBEH.ProjectAdd("A", "CDA")

            Assert.Equal(1, ClsVBEH.ProjectEntries.Count)
            Assert.Equal("A", ClsVBEH.ProjectEntries(0).ProjectName)
            Assert.Equal("CDA", ClsVBEH.ProjectEntries(0).CodeDirectrory)

            ClsVBEH.ProjectAdd("B", "CDB")

            Assert.Equal(2, ClsVBEH.ProjectEntries.Count)
            Assert.Equal("B", ClsVBEH.ProjectEntries(1).ProjectName)
            Assert.Equal("CDB", ClsVBEH.ProjectEntries(1).CodeDirectrory)

            ClsVBEH.ProjectAdd("A", "CDAA")

            Assert.Equal(2, ClsVBEH.ProjectEntries.Count)
            Assert.Equal("A", ClsVBEH.ProjectEntries(ClsVBEH.ProjectEntries.Count - 1).ProjectName)
            Assert.Equal("CDAA", ClsVBEH.ProjectEntries(ClsVBEH.ProjectEntries.Count - 1).CodeDirectrory)

        End Sub

        <Fact>
        Sub TestGetProjectDirectory()
            ClsVBEH.ProjectEntries.Clear()

            ClsVBEH.ProjectAdd("A", "CDA")
            ClsVBEH.ProjectAdd("B", "CDB")

            Assert.Equal("CDA", ClsVBEH.ProjectDirectoryByName("A"))
            Assert.Equal("CDB", ClsVBEH.ProjectDirectoryByName("B"))
            Assert.Null(ClsVBEH.ProjectDirectoryByName("C"))
            Assert.Null(ClsVBEH.ProjectDirectoryByName("D"))

            ClsVBEH.ProjectAdd("C", "CDC")
            Assert.Equal("CDA", ClsVBEH.ProjectDirectoryByName("A"))
            Assert.Equal("CDB", ClsVBEH.ProjectDirectoryByName("B"))
            Assert.Equal("CDC", ClsVBEH.ProjectDirectoryByName("C"))
            Assert.Null(ClsVBEH.ProjectDirectoryByName("D"))
        End Sub

        <Fact>
        Sub TestWriteProjectDirectory()
            ClsVBEH.ProjectEntries.Clear()
            ClsVBEH.ProjectAdd("A", "CDA")
            ClsVBEH.ProjectAdd("B", "CDB")
            ClsVBEH.ProjectAdd("C", "CDC")

            Dim strPath As String = Path.Combine(TestFolder, "WriteProject.txt")
            File.Delete(strPath)
            ClsVBEH.WriteProjectEntries(strPath)

            Using sr As New StreamReader(strPath)
                Assert.Equal("A; CDA", sr.ReadLine())
                Assert.Equal("B; CDB", sr.ReadLine())
                Assert.Equal("C; CDC", sr.ReadLine())
            End Using
        End Sub

        <Fact>
        Sub TestReadProjectDirectory()
            ClsVBEH.ProjectEntries.Clear()

            Dim strPath As String = Path.Combine(TestFolder, "WriteProject.txt")
            File.Delete(strPath)
            Using sw As New StreamWriter(strPath)

                sw.WriteLine("A; TA")
                sw.WriteLine("B; TB")
                sw.WriteLine("C; TC")
            End Using

            ClsVBEH.ReadProjectEntries(strPath)

            Assert.Equal(3, ClsVBEH.ProjectEntries.Count)
            Assert.Equal("A", ClsVBEH.ProjectEntries(0).ProjectName)
            Assert.Equal("TA", ClsVBEH.ProjectEntries(0).CodeDirectrory)
            Assert.Equal("B", ClsVBEH.ProjectEntries(1).ProjectName)
            Assert.Equal("TB", ClsVBEH.ProjectEntries(1).CodeDirectrory)
            Assert.Equal("C", ClsVBEH.ProjectEntries(2).ProjectName)
            Assert.Equal("TC", ClsVBEH.ProjectEntries(2).CodeDirectrory)
        End Sub


    End Class
End Namespace
