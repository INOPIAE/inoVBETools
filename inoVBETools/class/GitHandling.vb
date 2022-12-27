Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Threading

Public Class GitHandling

    Public Structure GitStatusEntry
        Public FileName As String
        Public Area As String
        Public Type As String
        Public Sub New(FN As String, A As String, T As String)
            FileName = FN
            Area = A
            Type = T
        End Sub
    End Structure

    Public GitStatusEntries As New List(Of GitStatusEntry)
    Public WorkingDirectory As String = ""
    Public CurrentBranch As String = ""
    Public LastCommitMessage As String = ""

    Private GitPath As String

    Public Sub New(GitExe As String)
        GitPath = GitExe
    End Sub

    Public Sub AppendToGitIgnoreFile(strPath As String, strIgnore As String)
        My.Computer.FileSystem.WriteAllText(Path.Combine(strPath, ".gitignore"), strIgnore, True)
    End Sub

    Public Sub GetGitFilesStatus(Optional blnClear As Boolean = True)

        Dim pr As New Process
        pr.StartInfo.FileName = GitPath
        pr.StartInfo.Arguments = "status"
        pr.StartInfo.WorkingDirectory = WorkingDirectory
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardOutput = True
        pr.Start()

        If blnClear = True Then
            GitStatusEntries.Clear()
        End If

        Dim sOutput As String = ""
        Dim blnNew As Boolean
        Dim blnChanged As Boolean
        Dim blnStaged As Boolean
        Using oStreamReader As System.IO.StreamReader = pr.StandardOutput
            Do While oStreamReader.Peek() >= 0
                Dim strLine As String = oStreamReader.ReadLine()
                If strLine.StartsWith("On branch") Then CurrentBranch = strLine.Substring(10)
                If strLine.Trim.StartsWith("no changes added to commit") Then blnNew = False
                If strLine.Trim.StartsWith("Untracked files:") Then blnChanged = False
                If strLine.Trim.StartsWith("Untracked files:") Then blnStaged = False
                If strLine.Trim.StartsWith("Changes not staged for commit:") Then blnStaged = False


                If blnStaged = True And strLine.Length > 1 Then
                    If IsNothing(FindGitStatusEntryByName(strLine.Trim.Substring(0, 12).Trim)) Then
                        GitStatusEntries.Add(New GitStatusEntry(strLine.Trim.Substring(12), My.Resources.GH_Stashed, strLine.Trim.Substring(0, 12).Trim))
                    End If
                End If

                If blnChanged = True And strLine.Length > 1 Then
                    GitStatusEntries.Add(New GitStatusEntry(strLine.Trim.Substring(12), My.Resources.GH_Changed, strLine.Trim.Substring(0, 12).Trim))
                End If

                If blnNew = True And strLine.Length > 1 Then
                    If strLine.Trim.EndsWith("/") = False Then
                        GitStatusEntries.Add(New GitStatusEntry(strLine.Trim, My.Resources.GH_New, "new file:"))
                    End If
                End If

                If strLine.Contains("to include in what will be committed)") Then blnNew = True
                If strLine.Contains("to discard changes in working directory)") Then blnChanged = True
                If strLine.Contains(" to unstage)") Then blnStaged = True
            Loop

        End Using

    End Sub

    Public Sub GetGitFilesStatusLastCommit()

        Dim pr As New Process
        pr.StartInfo.FileName = GitPath
        pr.StartInfo.Arguments = "log --name-status HEAD^..HEAD"
        pr.StartInfo.WorkingDirectory = WorkingDirectory
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardOutput = True
        pr.Start()

        GitStatusEntries.Clear()

        Dim sOutput As String = ""
        Dim intLine As Int16
        Using oStreamReader As System.IO.StreamReader = pr.StandardOutput
            Do While oStreamReader.Peek() >= 0
                Dim strLine As String = oStreamReader.ReadLine()
                intLine += 1

                If intLine = 5 Then
                    LastCommitMessage = strLine.Trim
                End If

                If intLine > 5 Then
                    If strLine.Trim.Length > 0 Then
                        Dim strType As String = ""
                        Select Case strLine.Trim.Substring(0, 1)
                            Case "M"
                                strType = "modified:"
                            Case "D"
                                strType = "deleted:"
                            Case "N"
                                strType = "new file:"
                        End Select

                        GitStatusEntries.Add(New GitStatusEntry(strLine.Trim.Substring(1), My.Resources.GH_Stashed, strType))
                    End If
                End If

            Loop

        End Using

    End Sub

    Public Sub GitCommand(strCommand As String, strFile As String)
        Dim pr As New Process
        pr.StartInfo.FileName = GitPath
        pr.StartInfo.Arguments = strCommand & " " & Chr(34) & strFile & Chr(34)
        pr.StartInfo.WorkingDirectory = WorkingDirectory
        pr.StartInfo.CreateNoWindow = True
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardOutput = True
        pr.Start()
        Do Until pr.HasExited = True
            pr.Refresh()
            Thread.Sleep(1000)
        Loop
    End Sub

    Public Sub GitCommit(strCommitMessage As String)
        Dim pr As New Process
        pr.StartInfo.FileName = GitPath
        pr.StartInfo.Arguments = "commit -m " & Chr(34) & strCommitMessage & Chr(34)
        pr.StartInfo.WorkingDirectory = WorkingDirectory
        pr.StartInfo.CreateNoWindow = True
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardOutput = True
        pr.Start()
        Do Until pr.HasExited = True
            pr.Refresh()
            Thread.Sleep(1000)
        Loop
    End Sub

    Public Function FindGitStatusEntryByName(FileName As String) As GitStatusEntry
        For Each g As GitStatusEntry In GitStatusEntries
            If g.FileName = FileName Then
                Return g
            End If
        Next
        Return Nothing
    End Function

    Public Function IsDirectoryRepo(strPath As String) As Boolean
        Return Directory.Exists(Path.Combine(strPath, ".git"))
    End Function

    Public Sub InitializeRepo(strPath As String)
        Dim pr As New Process
        pr.StartInfo.FileName = GitPath
        pr.StartInfo.Arguments = "init"
        pr.StartInfo.WorkingDirectory = strPath
        pr.StartInfo.CreateNoWindow = True
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardOutput = True
        pr.Start()
        Do Until pr.HasExited = True
            pr.Refresh()
            Thread.Sleep(1000)
        Loop
    End Sub
End Class
