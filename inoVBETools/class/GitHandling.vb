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
    Public WorkingDirectory As String = My.Settings.WorkingDirectory
    Public CurrentBranch As String

    Public Sub AppendToGitIgnoreFile(strPath As String, strIgnore As String)
        My.Computer.FileSystem.WriteAllText(Path.Combine(strPath, ".gitignore"), strIgnore, True)
    End Sub

    Public Sub GetGitFilesStatus()

        Dim pr As New Process
        pr.StartInfo.FileName = My.Settings.Git_Exe
        pr.StartInfo.Arguments = "status"
        pr.StartInfo.WorkingDirectory = WorkingDirectory
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardOutput = True
        pr.Start()

        GitStatusEntries.Clear()

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
                    GitStatusEntries.Add(New GitStatusEntry(strLine.Trim.Substring(12), My.Resources.GH_Stashed, strLine.Trim.Substring(0, 12).Trim))
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

    Public Sub GitCommand(strCommand As String, strFile As String)
        Dim pr As New Process
        pr.StartInfo.FileName = My.Settings.Git_Exe
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
        pr.StartInfo.FileName = My.Settings.Git_Exe
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
End Class
