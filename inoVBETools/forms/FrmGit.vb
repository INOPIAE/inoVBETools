Imports System.Drawing
Imports System.Management.Instrumentation
Imports System.Net.Http
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Xml

Public Class FrmGit
    Private ClsGit As New GitHandling
    Private Sub cmdTest_Click(sender As Object, e As EventArgs) Handles cmdTest.Click
        PopulateTreeView()
    End Sub

    Private Sub PopulateTreeView()
        ClsGit.GetGitFilesStatus()

        TvGit.Nodes.Clear()
        TvGit.ShowLines = False
        TvGit.CheckBoxes = True

        Dim nArea As TreeNode
        Dim strArea As String = ""
        For Each g As GitHandling.GitStatusEntry In ClsGit.GitStatusEntries
            If strArea <> g.Area Then
                strArea = g.Area
                nArea = TvGit.Nodes.Add(g.Area)
                nArea.Name = g.Area
                nArea.Checked = True
            End If
            Dim NewNode As TreeNode = nArea.Nodes.Add(g.Type & g.FileName)
            NewNode.Checked = True
            Select Case g.Area
                Case My.Resources.GH_New
                    NewNode.ForeColor = My.Settings.GitColorNew
                Case My.Resources.GH_Changed
                    NewNode.ForeColor = My.Settings.GitColorChanged
                Case My.Resources.GH_Stashed
                    NewNode.ForeColor = My.Settings.GitColorStashed
            End Select
        Next

        TvGit.ExpandAll()

        LblCurrentBranch.Text = String.Format(My.Resources.frmGitCurrentBranch, ClsGit.CurrentBranch)
    End Sub

    Private Sub TvGit_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TvGit.AfterCheck
        If e.Action <> TreeViewAction.Unknown Then
            If e.Node.Nodes.Count > 0 Then
                Me.CheckAllChildNodes(e.Node, e.Node.Checked)
            End If
        End If
    End Sub

    Private Sub CheckAllChildNodes(treeNode As TreeNode, nodeChecked As Boolean)
        Dim node As TreeNode
        For Each node In treeNode.Nodes
            node.Checked = nodeChecked
            If node.Nodes.Count > 0 Then
                Me.CheckAllChildNodes(node, nodeChecked)
            End If
        Next node
    End Sub

    Private Sub FrmGit_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = inoVBETools.My.Resources.FrmGitCaption
        Me.lblCommit.Text = inoVBETools.My.Resources.FrmGitLblCommitMsg
        Me.CmdAdd.Text = inoVBETools.My.Resources.FrmButtonAddToStage
        Me.CmdRemove.Text = inoVBETools.My.Resources.FrmButtonRemoveFromStage
        Me.CmdOK.Text = My.Resources.frmButtonOK

        PopulateTreeView()
    End Sub

    Private Sub CmdAdd_Click(sender As Object, e As EventArgs) Handles CmdAdd.Click
        Dim np() As TreeNode = TvGit.Nodes.Find(My.Resources.GH_Changed, True)
        If np.Count = 1 Then
            For Each n As TreeNode In np(0).Nodes
                If n.Checked = True Then
                    ClsGit.GitCommand("add", n.Text.Split(":").Last.Trim)
                End If
            Next
        End If
        np = TvGit.Nodes.Find(My.Resources.GH_New, True)
        If np.Count = 1 Then
            For Each n As TreeNode In np(0).Nodes
                If n.Checked = True Then
                    ClsGit.GitCommand("add", n.Text.Split(":").Last.Trim)
                End If
            Next
        End If
        PopulateTreeView()
    End Sub

    Private Sub CmdRemove_Click(sender As Object, e As EventArgs) Handles CmdRemove.Click
        Dim np() As TreeNode = TvGit.Nodes.Find(My.Resources.GH_Stashed, True)
        If np.Count = 1 Then
            For Each n As TreeNode In np(0).Nodes
                If n.Checked = True Then
                    ClsGit.GitCommand("restore --staged", n.Text.Split(":").Last.Trim)
                End If
            Next
        End If
        PopulateTreeView()
    End Sub

    Private Sub CmdCommit_Click(sender As Object, e As EventArgs) Handles CmdCommit.Click
        If Me.TxtCommit.Text.Trim = "" Then
            MessageBox.Show(inoVBETools.My.Resources.msgContinue)
            Me.TxtCommit.Select()
            Exit Sub
        End If
        Dim np() As TreeNode = TvGit.Nodes.Find(My.Resources.GH_Changed, False)
        If np.Count = 1 Then
            For Each n As TreeNode In np(0).Nodes
                If n.Checked = True Then
                    ClsGit.GitCommand("add", n.Text.Split(":").Last.Trim)
                End If
            Next
        End If
        np = TvGit.Nodes.Find(My.Resources.GH_New, False)
        If np.Count = 1 Then
            For Each n As TreeNode In np(0).Nodes
                If n.Checked = True Then
                    ClsGit.GitCommand("add", n.Text.Split(":").Last.Trim)
                End If
            Next
        End If
        ClsGit.GitCommit(Me.TxtCommit.Text)
        PopulateTreeView()
    End Sub

    Private Sub CmdOK_Click(sender As Object, e As EventArgs) Handles CmdOK.Click
        Me.Close()
    End Sub
End Class