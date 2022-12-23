Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms


Public Class FrmOptions
    Private Sub CmdOK_Click(sender As Object, e As EventArgs) Handles CmdOK.Click
        My.Settings.Language = Me.CboLanguage.Text
        My.Settings.GotoError = Me.TxtErrHandling.Text
        My.Settings.MakeBackup = Me.ChbBackup.Checked
        My.Settings.KeepBackup = Me.ChbKeepBackup.Checked
        My.Settings.Git_Exe = Me.TxtGit.Text
        My.Settings.GitColorNew = Me.PBGitNew.BackColor
        My.Settings.GitColorChanged = Me.PbGitChanged.BackColor
        My.Settings.GitColorStashed = Me.PbGitStashed.BackColor
        My.Settings.Save()
        My.Application.ChangeUICulture(My.Settings.Language)
        Me.Close()
    End Sub

    Private Sub CmdCancel_Click(sender As Object, e As EventArgs) Handles CmdCancel.Click
        Me.Close()
    End Sub

    Private Sub FrmOptions_Load(sender As Object, e As EventArgs) Handles Me.Load


        Dim EN As CultureInfo = CultureInfo.GetCultureInfo("en-US")
        Dim DE As CultureInfo = CultureInfo.GetCultureInfo("de-DE")

        Me.CboLanguage.Items.Add(EN.Name)
        Me.CboLanguage.Items.Add(DE.Name)

        Me.CboLanguage.Text = My.Settings.Language

        Me.GrpImport.Text = inoVBETools.My.Resources.FrmOptionsImportGrp
        Me.ChbBackup.Text = inoVBETools.My.Resources.FrmOptionsCreateBackup
        Me.ChbKeepBackup.Text = inoVBETools.My.Resources.FrmOptionsKeepBackup

        Me.TxtErrHandling.Text = My.Settings.GotoError

        Me.Text = inoVBETools.My.Resources.FrmOptionsCaption

        Me.GrpGit.Text = My.Resources.FrmOptionsGitGrp
        Me.LblGit.Text = My.Resources.FrmOptionsGitLocation

        Me.LblColour.Text = inoVBETools.My.Resources.frmOptionsColourGit
        Me.LblGitNew.Text = My.Resources.GH_New
        Me.PBGitNew.BackColor = My.Settings.GitColorNew
        Me.LblGitChanged.Text = My.Resources.GH_Changed
        Me.PbGitChanged.BackColor = My.Settings.GitColorChanged
        Me.LblGitStashed.Text = My.Resources.GH_Stashed
        Me.PbGitStashed.BackColor = My.Settings.GitColorStashed

        Me.LblLanguage.Text = inoVBETools.My.Resources.frmOptionsLanguage
        Me.LblErrHandling.Text = inoVBETools.My.Resources.FrmOptionsNameOfGoToStatement
        Me.CmdCancel.Text = inoVBETools.My.Resources.frmButtonCancel
        Me.CmdOK.Text = inoVBETools.My.Resources.frmButtonOK
        Me.LblLangInfo.Text = inoVBETools.My.Resources.FrmOptionsLangInfo
        Me.ChbBackup.Checked = My.Settings.MakeBackup
        Me.ChbKeepBackup.Checked = My.Settings.KeepBackup
        Me.TxtGit.Text = My.Settings.Git_Exe
    End Sub

    Private Sub CmdGit_Click(sender As Object, e As EventArgs) Handles CmdGit.Click
        Dim ofd As New OpenFileDialog
        With ofd
            .Multiselect = False
            .Title = My.Resources.FrmOptionsTitelGitSearch
            If System.IO.File.Exists(Me.TxtGit.Text) Then
                .InitialDirectory = Path.GetDirectoryName(Me.TxtGit.Text)
                .FileName = Path.GetFileName(Me.TxtGit.Text)
            End If
            If .ShowDialog = DialogResult.OK Then
                Me.TxtGit.Text = .FileName
            End If
        End With
    End Sub

    Private Sub CmdGitNew_Click(sender As Object, e As EventArgs) Handles CmdGitNew.Click
        Dim cd As New ColorDialog
        With cd
            .Color = PBGitNew.BackColor
            If .ShowDialog = vbOK Then
                PBGitNew.BackColor = .Color
            End If
        End With
    End Sub

    Private Sub cmdGitChanged_Click(sender As Object, e As EventArgs) Handles cmdGitChanged.Click
        Dim cd As New ColorDialog
        With cd
            .Color = PbGitChanged.BackColor
            If .ShowDialog = vbOK Then
                PbGitChanged.BackColor = .Color
            End If
        End With
    End Sub

    Private Sub cmdGitStashed_Click(sender As Object, e As EventArgs) Handles cmdGitStashed.Click
        Dim cd As New ColorDialog
        With cd
            .Color = PbGitStashed.BackColor
            If .ShowDialog = vbOK Then
                PbGitStashed.BackColor = .Color
            End If
        End With
    End Sub
End Class