Imports System.Configuration
Imports System.Globalization
Imports System.Net.Configuration
Imports System.Security
Imports System.Windows.Forms
Imports System.Xml.Serialization
Imports inoVBETools.My.Resources
Imports Microsoft.SqlServer.Server

Public Class FrmOptions
    Private Sub CmdOK_Click(sender As Object, e As EventArgs) Handles CmdOK.Click
        My.Settings.Language = Me.CboLanguage.Text
        My.Settings.GotoError = Me.TxtErrHandling.Text
        My.Settings.MakeBackup = Me.ChbBackup.Checked
        My.Settings.KeepBackup = Me.ChbKeepBackup.Checked
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

        Me.LblLanguage.Text = inoVBETools.My.Resources.frmOptionsLanguage
        Me.LblErrHandling.Text = inoVBETools.My.Resources.FrmOptionsNameOfGoToStatement
        Me.CmdCancel.Text = inoVBETools.My.Resources.frmButtonCancel
        Me.CmdOK.Text = inoVBETools.My.Resources.frmButtonOK
        Me.LblLangInfo.Text = inoVBETools.My.Resources.FrmOptionsLangInfo
        Me.ChbBackup.Checked = My.Settings.MakeBackup
        Me.ChbKeepBackup.Checked = My.Settings.KeepBackup
    End Sub
End Class