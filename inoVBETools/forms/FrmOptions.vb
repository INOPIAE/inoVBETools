Imports System.Globalization
Imports inoVBETools.My.Resources

Public Class FrmOptions
    Private Sub CmdOK_Click(sender As Object, e As EventArgs) Handles CmdOK.Click
        My.Settings.Language = Me.CboLanguage.Text
        My.Settings.GotoError = Me.TxtErrHandling.Text
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

        Me.TxtErrHandling.Text = My.Settings.GotoError

        Me.LblLanguage.Text = inoVBETools.My.Resources.frmOptionsLanguage
        Me.LblErrHandling.Text = inoVBETools.My.Resources.FrmOptionsNameOfGoToStatement
        Me.CmdCancel.Text = inoVBETools.My.Resources.frmButtonCancel
        Me.CmdOK.Text = inoVBETools.My.Resources.frmButtonOK
        Me.LblLangInfo.Text = inoVBETools.My.Resources.FrmOptionsLangInfo
    End Sub
End Class