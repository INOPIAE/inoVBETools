<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmOptions
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CboLanguage = New System.Windows.Forms.ComboBox()
        Me.LblLanguage = New System.Windows.Forms.Label()
        Me.CmdCancel = New System.Windows.Forms.Button()
        Me.CmdOK = New System.Windows.Forms.Button()
        Me.LblLangInfo = New System.Windows.Forms.Label()
        Me.TxtErrHandling = New System.Windows.Forms.TextBox()
        Me.LblErrHandling = New System.Windows.Forms.Label()
        Me.GrpImport = New System.Windows.Forms.GroupBox()
        Me.ChbBackup = New System.Windows.Forms.CheckBox()
        Me.ChbKeepBackup = New System.Windows.Forms.CheckBox()
        Me.GrpImport.SuspendLayout()
        Me.SuspendLayout()
        '
        'CboLanguage
        '
        Me.CboLanguage.FormattingEnabled = True
        Me.CboLanguage.Location = New System.Drawing.Point(187, 42)
        Me.CboLanguage.Name = "CboLanguage"
        Me.CboLanguage.Size = New System.Drawing.Size(154, 21)
        Me.CboLanguage.TabIndex = 0
        '
        'LblLanguage
        '
        Me.LblLanguage.AutoSize = True
        Me.LblLanguage.Location = New System.Drawing.Point(44, 45)
        Me.LblLanguage.Name = "LblLanguage"
        Me.LblLanguage.Size = New System.Drawing.Size(55, 13)
        Me.LblLanguage.TabIndex = 1
        Me.LblLanguage.Text = "Language"
        '
        'CmdCancel
        '
        Me.CmdCancel.Location = New System.Drawing.Point(99, 329)
        Me.CmdCancel.Name = "CmdCancel"
        Me.CmdCancel.Size = New System.Drawing.Size(88, 28)
        Me.CmdCancel.TabIndex = 2
        Me.CmdCancel.Text = "Cancel"
        Me.CmdCancel.UseVisualStyleBackColor = True
        '
        'CmdOK
        '
        Me.CmdOK.Location = New System.Drawing.Point(543, 329)
        Me.CmdOK.Name = "CmdOK"
        Me.CmdOK.Size = New System.Drawing.Size(88, 28)
        Me.CmdOK.TabIndex = 2
        Me.CmdOK.Text = "OK"
        Me.CmdOK.UseVisualStyleBackColor = True
        '
        'LblLangInfo
        '
        Me.LblLangInfo.AutoSize = True
        Me.LblLangInfo.Location = New System.Drawing.Point(44, 223)
        Me.LblLangInfo.Name = "LblLangInfo"
        Me.LblLangInfo.Size = New System.Drawing.Size(76, 13)
        Me.LblLangInfo.TabIndex = 3
        Me.LblLangInfo.Text = "Language Info"
        '
        'TxtErrHandling
        '
        Me.TxtErrHandling.Location = New System.Drawing.Point(187, 84)
        Me.TxtErrHandling.Name = "TxtErrHandling"
        Me.TxtErrHandling.Size = New System.Drawing.Size(154, 20)
        Me.TxtErrHandling.TabIndex = 4
        '
        'LblErrHandling
        '
        Me.LblErrHandling.AutoSize = True
        Me.LblErrHandling.Location = New System.Drawing.Point(44, 87)
        Me.LblErrHandling.Name = "LblErrHandling"
        Me.LblErrHandling.Size = New System.Drawing.Size(60, 13)
        Me.LblErrHandling.TabIndex = 1
        Me.LblErrHandling.Text = "Errhandling"
        '
        'GrpImport
        '
        Me.GrpImport.Controls.Add(Me.ChbKeepBackup)
        Me.GrpImport.Controls.Add(Me.ChbBackup)
        Me.GrpImport.Location = New System.Drawing.Point(47, 110)
        Me.GrpImport.Name = "GrpImport"
        Me.GrpImport.Size = New System.Drawing.Size(381, 110)
        Me.GrpImport.TabIndex = 5
        Me.GrpImport.TabStop = False
        Me.GrpImport.Text = "ImportCode"
        '
        'ChbBackup
        '
        Me.ChbBackup.AutoSize = True
        Me.ChbBackup.Location = New System.Drawing.Point(22, 19)
        Me.ChbBackup.Name = "ChbBackup"
        Me.ChbBackup.Size = New System.Drawing.Size(63, 17)
        Me.ChbBackup.TabIndex = 0
        Me.ChbBackup.Text = "Backup"
        Me.ChbBackup.UseVisualStyleBackColor = True
        '
        'ChbKeepBackup
        '
        Me.ChbKeepBackup.AutoSize = True
        Me.ChbKeepBackup.Location = New System.Drawing.Point(22, 42)
        Me.ChbKeepBackup.Name = "ChbKeepBackup"
        Me.ChbKeepBackup.Size = New System.Drawing.Size(88, 17)
        Me.ChbKeepBackup.TabIndex = 0
        Me.ChbKeepBackup.Text = "KeepBackup"
        Me.ChbKeepBackup.UseVisualStyleBackColor = True
        '
        'FrmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.GrpImport)
        Me.Controls.Add(Me.TxtErrHandling)
        Me.Controls.Add(Me.LblLangInfo)
        Me.Controls.Add(Me.CmdOK)
        Me.Controls.Add(Me.CmdCancel)
        Me.Controls.Add(Me.LblErrHandling)
        Me.Controls.Add(Me.LblLanguage)
        Me.Controls.Add(Me.CboLanguage)
        Me.Name = "FrmOptions"
        Me.Text = "FrmOptions"
        Me.GrpImport.ResumeLayout(False)
        Me.GrpImport.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CboLanguage As Windows.Forms.ComboBox
    Friend WithEvents LblLanguage As Windows.Forms.Label
    Friend WithEvents CmdCancel As Windows.Forms.Button
    Friend WithEvents CmdOK As Windows.Forms.Button
    Friend WithEvents LblLangInfo As Windows.Forms.Label
    Friend WithEvents TxtErrHandling As Windows.Forms.TextBox
    Friend WithEvents LblErrHandling As Windows.Forms.Label
    Friend WithEvents GrpImport As Windows.Forms.GroupBox
    Friend WithEvents ChbKeepBackup As Windows.Forms.CheckBox
    Friend WithEvents ChbBackup As Windows.Forms.CheckBox
End Class
