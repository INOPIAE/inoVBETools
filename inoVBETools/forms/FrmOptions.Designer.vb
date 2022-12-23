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
        Me.ChbKeepBackup = New System.Windows.Forms.CheckBox()
        Me.ChbBackup = New System.Windows.Forms.CheckBox()
        Me.GrpGit = New System.Windows.Forms.GroupBox()
        Me.LblGitStashed = New System.Windows.Forms.Label()
        Me.LblGitChanged = New System.Windows.Forms.Label()
        Me.LblGitNew = New System.Windows.Forms.Label()
        Me.PbGitStashed = New System.Windows.Forms.PictureBox()
        Me.PbGitChanged = New System.Windows.Forms.PictureBox()
        Me.PBGitNew = New System.Windows.Forms.PictureBox()
        Me.cmdGitStashed = New System.Windows.Forms.Button()
        Me.cmdGitChanged = New System.Windows.Forms.Button()
        Me.CmdGitNew = New System.Windows.Forms.Button()
        Me.CmdGit = New System.Windows.Forms.Button()
        Me.TxtGit = New System.Windows.Forms.TextBox()
        Me.LblGit = New System.Windows.Forms.Label()
        Me.LblColour = New System.Windows.Forms.Label()
        Me.GrpImport.SuspendLayout()
        Me.GrpGit.SuspendLayout()
        CType(Me.PbGitStashed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PbGitChanged, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PBGitNew, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CboLanguage
        '
        Me.CboLanguage.FormattingEnabled = True
        Me.CboLanguage.Location = New System.Drawing.Point(187, 42)
        Me.CboLanguage.Name = "CboLanguage"
        Me.CboLanguage.Size = New System.Drawing.Size(154, 21)
        Me.CboLanguage.TabIndex = 1
        '
        'LblLanguage
        '
        Me.LblLanguage.AutoSize = True
        Me.LblLanguage.Location = New System.Drawing.Point(44, 45)
        Me.LblLanguage.Name = "LblLanguage"
        Me.LblLanguage.Size = New System.Drawing.Size(55, 13)
        Me.LblLanguage.TabIndex = 0
        Me.LblLanguage.Text = "Language"
        '
        'CmdCancel
        '
        Me.CmdCancel.Location = New System.Drawing.Point(99, 373)
        Me.CmdCancel.Name = "CmdCancel"
        Me.CmdCancel.Size = New System.Drawing.Size(88, 28)
        Me.CmdCancel.TabIndex = 8
        Me.CmdCancel.Text = "Cancel"
        Me.CmdCancel.UseVisualStyleBackColor = True
        '
        'CmdOK
        '
        Me.CmdOK.Location = New System.Drawing.Point(544, 373)
        Me.CmdOK.Name = "CmdOK"
        Me.CmdOK.Size = New System.Drawing.Size(88, 28)
        Me.CmdOK.TabIndex = 7
        Me.CmdOK.Text = "OK"
        Me.CmdOK.UseVisualStyleBackColor = True
        '
        'LblLangInfo
        '
        Me.LblLangInfo.AutoSize = True
        Me.LblLangInfo.Location = New System.Drawing.Point(44, 74)
        Me.LblLangInfo.Name = "LblLangInfo"
        Me.LblLangInfo.Size = New System.Drawing.Size(76, 13)
        Me.LblLangInfo.TabIndex = 2
        Me.LblLangInfo.Text = "Language Info"
        '
        'TxtErrHandling
        '
        Me.TxtErrHandling.Location = New System.Drawing.Point(187, 95)
        Me.TxtErrHandling.Name = "TxtErrHandling"
        Me.TxtErrHandling.Size = New System.Drawing.Size(154, 20)
        Me.TxtErrHandling.TabIndex = 4
        '
        'LblErrHandling
        '
        Me.LblErrHandling.AutoSize = True
        Me.LblErrHandling.Location = New System.Drawing.Point(44, 98)
        Me.LblErrHandling.Name = "LblErrHandling"
        Me.LblErrHandling.Size = New System.Drawing.Size(60, 13)
        Me.LblErrHandling.TabIndex = 3
        Me.LblErrHandling.Text = "Errhandling"
        '
        'GrpImport
        '
        Me.GrpImport.Controls.Add(Me.ChbKeepBackup)
        Me.GrpImport.Controls.Add(Me.ChbBackup)
        Me.GrpImport.Location = New System.Drawing.Point(47, 121)
        Me.GrpImport.Name = "GrpImport"
        Me.GrpImport.Size = New System.Drawing.Size(703, 75)
        Me.GrpImport.TabIndex = 5
        Me.GrpImport.TabStop = False
        Me.GrpImport.Text = "ImportCode"
        '
        'ChbKeepBackup
        '
        Me.ChbKeepBackup.AutoSize = True
        Me.ChbKeepBackup.Location = New System.Drawing.Point(22, 42)
        Me.ChbKeepBackup.Name = "ChbKeepBackup"
        Me.ChbKeepBackup.Size = New System.Drawing.Size(88, 17)
        Me.ChbKeepBackup.TabIndex = 1
        Me.ChbKeepBackup.Text = "KeepBackup"
        Me.ChbKeepBackup.UseVisualStyleBackColor = True
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
        'GrpGit
        '
        Me.GrpGit.Controls.Add(Me.LblColour)
        Me.GrpGit.Controls.Add(Me.LblGitStashed)
        Me.GrpGit.Controls.Add(Me.LblGitChanged)
        Me.GrpGit.Controls.Add(Me.LblGitNew)
        Me.GrpGit.Controls.Add(Me.PbGitStashed)
        Me.GrpGit.Controls.Add(Me.PbGitChanged)
        Me.GrpGit.Controls.Add(Me.PBGitNew)
        Me.GrpGit.Controls.Add(Me.cmdGitStashed)
        Me.GrpGit.Controls.Add(Me.cmdGitChanged)
        Me.GrpGit.Controls.Add(Me.CmdGitNew)
        Me.GrpGit.Controls.Add(Me.CmdGit)
        Me.GrpGit.Controls.Add(Me.TxtGit)
        Me.GrpGit.Controls.Add(Me.LblGit)
        Me.GrpGit.Location = New System.Drawing.Point(47, 202)
        Me.GrpGit.Name = "GrpGit"
        Me.GrpGit.Size = New System.Drawing.Size(703, 165)
        Me.GrpGit.TabIndex = 6
        Me.GrpGit.TabStop = False
        Me.GrpGit.Text = "Git"
        '
        'LblGitStashed
        '
        Me.LblGitStashed.AutoSize = True
        Me.LblGitStashed.Location = New System.Drawing.Point(23, 134)
        Me.LblGitStashed.Name = "LblGitStashed"
        Me.LblGitStashed.Size = New System.Drawing.Size(46, 13)
        Me.LblGitStashed.TabIndex = 8
        Me.LblGitStashed.Text = "Stashed"
        '
        'LblGitChanged
        '
        Me.LblGitChanged.AutoSize = True
        Me.LblGitChanged.Location = New System.Drawing.Point(23, 109)
        Me.LblGitChanged.Name = "LblGitChanged"
        Me.LblGitChanged.Size = New System.Drawing.Size(50, 13)
        Me.LblGitChanged.TabIndex = 6
        Me.LblGitChanged.Text = "Changed"
        '
        'LblGitNew
        '
        Me.LblGitNew.AutoSize = True
        Me.LblGitNew.Location = New System.Drawing.Point(23, 82)
        Me.LblGitNew.Name = "LblGitNew"
        Me.LblGitNew.Size = New System.Drawing.Size(29, 13)
        Me.LblGitNew.TabIndex = 4
        Me.LblGitNew.Text = "New"
        '
        'PbGitStashed
        '
        Me.PbGitStashed.Location = New System.Drawing.Point(140, 134)
        Me.PbGitStashed.Name = "PbGitStashed"
        Me.PbGitStashed.Size = New System.Drawing.Size(19, 19)
        Me.PbGitStashed.TabIndex = 3
        Me.PbGitStashed.TabStop = False
        '
        'PbGitChanged
        '
        Me.PbGitChanged.Location = New System.Drawing.Point(140, 109)
        Me.PbGitChanged.Name = "PbGitChanged"
        Me.PbGitChanged.Size = New System.Drawing.Size(19, 19)
        Me.PbGitChanged.TabIndex = 3
        Me.PbGitChanged.TabStop = False
        '
        'PBGitNew
        '
        Me.PBGitNew.Location = New System.Drawing.Point(140, 82)
        Me.PBGitNew.Name = "PBGitNew"
        Me.PBGitNew.Size = New System.Drawing.Size(19, 19)
        Me.PBGitNew.TabIndex = 3
        Me.PBGitNew.TabStop = False
        '
        'cmdGitStashed
        '
        Me.cmdGitStashed.Location = New System.Drawing.Point(159, 134)
        Me.cmdGitStashed.Name = "cmdGitStashed"
        Me.cmdGitStashed.Size = New System.Drawing.Size(34, 19)
        Me.cmdGitStashed.TabIndex = 9
        Me.cmdGitStashed.Text = "..."
        Me.cmdGitStashed.UseVisualStyleBackColor = True
        '
        'cmdGitChanged
        '
        Me.cmdGitChanged.Location = New System.Drawing.Point(159, 109)
        Me.cmdGitChanged.Name = "cmdGitChanged"
        Me.cmdGitChanged.Size = New System.Drawing.Size(34, 19)
        Me.cmdGitChanged.TabIndex = 7
        Me.cmdGitChanged.Text = "..."
        Me.cmdGitChanged.UseVisualStyleBackColor = True
        '
        'CmdGitNew
        '
        Me.CmdGitNew.Location = New System.Drawing.Point(159, 82)
        Me.CmdGitNew.Name = "CmdGitNew"
        Me.CmdGitNew.Size = New System.Drawing.Size(34, 19)
        Me.CmdGitNew.TabIndex = 5
        Me.CmdGitNew.Text = "..."
        Me.CmdGitNew.UseVisualStyleBackColor = True
        '
        'CmdGit
        '
        Me.CmdGit.Location = New System.Drawing.Point(632, 21)
        Me.CmdGit.Name = "CmdGit"
        Me.CmdGit.Size = New System.Drawing.Size(34, 19)
        Me.CmdGit.TabIndex = 2
        Me.CmdGit.Text = "..."
        Me.CmdGit.UseVisualStyleBackColor = True
        '
        'TxtGit
        '
        Me.TxtGit.Location = New System.Drawing.Point(137, 21)
        Me.TxtGit.Name = "TxtGit"
        Me.TxtGit.Size = New System.Drawing.Size(489, 20)
        Me.TxtGit.TabIndex = 1
        '
        'LblGit
        '
        Me.LblGit.AutoSize = True
        Me.LblGit.Location = New System.Drawing.Point(19, 24)
        Me.LblGit.Name = "LblGit"
        Me.LblGit.Size = New System.Drawing.Size(94, 13)
        Me.LblGit.TabIndex = 0
        Me.LblGit.Text = "Location of git exe"
        '
        'LblColour
        '
        Me.LblColour.AutoSize = True
        Me.LblColour.Location = New System.Drawing.Point(23, 59)
        Me.LblColour.Name = "LblColour"
        Me.LblColour.Size = New System.Drawing.Size(37, 13)
        Me.LblColour.TabIndex = 3
        Me.LblColour.Text = "Colour"
        '
        'FrmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.GrpGit)
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
        Me.GrpGit.ResumeLayout(False)
        Me.GrpGit.PerformLayout()
        CType(Me.PbGitStashed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PbGitChanged, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PBGitNew, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents GrpGit As Windows.Forms.GroupBox
    Friend WithEvents CmdGit As Windows.Forms.Button
    Friend WithEvents TxtGit As Windows.Forms.TextBox
    Friend WithEvents LblGit As Windows.Forms.Label
    Friend WithEvents LblGitStashed As Windows.Forms.Label
    Friend WithEvents LblGitChanged As Windows.Forms.Label
    Friend WithEvents LblGitNew As Windows.Forms.Label
    Friend WithEvents PbGitStashed As Windows.Forms.PictureBox
    Friend WithEvents PbGitChanged As Windows.Forms.PictureBox
    Friend WithEvents PBGitNew As Windows.Forms.PictureBox
    Friend WithEvents cmdGitStashed As Windows.Forms.Button
    Friend WithEvents cmdGitChanged As Windows.Forms.Button
    Friend WithEvents CmdGitNew As Windows.Forms.Button
    Friend WithEvents LblColour As Windows.Forms.Label
End Class
