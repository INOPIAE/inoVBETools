<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmGit
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
        Me.cmdTest = New System.Windows.Forms.Button()
        Me.TxtCommit = New System.Windows.Forms.TextBox()
        Me.lblCommit = New System.Windows.Forms.Label()
        Me.TvGit = New System.Windows.Forms.TreeView()
        Me.CmdAdd = New System.Windows.Forms.Button()
        Me.CmdRemove = New System.Windows.Forms.Button()
        Me.CmdCommit = New System.Windows.Forms.Button()
        Me.CmdOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdTest
        '
        Me.cmdTest.Location = New System.Drawing.Point(650, 49)
        Me.cmdTest.Name = "cmdTest"
        Me.cmdTest.Size = New System.Drawing.Size(75, 23)
        Me.cmdTest.TabIndex = 0
        Me.cmdTest.Text = "Button1"
        Me.cmdTest.UseVisualStyleBackColor = True
        '
        'TxtCommit
        '
        Me.TxtCommit.Location = New System.Drawing.Point(146, 49)
        Me.TxtCommit.Multiline = True
        Me.TxtCommit.Name = "TxtCommit"
        Me.TxtCommit.Size = New System.Drawing.Size(289, 52)
        Me.TxtCommit.TabIndex = 1
        '
        'lblCommit
        '
        Me.lblCommit.AutoSize = True
        Me.lblCommit.Location = New System.Drawing.Point(32, 54)
        Me.lblCommit.Name = "lblCommit"
        Me.lblCommit.Size = New System.Drawing.Size(86, 13)
        Me.lblCommit.TabIndex = 3
        Me.lblCommit.Text = "Commit message"
        '
        'TvGit
        '
        Me.TvGit.Location = New System.Drawing.Point(35, 126)
        Me.TvGit.Name = "TvGit"
        Me.TvGit.Size = New System.Drawing.Size(400, 244)
        Me.TvGit.TabIndex = 5
        '
        'CmdAdd
        '
        Me.CmdAdd.Location = New System.Drawing.Point(502, 126)
        Me.CmdAdd.Name = "CmdAdd"
        Me.CmdAdd.Size = New System.Drawing.Size(130, 44)
        Me.CmdAdd.TabIndex = 0
        Me.CmdAdd.Text = "Add"
        Me.CmdAdd.UseVisualStyleBackColor = True
        '
        'CmdRemove
        '
        Me.CmdRemove.Location = New System.Drawing.Point(502, 176)
        Me.CmdRemove.Name = "CmdRemove"
        Me.CmdRemove.Size = New System.Drawing.Size(130, 44)
        Me.CmdRemove.TabIndex = 0
        Me.CmdRemove.Text = "Remove"
        Me.CmdRemove.UseVisualStyleBackColor = True
        '
        'CmdCommit
        '
        Me.CmdCommit.Location = New System.Drawing.Point(502, 347)
        Me.CmdCommit.Name = "CmdCommit"
        Me.CmdCommit.Size = New System.Drawing.Size(130, 44)
        Me.CmdCommit.TabIndex = 0
        Me.CmdCommit.Text = "Commit"
        Me.CmdCommit.UseVisualStyleBackColor = True
        '
        'CmdOK
        '
        Me.CmdOK.Location = New System.Drawing.Point(656, 347)
        Me.CmdOK.Name = "CmdOK"
        Me.CmdOK.Size = New System.Drawing.Size(130, 44)
        Me.CmdOK.TabIndex = 6
        Me.CmdOK.Text = "OK"
        Me.CmdOK.UseVisualStyleBackColor = True
        '
        'FrmGit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.CmdOK)
        Me.Controls.Add(Me.TvGit)
        Me.Controls.Add(Me.lblCommit)
        Me.Controls.Add(Me.TxtCommit)
        Me.Controls.Add(Me.CmdCommit)
        Me.Controls.Add(Me.CmdRemove)
        Me.Controls.Add(Me.CmdAdd)
        Me.Controls.Add(Me.cmdTest)
        Me.Name = "FrmGit"
        Me.Text = "FrmGit"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdTest As Windows.Forms.Button
    Friend WithEvents TxtCommit As Windows.Forms.TextBox
    Friend WithEvents lblCommit As Windows.Forms.Label
    Friend WithEvents TvGit As Windows.Forms.TreeView
    Friend WithEvents CmdAdd As Windows.Forms.Button
    Friend WithEvents CmdRemove As Windows.Forms.Button
    Friend WithEvents CmdCommit As Windows.Forms.Button
    Friend WithEvents CmdOK As Windows.Forms.Button
End Class
