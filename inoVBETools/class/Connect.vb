Imports Microsoft.Office.Interop
Imports Extensibility
'Imports Microsoft.Office.Interop.Access
'Imports Microsoft.Office.Interop.Access.Dao
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Vbe.Interop
Imports Microsoft
Imports Microsoft.Office.Core
Imports System.Runtime.ConstrainedExecution
Imports System.Reflection
Imports System.IO
'Imports Microsoft.Vbe.Interop.Forms

<ComVisible(True), Guid("1B3515B2-6A73-40C8-9DA4-1766ED6600ED"), ProgId("inoVBETools.Connect")>
Public Class Connect
    Implements Extensibility.IDTExtensibility2

    Private _VBE As VBE
    Private _AddIn As AddIn

    Private WithEvents _MyLineNummeringButton1 As CommandBarButton = Nothing
    Private WithEvents _MyLineNummeringButton2 As CommandBarButton = Nothing
    Private WithEvents _MyLineNummeringButton3 As CommandBarButton = Nothing
    Private WithEvents _MyLineNummeringButton4 As CommandBarButton = Nothing
    Private WithEvents _MyErrorHandling1 As CommandBarButton = Nothing
    Private WithEvents _MyErrorHandling2 As CommandBarButton = Nothing
    Private WithEvents _MyExport As CommandBarButton = Nothing
    Private WithEvents _MyGitExport As CommandBarButton = Nothing
    Private WithEvents _MyImport As CommandBarButton = Nothing
    Private WithEvents _MySettings As CommandBarButton = Nothing
    Private WithEvents _MyIndentation As CommandBarButton = Nothing

    Private ClsIndent As New Indentation
    Private ClsVBEHandling As New VBEHandling
    Private ClsLineNumbering As New LineNumbering
    Private ClsCodeModuleHandling As New CodeModuleHandling

    Public HostApplicationName As String

    Public Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        Try
            _VBE = DirectCast(Application, VBE)
            _AddIn = DirectCast(AddInInst, AddIn)

            For Each refVBE As Reference In _VBE.ActiveVBProject.References
                With refVBE
                    If .BuiltIn = True And .Name <> "VBA" Then
                        HostApplicationName = .Name
                        Exit For
                    End If
                End With
            Next refVBE

        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        'Throw New NotImplementedException()
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
        ' Throw New NotImplementedException()
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        'MessageBox.Show("Add-In geladen (OnStartupComplete): " & _AddIn.ProgId)
        InitializeAddIn()
        ' Throw New NotImplementedException()
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        ' Throw New NotImplementedException()
    End Sub

    Private Sub InitializeAddIn()
        Dim cbr As CommandBar
        Dim cbrAddIns As CommandBarPopup = Nothing
        Dim cbrSub As CommandBarPopup = Nothing
        Dim cbrExport As CommandBarPopup = Nothing
        SetLanguage()

        cbr = _VBE.CommandBars("Menüleiste")
        cbrSub = cbr.Controls.Add(MsoControlType.msoControlPopup, Before:=10)
        With cbrSub
            .Caption = "inoVBETools"
            '.BeginGroup = True
            _MyLineNummeringButton1 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyLineNummeringButton1
                .Caption = inoVBETools.My.Resources.menuLineNumber1
            End With
            _MyLineNummeringButton2 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyLineNummeringButton2
                .Caption = inoVBETools.My.Resources.menuLineNumber2
            End With
            _MyLineNummeringButton3 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyLineNummeringButton3
                .Caption = inoVBETools.My.Resources.menuLineNumber3
            End With
            _MyLineNummeringButton4 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyLineNummeringButton4
                .Caption = inoVBETools.My.Resources.menuLineNumber4
            End With
            _MyErrorHandling1 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyErrorHandling1
                .Caption = inoVBETools.My.Resources.menuErrorHandling
                .BeginGroup = True
            End With
            _MyErrorHandling2 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyErrorHandling2
                .Caption = inoVBETools.My.Resources.menuErrorHandlingDebug
            End With
            _MyIndentation = .Controls.Add(MsoControlType.msoControlButton)
            With _MyIndentation
                .Caption = inoVBETools.My.Resources.menuIndentationAll
                .BeginGroup = True
            End With

            cbrExport = .Controls.Add(MsoControlType.msoControlPopup)
            With cbrExport
                .Caption = inoVBETools.My.Resources.menuCodeExportImport
                .BeginGroup = True
            End With
            _MyExport = cbrExport.Controls.Add(MsoControlType.msoControlButton)
            With _MyExport
                .Caption = inoVBETools.My.Resources.menuExportCode
                .BeginGroup = True
            End With
            _MyGitExport = cbrExport.Controls.Add(MsoControlType.msoControlButton)
            With _MyGitExport
                .Caption = inoVBETools.My.Resources.menuGitExport
            End With
            _MyImport = cbrExport.Controls.Add(MsoControlType.msoControlButton)
            With _MyImport
                .Caption = inoVBETools.My.Resources.menuImport
                .BeginGroup = True
            End With
            _MySettings = .Controls.Add(MsoControlType.msoControlButton)
            With _MySettings
                .Caption = inoVBETools.My.Resources.menuSettings
                .BeginGroup = True
            End With
        End With
    End Sub

    Private Sub _MyErrorHandling1_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyErrorHandling1.Click
        Dim startline As Long
        Dim startcol As Long
        Dim endline As Long
        Dim endcol As Long

        _VBE.ActiveCodePane.GetSelection(startline, startcol, endline, endcol)

        Dim strVBA As New System.Text.StringBuilder

        strVBA.Append("    On Error Goto " & My.Settings.GotoError)
        strVBA.Append(vbCrLf & vbCrLf & vbCrLf)
        strVBA.Append("    Exit " & ClsVBEHandling.GetFnOrSubTypeCurrentPosition(_VBE) & vbCrLf)
        strVBA.Append(My.Settings.GotoError & ":" & vbCrLf)
        strVBA.Append("    Select Case Err.Number" & vbCrLf)
        strVBA.Append("        Case Else" & vbCrLf)
        Dim strErr As String = String.Format("            MsgBox ""{0} "" & Erl & "" {1} '{2}'"" & vbNewLine _", inoVBETools.My.Resources.ErrorInLine, inoVBETools.My.Resources.ErrorInProcedure, ClsVBEHandling.GetFnOrSubNameOfCurrentPosition(_VBE))
        strVBA.Append(strErr & vbCrLf)
        strVBA.Append("                 & Err.Number & "" - "" & Err.Description" & vbCrLf)
        strVBA.Append("    End Select")

        _VBE.ActiveCodePane.CodeModule.InsertLines(startline + 1, strVBA.ToString)

    End Sub

    Private Sub _MyErrorHandling2_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyErrorHandling2.Click
        Dim startline As Long
        Dim startcol As Long
        Dim endline As Long
        Dim endcol As Long

        _VBE.ActiveCodePane.GetSelection(startline, startcol, endline, endcol)

        Dim strVBA As New System.Text.StringBuilder

        strVBA.Append("    On Error Goto " & My.Settings.GotoError)
        strVBA.Append(vbCrLf & vbCrLf & vbCrLf)
        strVBA.Append("    Exit " & ClsVBEHandling.GetFnOrSubTypeCurrentPosition(_VBE) & vbCrLf)
        strVBA.Append(My.Settings.GotoError & ":" & vbCrLf)
        strVBA.Append("    Select Case Err.Number" & vbCrLf)
        strVBA.Append("        Case Else" & vbCrLf)
        strVBA.Append("            Dim errMsg as String" & vbCrLf)
        Dim strErr As String = String.Format("            errMsg = ""{0} "" & Erl & "" {1} '{2}'"" & vbNewLine _", inoVBETools.My.Resources.ErrorInLine, inoVBETools.My.Resources.ErrorInProcedure, ClsVBEHandling.GetFnOrSubNameOfCurrentPosition(_VBE))
        strVBA.Append(strErr & vbCrLf)
        strVBA.Append("                 & Err.Number & "" - "" & Err.Description" & vbCrLf)
        strVBA.Append("            Select Case frmInoVBEError.ShowForm(errMsg)" & vbCrLf)
        strVBA.Append("                 Case 1" & vbCrLf & vbCrLf)
        strVBA.Append("                 Case 2" & vbCrLf)
        strVBA.Append("                     Debug.Print errMsg" & vbCrLf)
        strVBA.Append("                     Debug.Assert False" & vbCrLf)
        strVBA.Append("            End Select" & vbCrLf)
        strVBA.Append("    End Select")

        _VBE.ActiveCodePane.CodeModule.InsertLines(startline + 1, strVBA.ToString)

        Dim strErrorForm As String = Path.Combine(My.Application.Info.DirectoryPath, "ressources\vbafiles\frmInoVBEError.frm")
        If ClsCodeModuleHandling.ModuleExists(strErrorForm, _VBE.ActiveVBProject) = False Then
            ClsCodeModuleHandling.ImportCodeModule(_VBE.ActiveVBProject, strErrorForm)
        End If
    End Sub

    Private Sub _MyLineNummeringButton1_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton1.Click
        ClsLineNumbering.AddLineNumbersToComponent(_VBE.ActiveCodePane.CodeModule)
    End Sub

    Private Sub _MyLineNummeringButton2_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton2.Click
        ClsLineNumbering.AddLineNumbersToComponent(_VBE.ActiveCodePane.CodeModule, True)
    End Sub

    Private Sub _MyLineNummeringButton3_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton3.Click
        Dim StartPos As Long = 0
        Dim Countlines As Long = 0
        Dim strCode As String = ClsLineNumbering.AddLineNumbersToCurrentProcedure(ClsVBEHandling.GetCurrentProcedureCode(_VBE, StartPos, Countlines))

        _VBE.ActiveCodePane.CodeModule.DeleteLines(StartPos, Countlines)
        _VBE.ActiveCodePane.CodeModule.InsertLines(StartPos, strCode)
    End Sub

    Private Sub _MyLineNummeringButton4_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton4.Click
        Dim StartPos As Long = 0
        Dim Countlines As Long = 0
        Dim strCode As String = ClsLineNumbering.AddLineNumbersToCurrentProcedure(ClsVBEHandling.GetCurrentProcedureCode(_VBE, StartPos, Countlines), True)

        _VBE.ActiveCodePane.CodeModule.DeleteLines(StartPos, Countlines)
        _VBE.ActiveCodePane.CodeModule.InsertLines(StartPos, strCode)
    End Sub

    Private Sub _MySettings_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MySettings.Click
        Dim frm As New FrmOptions
        frm.Show()
    End Sub

    Private Sub _MyIndentation_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyIndentation.Click
        Dim StartPos As Long = 0
        Dim Countlines As Long = 0
        Dim strCode As String = ClsIndent.IndentCode(ClsVBEHandling.GetCurrentProcedureCode(_VBE, StartPos, Countlines))

        _VBE.ActiveCodePane.CodeModule.DeleteLines(StartPos, Countlines)
        _VBE.ActiveCodePane.CodeModule.InsertLines(StartPos, strCode)
    End Sub

    Private Sub _MyExport_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyExport.Click
        If ClsCodeModuleHandling.CheckProjectHasName(_VBE.ActiveVBProject) = False Then
            Exit Sub
        End If

        Dim cd As String = SelectCodeDirectory()
        If cd <> "" Then
            ClsCodeModuleHandling.ExportModules(_VBE.ActiveVBProject, cd & "\")
            My.Settings.WorkingDirectory = cd
            My.Settings.Save()
            If MessageBox.Show(My.Resources.msgGitDirect, "inoVBETools", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                OpenGitForm()
            End If
        End If
    End Sub

    Private Sub _MyImport_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyImport.Click
        If ClsCodeModuleHandling.CheckProjectHasName(_VBE.ActiveVBProject) = False Then
            Exit Sub
        End If

        Dim cd As String = SelectCodeDirectory()
        If cd <> "" Then
            ClsCodeModuleHandling.ImportModules(_VBE.ActiveVBProject, cd & "\")
        End If
    End Sub

    Private Sub _MyGitExport_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyGitExport.Click
        OpenGitForm()
    End Sub

    Private Sub SetLanguage()
        My.Application.ChangeUICulture(My.Settings.Language)
    End Sub

    Private Function SelectCodeDirectory() As String
        Dim startPath As String = ClsVBEHandling.ProjectDirectoryByName(_VBE.ActiveVBProject.Name)
        If startPath = "" Then
            startPath = My.Settings.LastExportFolder
        End If
        Dim ofd As New OpenFileDialog
        With ofd
            .ValidateNames = False
            .CheckFileExists = False
            .CheckPathExists = True
            .InitialDirectory = startPath
            .Multiselect = False
            .Title = String.Format(inoVBETools.My.Resources.ConnectExportTitle, inoVBETools.My.Resources.ConnectTemporaryFileName)
            .FileName = inoVBETools.My.Resources.ConnectTemporaryFileName
            If .ShowDialog = DialogResult.OK Then
                My.Settings.LastExportFolder = Path.GetDirectoryName(.FileName)
                My.Settings.Save()
                ClsVBEHandling.ProjectAdd(_VBE.ActiveVBProject.Name, My.Settings.LastExportFolder)
                ClsVBEHandling.WriteProjectEntries()
                Return My.Settings.LastExportFolder
            End If
        End With
        Return ""
    End Function

    Private Sub OpenGitForm()
        If Not System.IO.File.Exists(My.Settings.Git_Exe) Then
            MessageBox.Show(My.Resources.msgMissingGit)
            Exit Sub
        End If
        Dim ClsGit As New GitHandling(My.Settings.Git_Exe)
        If Not ClsGit.IsDirectoryRepo(My.Settings.WorkingDirectory) Then
            ClsGit.InitializeRepo(My.Settings.WorkingDirectory)
        End If
        Dim frm As New FrmGit
        frm.Show()
    End Sub
End Class
