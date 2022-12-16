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
    Private WithEvents _MySettings As CommandBarButton = Nothing
    Private WithEvents _MyIndentation As CommandBarButton = Nothing

    Private ClsIndent As New Indentation
    Private ClsVBEHandling As New VBEHandling
    Private ClsLineNumbering As New LineNumbering
    Private ClsCodeModuleHandling As New CodeModuleHandling

    Public Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        Try
            _VBE = DirectCast(Application, VBE)
            _AddIn = DirectCast(AddInInst, AddIn)
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
            _MySettings = .Controls.Add(MsoControlType.msoControlButton)
            With _MySettings
                .Caption = inoVBETools.My.Resources.menuSettings
                .BeginGroup = True
            End With
        End With
    End Sub

    Private Sub _MyLineNummeringButton1_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton1.Click
        ClsLineNumbering.AddLineNumbersToComponent(_VBE.ActiveCodePane.CodeModule)
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

    Private Sub _MyLineNummeringButton2_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton2.Click
        ClsLineNumbering.AddLineNumbersToComponent(_VBE.ActiveCodePane.CodeModule, True)
    End Sub

    Private Sub _MySettings_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MySettings.Click
        Dim frm As New FrmOptions
        frm.Show()
    End Sub

    Private Sub SetLanguage()
        My.Application.ChangeUICulture(My.Settings.Language)
    End Sub

    Private Sub _MyIndentation_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyIndentation.Click
        Dim StartPos As Long = 0
        Dim Countlines As Long = 0
        Dim strCode As String = ClsIndent.IndentCode(ClsVBEHandling.GetCurrentProcedureCode(_VBE, StartPos, Countlines))

        _VBE.ActiveCodePane.CodeModule.DeleteLines(StartPos, Countlines)
        _VBE.ActiveCodePane.CodeModule.InsertLines(StartPos, strCode)
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


End Class
