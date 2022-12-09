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
'Imports Microsoft.Vbe.Interop.Forms

<ComVisible(True), Guid("1B3515B2-6A73-40C8-9DA4-1766ED6600ED"), ProgId("inoVBETools.Connect")>
Public Class Connect
    Implements Extensibility.IDTExtensibility2

    Private _VBE As VBE
    Private _AddIn As AddIn
    'CommandBar
    Private WithEvents _myStandardCommandBarButton As CommandBarButton
    Private WithEvents _myToolsCommandBarButton As CommandBarButton
    Private WithEvents _myCodeWindowCommandBarButton As CommandBarButton
    Private WithEvents _myToolBarButton As CommandBarButton
    Private WithEvents _myCommandBarPopup1Button As CommandBarButton
    Private WithEvents _myCommandBarPopup2Button As CommandBarButton
    ' CommandBars created by the add-in
    Private _myToolbar As CommandBar
    Private _myCommandBarPopup1 As CommandBarPopup
    Private _myCommandBarPopup2 As CommandBarPopup

    Private WithEvents _MyLineNummeringButton1 As CommandBarButton = Nothing
    Private WithEvents _MyLineNummeringButton2 As CommandBarButton = Nothing
    Private WithEvents _MyErrorHandling As CommandBarButton = Nothing

    Public Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        Try
            _VBE = DirectCast(Application, VBE)
            _AddIn = DirectCast(AddInInst, AddIn)
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        Try
            Select Case RemoveMode
                Case ext_DisconnectMode.ext_dm_HostShutdown, ext_DisconnectMode.ext_dm_UserClosed
                    ' Delete buttons on built-in commandbars
                    If Not (_myStandardCommandBarButton Is Nothing) Then
                        _myStandardCommandBarButton.Delete()
                    End If
                    If Not (_myCodeWindowCommandBarButton Is Nothing) Then
                        _myCodeWindowCommandBarButton.Delete()
                    End If
                    If Not (_myToolsCommandBarButton Is Nothing) Then
                        _myToolsCommandBarButton.Delete()
                    End If
                    ' Disconnect event handlers
                    _myToolBarButton = Nothing
                    _myCommandBarPopup1Button = Nothing
                    _myCommandBarPopup2Button = Nothing
                    ' Delete commandbars created by the add-in
                    If Not (_myToolbar Is Nothing) Then
                        _myToolbar.Delete()
                    End If
                    If Not (_myCommandBarPopup1 Is Nothing) Then
                        _myCommandBarPopup1.Delete()
                    End If
                    If Not (_myCommandBarPopup2 Is Nothing) Then
                        _myCommandBarPopup2.Delete()
                    End If
            End Select
        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show(e.ToString)
        End Try

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

    Private Function AddCommandBarButton(ByVal commandBar As CommandBar) As CommandBarButton
        Dim commandBarButton As CommandBarButton
        Dim commandBarControl As CommandBarControl
        commandBarControl = commandBar.Controls.Add(MsoControlType.msoControlButton)
        commandBarButton = DirectCast(commandBarControl, CommandBarButton)
        commandBarButton.Caption = "My button"
        commandBarButton.FaceId = 59
        Return commandBarButton
    End Function
    Private Sub InitializeAddIn()
        ' Constants for names of built-in commandbars of the VBA editor
        Const STANDARD_COMMANDBAR_NAME As String = "Voreinstellung"  '"Standard"
        Const MENUBAR_COMMANDBAR_NAME As String = "Menüleiste" ' "Menu Bar"
        Const TOOLS_COMMANDBAR_NAME As String = "Tools"
        Const CODE_WINDOW_COMMANDBAR_NAME As String = "Code Window"
        ' Constants for names of commandbars created by the add-in
        Const MY_COMMANDBAR_POPUP1_NAME As String = "MyTemporaryCommandBarPopup1"
        Const MY_COMMANDBAR_POPUP2_NAME As String = "MyTemporaryCommandBarPopup2"
        ' Constants for captions of commandbars created by the add-in
        Const MY_COMMANDBAR_POPUP1_CAPTION As String = "My sub menu"
        Const MY_COMMANDBAR_POPUP2_CAPTION As String = "My main menu"
        Const MY_TOOLBAR_CAPTION As String = "My toolbar"
        ' Built-in commandbars of the VBA editor
        Dim standardCommandBar As CommandBar
        Dim menuCommandBar As CommandBar
        Dim toolsCommandBar As CommandBar
        Dim codeCommandBar As CommandBar
        ' Other variables
        Dim toolsCommandBarControl As CommandBarControl
        Dim position As Integer
        Try
            ' Retrieve some built-in commandbars
            standardCommandBar = _VBE.CommandBars.Item(STANDARD_COMMANDBAR_NAME)
            menuCommandBar = _VBE.CommandBars.Item(MENUBAR_COMMANDBAR_NAME)
            'toolsCommandBar = _VBE.CommandBars.Item(TOOLS_COMMANDBAR_NAME)
            codeCommandBar = _VBE.CommandBars.Item(CODE_WINDOW_COMMANDBAR_NAME)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString)
        End Try



        Dim cbr As CommandBar
        Dim cbrAddIns As CommandBarPopup = Nothing
        Dim cbrSub As CommandBarPopup = Nothing


        cbr = _VBE.CommandBars("Menüleiste")
        cbrAddIns = cbr.Controls.Item("Add-&Ins")
        cbrSub = cbrAddIns.Controls.Add(MsoControlType.msoControlPopup)
        With cbrSub
            .Caption = "inoVBETools"
            .BeginGroup = True
            _MyLineNummeringButton1 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyLineNummeringButton1
                .Caption = "Zeilennummerierung"
            End With
            _MyLineNummeringButton2 = .Controls.Add(MsoControlType.msoControlButton)
            With _MyLineNummeringButton2
                .Caption = "Zeilennummerierung entfernen"
            End With
            _MyErrorHandling = .Controls.Add(MsoControlType.msoControlButton)
            With _MyErrorHandling
                .Caption = "Fehlerbehandlung"
            End With
        End With
    End Sub

    Private Sub _MyLineNummeringButton1_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton1.Click
        AddLineNumbersToComponent(_VBE.ActiveCodePane.CodeModule)
    End Sub

    Private Sub _MyErrorHandling_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyErrorHandling.Click
        Dim startline As Long
        Dim startcol As Long
        Dim endline As Long
        Dim endcol As Long

        _VBE.ActiveCodePane.GetSelection(startline, startcol, endline, endcol)

        Dim strVBA As String = "    On Error Goto ErrHandling" & vbNewLine _
            & vbNewLine & vbNewLine _
            & "    Exit " & GetFnOrSubTypeCurrentPosition() & vbNewLine _
            & "ErrHandling:" & vbNewLine _
            & "    Select Case Err.Number" & vbNewLine _
            & "        Case Else" & vbNewLine _
            & "            MsgBox ""Fehler In Zeile "" & Erl & "" in der Routine '" & GetFnOrSubNameOfCurrentPosition() & "'"" & vbNewLine _" & vbNewLine _
            & "                 & Err.Number & "" - "" & Err.Description" & vbNewLine _
            & "    End Select"
        _VBE.ActiveCodePane.CodeModule.InsertLines(startline + 1, strVBA)

    End Sub

    Private Function GetFnOrSubNameOfCurrentPosition() As String

        Dim CodeMod As CodeModule

        Dim startline As Long
        Dim startcol As Long
        Dim endline As Long
        Dim endcol As Long

        _VBE.ActiveCodePane.GetSelection(startline, startcol, endline, endcol)

        CodeMod = _VBE.ActiveCodePane.CodeModule

        For intC As Int16 = startline To 1 Step -1
            Dim strTest As String = CodeMod.Lines(intC, 1)
            Dim strTestA() As String
            Dim strTestA1() As String
            If strTest.Contains("Sub") Then
                strTestA = strTest.Split("(")
                strTestA1 = strTestA(0).Split(" ")
                Return strTestA1.Last
            End If
            If strTest.Contains("Function") Then
                strTestA = strTest.Split("(")
                strTestA1 = strTestA(0).Split(" ")
                Return strTestA1.Last
            End If
        Next

        Return String.Empty
    End Function

    Private Function GetFnOrSubTypeCurrentPosition() As String

        Dim CodeMod As CodeModule

        Dim startline As Long
        Dim startcol As Long
        Dim endline As Long
        Dim endcol As Long

        _VBE.ActiveCodePane.GetSelection(startline, startcol, endline, endcol)

        CodeMod = _VBE.ActiveCodePane.CodeModule

        For intC As Int16 = startline To 1 Step -1
            Dim strTest As String = CodeMod.Lines(intC, 1)
            If strTest.Contains("Sub") Then
                Return "Sub"
            End If
            If strTest.Contains("Function") Then
                Return "Function"
            End If
        Next

        Return String.Empty
    End Function

    Public Function AddLineNumbersToComponent(vbaCodeModule As CodeModule, Optional blnNoNumber As Boolean = False, Optional blnEachProcedure As Boolean = True) As Long
        ' returns total line numbers added to code of a single code object as passed to the function
        Dim intLine As Integer
        Dim intColumn As Integer, intLineCounter As Integer
        Dim strModulname As String = vbNullString
        Dim bolUnderscore As Boolean, bolSelect As Boolean
        Dim lngCount As Long

        With vbaCodeModule 'vbaComponent.CodeModule
            For intLine = .CountOfDeclarationLines + 1 To .CountOfLines
                If .Lines(intLine, 1).Trim <> vbNullString And Left$(Trim$(.Lines(intLine, 1)), 1) <> "'" Then '.Lines(intLine, 1).Trim.Substring(0, 1) <> "'" Then
                    If .ProcOfLine(intLine, 0) <> strModulname Then
                        strModulname = .ProcOfLine(intLine, 4)
                        If blnEachProcedure = True Then
                            intLineCounter = 0
                        End If
                        If Left$(Trim$(StrReverse(.Lines(intLine, 1))), 1) = "_" Then
                            bolUnderscore = True
                        Else
                            bolUnderscore = False
                        End If
                    Else
                        If InStr(1, "End Sub End Function End Property", .Lines(intLine, 1)) = 0 Then
                            If Not bolUnderscore And Not bolSelect Then

                                If Left$(Trim$(StrReverse(.Lines(intLine, 1))), 1) = "_" Then bolUnderscore = True
                                If InStr(1, .Lines(intLine, 1), "Select Case") <> 0 Then bolSelect = True
                                If IsNumeric(Left$(.Lines(intLine, 1), 1)) Then
                                    For intColumn = 1 To Len(.Lines(intLine, 1))
                                        If Not IsNumeric(Left$(.Lines(intLine, 1), intColumn)) Then
                                            Exit For
                                        End If
                                    Next
                                    .ReplaceLine(intLine, StrDup(intColumn - 1, " ") & Mid$(.Lines(intLine, 1), intColumn))
                                End If
                                intLineCounter = intLineCounter + 1
                                If blnNoNumber = False Then
                                    If Trim$(Left$(.Lines(intLine, 1), Len(Trim(intLineCounter)) + 2)) = "" Then
                                        .ReplaceLine(intLine, Mid$(.Lines(intLine, 1), Len(Trim(intLineCounter)) + 2))
                                    Else
                                        .ReplaceLine(intLine, Trim$(.Lines(intLine, 1)))
                                    End If
                                    .ReplaceLine(intLine, Trim$(CStr(intLineCounter)) & " " & .Lines(intLine, 1))
                                    lngCount = lngCount + 1
                                End If
                            Else
                                If Left$(Trim$(StrReverse(.Lines(intLine, 1))), 1) <> "_" Then bolUnderscore = False
                                If InStr(1, .Lines(intLine, 1), "Case") <> 0 Then bolSelect = False
                            End If
                        Else
                            strModulname = vbNullString
                        End If
                    End If
                End If
            Next
        End With
        AddLineNumbersToComponent = lngCount
    End Function

    Private Sub _MyLineNummeringButton2_Click(Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _MyLineNummeringButton2.Click
        AddLineNumbersToComponent(_VBE.ActiveCodePane.CodeModule, True)
    End Sub
End Class
