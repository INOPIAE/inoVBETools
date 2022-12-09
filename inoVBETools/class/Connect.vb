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


        cbr = _VBE.CommandBars("Menüleiste")
        '  cbrAddIns = cbr.Controls.Item("Add-&Ins")
        cbrSub = cbr.Controls.Add(MsoControlType.msoControlPopup)
        With cbrSub
            .Caption = "inoVBETools"
            '.BeginGroup = True
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
                .BeginGroup = True
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

        With vbaCodeModule
            For intLine = .CountOfDeclarationLines + 1 To .CountOfLines
                If .Lines(intLine, 1).Trim <> vbNullString And If(.Lines(intLine, 1).Trim <> vbNullString, .Lines(intLine, 1).Trim.First <> "'", False) Then
                    If .ProcOfLine(intLine, 0) <> strModulname Then
                        strModulname = .ProcOfLine(intLine, 4)
                        If blnEachProcedure = True Then
                            intLineCounter = 0
                        End If
                        If .Lines(intLine, 1).Trim.Last = "_" Then
                            bolUnderscore = True
                        Else
                            bolUnderscore = False
                        End If
                    Else
                        If "End Sub End Function End Property".Contains(.Lines(intLine, 1)) = False Then
                            If Not bolUnderscore And Not bolSelect Then
                                If .Lines(intLine, 1).Trim.Last = "_" Then bolUnderscore = True
                                If .Lines(intLine, 1).Contains("Select Case") Then bolSelect = True
                                If IsNumeric(.Lines(intLine, 1).Substring(0, 1)) Then
                                    For intColumn = 1 To .Lines(intLine, 1).Length
                                        If Not IsNumeric(.Lines(intLine, 1).Substring(0, intColumn)) Then
                                            Exit For
                                        End If
                                    Next
                                    .ReplaceLine(intLine, StrDup(intColumn, " ") & .Lines(intLine, 1).Substring(intColumn - 1))
                                End If
                                intLineCounter += 1
                                If blnNoNumber = False Then
                                    If .Lines(intLine, 1).Length > 4 Then
                                        For intColumn = 0 To 3
                                            If .Lines(intLine, 1).Substring(intColumn, 1) <> " " Then
                                                Exit For
                                            End If
                                        Next
                                        .ReplaceLine(intLine, intLineCounter.ToString.PadRight(4) & .Lines(intLine, 1).Substring(intColumn))
                                    Else
                                        .ReplaceLine(intLine, intLineCounter.ToString.PadRight(4) & .Lines(intLine, 1).Trim)
                                    End If

                                    lngCount += 1
                                End If
                            Else
                                If .Lines(intLine, 1).Trim.Last <> "_" Then bolUnderscore = False
                                If .Lines(intLine, 1).Contains("Case") Then bolSelect = False
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
