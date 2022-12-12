Imports Microsoft
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop

Public Class VBEHandling
    Public Function GetCurrentProcedureCode(_VBE As VBE, ByRef StartPos As Long, ByRef CountLines As Long, Optional blnIncHeader As Boolean = False) As String
        Dim startline As Long
        Dim startcol As Long
        Dim endline As Long
        Dim endcol As Long


        Dim CodeMod As CodeModule
        CodeMod = _VBE.ActiveCodePane.CodeModule

        _VBE.ActiveCodePane.GetSelection(startline, startcol, endline, endcol)
        Dim strProc As String = CodeMod.ProcOfLine(startline, vbext_ProcKind.vbext_pk_Proc)

        Dim lProcStart As Long = CodeMod.ProcStartLine(strProc, vbext_ProcKind.vbext_pk_Proc)
        Dim lProcBodyStart As Long = CodeMod.ProcBodyLine(strProc, vbext_ProcKind.vbext_pk_Proc)
        CountLines = CodeMod.ProcCountLines(strProc, vbext_ProcKind.vbext_pk_Proc)



        If blnIncHeader = True Then
            StartPos = lProcStart
            Return CodeMod.Lines(StartPos, CountLines)
        Else
            StartPos = lProcBodyStart + 1
            CountLines = CountLines - (lProcBodyStart - lProcStart) - 2
            Return CodeMod.Lines(StartPos, CountLines)
        End If



    End Function

    Public Function GetFnOrSubNameOfCurrentPosition(_VBE As VBE) As String

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

    Public Function GetFnOrSubTypeCurrentPosition(_VBE As VBE) As String

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
End Class
