Imports Microsoft.Vbe.Interop

Public Class LineNumbering
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
End Class
