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
                Dim strFill As String = .Lines(intLine, 1).Trim & " -"
                If .Lines(intLine, 1).Trim <> vbNullString And IIf(.Lines(intLine, 1).Trim <> vbNullString, strFill.First <> "'", False) Then
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
        Return lngCount
    End Function

    Public Function AddLineNumbersToCurrentProcedure(CodeString As String, Optional blnNoNumber As Boolean = False) As String
        Dim intColumn As Integer, intLineCounter As Integer
        Dim bolUnderscore As Boolean
        Dim bolSelect As Boolean
        Dim bolSelectCase As Boolean
        Dim lngCount As Long

        Dim lines() As String = CodeString.Split({ControlChars.CrLf}, StringSplitOptions.None)
        Dim strTest As String = ""
        Dim strReturn As String = ""

        For i As Int16 = 0 To lines.Count - 1

            Dim strFill As String = lines(i).Trim & " -"
            If String.IsNullOrEmpty(lines(i).Trim) = False And IIf(String.IsNullOrEmpty(lines(i).Trim) = False, strFill.First <> "'", False) Then

                If lines(i).Trim.Last = "_" Then
                    bolUnderscore = True
                Else
                    bolUnderscore = False
                End If

                'If Not bolUnderscore And Not bolSelect Then
                If Not bolUnderscore Then
                    If lines(i).Trim.Last = "_" Then bolUnderscore = True
                    If lines(i).Contains("Select Case") Then bolSelect = True

                    If lines(i).Contains("Case") Then
                        bolSelectCase = True
                    Else
                        bolSelectCase = False
                    End If
                    If lines(i).Contains("End Case") Then bolSelectCase = False
                    If lines(i).Contains("Select Case") Then bolSelectCase = False

                    If IsNumeric(lines(i).Substring(0, 1)) Then
                        For intColumn = 0 To lines(i).Length - 1
                            If Not IsNumeric(lines(i).Substring(intColumn, 1)) Then
                                Exit For
                            End If
                        Next
                        strTest = StrDup(intColumn, " ") & lines(i).Substring(intColumn)
                    End If

                    If blnNoNumber = False And bolSelectCase = False Then
                        intLineCounter += 1
                        If lines(i).Length > 4 Then
                            For intColumn = 0 To 3
                                If lines(i).Substring(intColumn, 1) <> " " Then
                                    Exit For
                                End If
                            Next
                            strTest = intLineCounter.ToString.PadRight(4) & lines(i).Substring(intColumn).TrimEnd
                        Else
                            strTest = intLineCounter.ToString.PadRight(4) & lines(i).Trim
                        End If

                        lngCount += 1
                    End If
                    If bolSelectCase = True Then
                        strTest = lines(i)
                    End If
                Else
                    If lines(i).Trim.Last <> "_" Then bolUnderscore = False
                    'If lines(i).Contains("Case") Then bolSelect = False
                    strTest = lines(i)
                End If
            Else
                strTest = lines(i)
            End If


            strReturn &= IIf(strReturn = "", "", vbCrLf) & strTest
        Next
        Return strReturn
    End Function
End Class
