Imports System.Data.Common

Public Class Indentation
    Public ColIndentLevel As New Collection

    Public Function IndentCode(CodeString As String) As String
        Dim lines() As String = CodeString.Split(vbCrLf)
        Dim strTest As String
        Dim strReturn As String = ""
        Dim level As Int16 = 1
        Dim intIndentNextLine As Int16 = 0
        For i As Int16 = 0 To lines.Count - 1
            Dim strLine() As String = lines(i).Trim.Split(" ")
            Select Case strLine(0).ToLower
                Case "if"
                    If strLine(strLine.Count - 1).ToLower = "then" Then
                        ColIndentLevel.Add("if")
                        level = ColIndentLevel.Count
                    Else
                        Dim iThen As Int16
                        Dim iC As Int16
                        Dim blnNotBlank As Boolean = False
                        For iT As Int16 = 0 To strLine.Count - 1
                            If strLine(iT).ToLower = "then" Then iThen = iT
                            If iThen > 0 And strLine(iT) <> "" Then blnNotBlank = True
                            If strLine(iT).Length > 0 Then
                                If strLine(iT).Substring(0, 1) = "'" Then iC = iT
                            End If
                            If iThen < iC Then
                                If iC - iThen > 2 And blnNotBlank = False Then
                                    level = ColIndentLevel.Count + 1
                                Else
                                    ColIndentLevel.Add("if")
                                    level = ColIndentLevel.Count
                                End If
                                Exit For
                            End If
                        Next
                    End If
                Case "else", "else"
                    level = ColIndentLevel.Count
                Case "select", "with", "for", "do", "while"
                    ColIndentLevel.Add(strLine(0).ToLower)
                    level = ColIndentLevel.Count
                Case "case"
                    If ColIndentLevel.Item(ColIndentLevel.Count) = "select" Then
                        ColIndentLevel.Add("case")
                    End If
                    level = ColIndentLevel.Count
                Case "end"
                    Select Case strLine(1).ToLower
                        Case "if", "with"
                            level = ColIndentLevel.Count
                            ColIndentLevel.Remove(ColIndentLevel.Count)
                        Case "select"
                            ColIndentLevel.Remove(ColIndentLevel.Count) 'remove Case
                            level = ColIndentLevel.Count
                            ColIndentLevel.Remove(ColIndentLevel.Count) 'remove Select
                    End Select
                Case "next", "loop", "wend"
                    level = ColIndentLevel.Count
                    ColIndentLevel.Remove(ColIndentLevel.Count)
                Case Else
                    level = ColIndentLevel.Count + 1
            End Select
            strTest = StrDup(4 * (level + intIndentNextLine), " ") & lines(i).Trim
            strReturn &= IIf(strReturn = "", "", vbCrLf) & strTest
            If strLine(strLine.Count - 1).ToLower = "_" Then
                intIndentNextLine = 1
            Else
                intIndentNextLine = 0
            End If
        Next
        Return strReturn
    End Function
End Class
