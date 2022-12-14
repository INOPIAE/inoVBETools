Imports System
Imports Xunit

Namespace inoVBEToolsTest

    Public Class UnitTestIndentation
        Private clsI As New inoVBETools.Indentation

        <Fact>
        Sub TestIndentIf()

            Dim strStart As String

            strStart = "    If test = 1 Then" & vbCrLf _
                & "       test = 2" & vbCrLf _
                & "    Else" & vbCrLf _
                & "        test = 2" & vbCrLf _
                & "    End If"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)
            TestIndentationLevel(strtest(3), 2)
            TestIndentationLevel(strtest(4), 1)
        End Sub

        <Fact>
        Sub TestIndentIfLine()
            Dim strStart As String

            strStart = "    If test = 1 Then test = 2" & vbCrLf _
                & "    test = 2"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)
            Dim strtest() As String = strResult.Split(vbCrLf)
            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 1)

            strStart = "    If test = 1 Then test = 2 'bla" & vbCrLf _
                & "    test = 2"


            strResult = clsI.IndentCode(strStart)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 1)

        End Sub

        <Fact>
        Sub TestIndentIfNested()

            Dim strStart As String

            strStart = "    If test = 1 Then" & vbCrLf _
                & "       test = 2" & vbCrLf _
                & "    Else" & vbCrLf _
                & "        If test = 2 Then" & vbCrLf _
                & "            test = 3" & vbCrLf _
                & "        Else" & vbCrLf _
                & "            Test = 4" & vbCrLf _
                & "        End If" & vbCrLf _
                & "        test = 2" & vbCrLf _
                & "    End If"


            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)
            TestIndentationLevel(strtest(3), 2)
            TestIndentationLevel(strtest(4), 3)
            TestIndentationLevel(strtest(5), 2)
            TestIndentationLevel(strtest(6), 3)
            TestIndentationLevel(strtest(7), 2)
            TestIndentationLevel(strtest(8), 2)
            TestIndentationLevel(strtest(9), 1)
        End Sub

        <Fact>
        Sub TestIndentIfElseIF()

            Dim strStart As String

            strStart = "    If test = 1 Then" & vbCrLf _
                & "       test = 2" & vbCrLf _
                & "    Else" & vbCrLf _
                & "        test = 2" & vbCrLf _
                & "    Else If" & vbCrLf _
                & "        test = 2" & vbCrLf _
                & "    End If"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)
            TestIndentationLevel(strtest(3), 2)
            TestIndentationLevel(strtest(4), 1)
            TestIndentationLevel(strtest(5), 2)
            TestIndentationLevel(strtest(6), 1)
        End Sub

        <Fact>
        Sub TestIndentIfComment()

            Dim strStart As String

            strStart = "    If test = 1 Then ' bla" & vbCrLf _
                & "       test = 2" & vbCrLf _
                & "    Else" & vbCrLf _
                & "        test = 2" & vbCrLf _
                & "    End If"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)
            TestIndentationLevel(strtest(3), 2)
            TestIndentationLevel(strtest(4), 1)

            strStart = "    If test = 1 Then     ' bla" & vbCrLf _
    & "       test = 2" & vbCrLf _
    & "    Else" & vbCrLf _
    & "        test = 2" & vbCrLf _
    & "    End If"


            strResult = clsI.IndentCode(strStart)
            Debug.Print(strStart)
            Debug.Print(strResult)


            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)
            TestIndentationLevel(strtest(3), 2)
            TestIndentationLevel(strtest(4), 1)
        End Sub
        <Fact>
        Sub TestIndentSelect()

            Dim strStart As String

            strStart = "    Select Case foo" & vbCrLf _
                & "        Case 1" & vbCrLf _
                & "            test = 1" & vbCrLf _
                & "        Case 2" & vbCrLf _
                & "            test = 2" & vbCrLf _
                & "    End Select"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)
            Debug.Print(strStart)
            Debug.Print(strResult)
            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 3)
            TestIndentationLevel(strtest(3), 2)
            TestIndentationLevel(strtest(4), 3)
            TestIndentationLevel(strtest(5), 1)
        End Sub

        <Fact>
        Sub TestIndentSelectIf()

            Dim strStart As String

            strStart = "    Select Case foo" & vbCrLf _
                & "        Case 1" & vbCrLf _
                & "            test = 1" & vbCrLf _
                & "        Case 2" & vbCrLf _
                & "            If test = 2 Then" & vbCrLf _
                & "                test = 2" & vbCrLf _
                & "            Else" & vbCrLf _
                & "                test = 2" & vbCrLf _
                & "            End If" & vbCrLf _
                & "    End Select"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)
            Debug.Print(strStart)
            Debug.Print(strResult)
            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 3)
            TestIndentationLevel(strtest(3), 2)
            TestIndentationLevel(strtest(4), 3)
            TestIndentationLevel(strtest(5), 4)
            TestIndentationLevel(strtest(6), 3)
            TestIndentationLevel(strtest(7), 4)
            TestIndentationLevel(strtest(8), 3)
            TestIndentationLevel(strtest(9), 1)
        End Sub

        <Fact>
        Sub TestIndentFor()

            Dim strStart As String

            strStart = "    For i = 1 to 4" & vbCrLf _
                & "        test = 2" & vbCrLf _
                & "    Next"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)
            Debug.Print(strStart)
            Debug.Print(strResult)
            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)
        End Sub

        <Fact>
        Sub TestIndentWith()

            Dim strStart As String

            strStart = "    With test" & vbCrLf _
                & "       .test = 2" & vbCrLf _
                & "    End With"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)

        End Sub

        <Fact>
        Sub TestIndentDo()

            Dim strStart As String

            strStart = "    Do" & vbCrLf _
                & "       test = 2" & vbCrLf _
                & "    Loop"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)

        End Sub

        <Fact>
        Sub TestIndentWhile()

            Dim strStart As String

            strStart = "    Msgbox ""This is a message"" _" & vbCrLf _
                & "       & ""more text"" & _" & vbCrLf _
                & "        ""next line""" & vbCrLf _
                & "    test=1"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 2)
            TestIndentationLevel(strtest(3), 1)

        End Sub

        <Fact>
        Sub TestIndentUnderscore()

            Dim strStart As String

            strStart = "    While Counter < 20" & vbCrLf _
                & "       Counter = Counter + 1 " & vbCrLf _
                & "    Wend"

            Dim strResult As String
            strResult = clsI.IndentCode(strStart)

            Dim strtest() As String = strResult.Split(vbCrLf)

            TestIndentationLevel(strtest(0), 1)
            TestIndentationLevel(strtest(1), 2)
            TestIndentationLevel(strtest(2), 1)

        End Sub

        <Fact>
        Sub TestIndentLineNumbers()

            Dim strStart As String

            strStart = "1   If test = 1 Then" & vbCrLf _
            & "2      test = 2" & vbCrLf _
            & "3   Else" & vbCrLf _
            & "4       test = 2" & vbCrLf _
            & "5   End If"


            Dim strResult As String
            strResult = clsI.IndentCode(strStart)
            Debug.Print(strStart & "|")
            Debug.Print(strResult & "|")

            Dim strtest() As String = strResult.Split(vbCrLf)
            TestIndentationLevelNumber(strtest(0), 1, 1)
            TestIndentationLevelNumber(strtest(1), 2, 2)
            TestIndentationLevelNumber(strtest(2), 1, 3)
            TestIndentationLevelNumber(strtest(3), 2, 4)
            TestIndentationLevelNumber(strtest(4), 1, 5)


        End Sub
        Private Shared Sub TestIndentationLevel(strtest As String, Level As Int16)
            Assert.StartsWith(StrDup(4 * Level, " "), strtest)
            Assert.NotEqual(" ", strtest.Substring(4 * Level, 1))
        End Sub
        Private Shared Sub TestIndentationLevelNumber(strtest As String, Level As Int16, Line As Int16)
            Dim strNumTest As String = StrDup(4 * Level, " ")
            If Line = 0 Then
                strNumTest = strNumTest
            Else
                strNumTest = Line & strNumTest.Substring(Line.ToString.Length)
            End If
            Assert.StartsWith(strNumTest, strtest)
            Assert.NotEqual(" ", strtest.Substring(4 * Level, 1))
        End Sub
    End Class
End Namespace

