Imports inoVBETools
Imports Xunit

Namespace inoVBEToolsTest
    Public Class UnitTestLineNumbering
        Private ClsLineNumbering As New LineNumbering

        <Fact>
        Sub TestLineNumberingIf()
            Dim strStart As String

            strStart = "    If test = 1 Then" & vbCrLf _
            & "       test = 2" & vbCrLf _
            & "    Else" & vbCrLf _
            & "        test = 2" & vbCrLf _
            & "    End If"

            Dim strResult As String
            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strStart)

            Dim strTest As String = "1   If test = 1 Then" & vbCrLf _
            & "2      test = 2" & vbCrLf _
            & "3   Else" & vbCrLf _
            & "4       test = 2" & vbCrLf _
            & "5   End If"
            Assert.Equal(strTest, strResult)

            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strResult, True)
            Assert.Equal(strStart, strResult)
        End Sub

        <Fact>
        Sub TestLineNumberingSelect()
            Dim strStart As String

            strStart = "    Select Case foo" & vbCrLf _
                & "        Case 1" & vbCrLf _
                & "            test = 1" & vbCrLf _
                & "        Case 2" & vbCrLf _
                & "            test = 2" & vbCrLf _
                & "    End Select"

            Dim strResult As String
            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strStart)

            Dim strTest As String = "1   Select Case foo" & vbCrLf _
                & "        Case 1" & vbCrLf _
                & "2           test = 1" & vbCrLf _
                & "        Case 2" & vbCrLf _
                & "3           test = 2" & vbCrLf _
                & "4   End Select"
            Assert.Equal(strTest, strResult)

            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strResult, True)
            Assert.Equal(strStart, strResult)
        End Sub

        <Fact>
        Sub TestLineNumberingComment()
            Dim strStart As String

            strStart = "    Test = 1" & vbCrLf _
                & "     'Test = 2" & vbCrLf _
                & "     Test = 3" & vbCrLf _
                & "     Test = 4 ' Comment" & vbCrLf _
                & "' Test 5"

            Dim strResult As String
            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strStart)

            Dim strTest As String = "1   Test = 1" & vbCrLf _
                & "     'Test = 2" & vbCrLf _
                & "2    Test = 3" & vbCrLf _
                & "3    Test = 4 ' Comment" & vbCrLf _
                & "' Test 5"
            Assert.Equal(strTest, strResult)

            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strResult, True)
            Assert.Equal(strStart, strResult)
        End Sub

        <Fact>
        Sub TestLineNumberingBlankLines()

            Dim strStart As String

            strStart = "    Test = 1" & vbCrLf _
                & "" & vbCrLf _
                & "     Test = 3" & vbCrLf _
                & "     " & vbCrLf _
                & "     Test = 4 ' Comment" & vbCrLf _
                & "     Test 5"

            Dim strResult As String
            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strStart)
            Debug.Print(strResult)
            Dim strTest As String = "1   Test = 1" & vbCrLf _
                & "" & vbCrLf _
                & "2    Test = 3" & vbCrLf _
                & "     " & vbCrLf _
                & "3    Test = 4 ' Comment" & vbCrLf _
                & "4    Test 5"
            Assert.Equal(strTest, strResult)

            strResult = ClsLineNumbering.AddLineNumbersToCurrentProcedure(strResult, True)
            Assert.Equal(strStart, strResult)
        End Sub

        <Fact>
        Sub TestLineNumberingExists()

            Dim strStart As String

            strStart = "    Test = 1" & vbCrLf _
                & "" & vbCrLf _
                & "     Test = 3" & vbCrLf _
                & "     " & vbCrLf _
                & "     Test = 4 ' Comment" & vbCrLf _
                & "     Test 5"

            Assert.False(ClsLineNumbering.HasLineNumbersInCode(strStart))

            strStart = "1    Test = 1" & vbCrLf _
               & "" & vbCrLf _
               & "     Test = 3" & vbCrLf _
               & "     " & vbCrLf _
               & "     Test = 4 ' Comment" & vbCrLf _
               & "     Test 5"
            Assert.True(ClsLineNumbering.HasLineNumbersInCode(strStart))

            strStart = "'    Test = 1" & vbCrLf _
               & "1  Test=1a" & vbCrLf _
               & "     Test = 3" & vbCrLf _
               & "     " & vbCrLf _
               & "     Test = 4 ' Comment" & vbCrLf _
               & "     Test 5"
            Assert.True(ClsLineNumbering.HasLineNumbersInCode(strStart))
        End Sub
    End Class
End Namespace
