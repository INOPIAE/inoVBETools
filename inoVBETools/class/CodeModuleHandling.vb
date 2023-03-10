Imports System.Drawing
Imports System.IO
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.AxHost
Imports Microsoft
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop
Public Class CodeModuleHandling

    Private ClsGit As New GitHandling(My.Settings.Git_Exe)
    Public Sub ImportCodeModule(vbeProject As VBProject, ModuleFullPath As String, Optional blnMessage As Boolean = False)

        Dim ModuleName As String = getModuleNameFromPath(ModuleFullPath)

        If ModuleExists(ModuleName, vbeProject) And blnMessage Then
            If MessageBox.Show(String.Format(inoVBETools.My.Resources.CMH_ModuleImported, ModuleName) & vbCrLf & inoVBETools.My.Resources.CMH_Replace, inoVBETools.My.Resources.Msg_Hint, MessageBoxButtons.YesNo) = vbYes Then
                vbeProject.VBComponents.Remove(GetComponentByName(ModuleName, vbeProject))
            Else
                Exit Sub
            End If
        ElseIf ModuleExists(ModuleName, vbeProject) Then
            vbeProject.VBComponents.Remove(GetComponentByName(ModuleName, vbeProject))
        End If
        vbeProject.VBComponents.Import(ModuleFullPath)
    End Sub

    Public Function ModuleExists(Modulename As String, vbeProject As VBProject) As Boolean
        For Each vbmodule As VBComponent In vbeProject.VBComponents
            If vbmodule.Name.ToLower = Modulename.ToLower Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetComponentByName(Modulename As String, vbeProject As VBProject) As VBComponent
        For Each vbmodule As VBComponent In vbeProject.VBComponents
            If vbmodule.Name.ToLower = Modulename.ToLower Then
                Return vbmodule
            End If
        Next
        Return Nothing
    End Function

    Public Sub ExportModules(vbeProject As VBProject, strPath As String)

        ExportModules(vbeProject, strPath, "")

    End Sub

    Public Sub ExportModules(vbeProject As VBProject, strPath As String, strDate As String)
        Dim LFiles As New List(Of String)
        Dim strFilename As String
        For Each vbmodule As VBComponent In vbeProject.VBComponents
            Dim strExtension As String = ""
            Select Case vbmodule.Type
                Case vbext_ComponentType.vbext_ct_StdModule
                    strExtension = ".bas"
                Case vbext_ComponentType.vbext_ct_ClassModule
                    strExtension = ".cls"
                Case vbext_ComponentType.vbext_ct_Document
                    strExtension = ".dcls"
                Case vbext_ComponentType.vbext_ct_MSForm
                    strExtension = ".frm"
            End Select

            If strExtension <> "" Then
                strFilename = strPath & vbmodule.Name & strExtension
                vbmodule.Export(strFilename)
                LFiles.Add(strFilename)
                If strDate <> "" Then
                    strDate = "_" & strDate
                    My.Computer.FileSystem.RenameFile(strFilename, vbmodule.Name & strDate & strExtension)
                End If
            End If
        Next
        If strDate = "" Then
            Dim di As New IO.DirectoryInfo(strPath)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.*")
            Dim fi As IO.FileInfo

            For Each fi In aryFi
                Select Case fi.Extension
                    Case ".cls", ".frm", ".bas", ".dcls"
                        If Not LFiles.Contains(fi.FullName) Then
                            If MessageBox.Show(String.Format(inoVBETools.My.Resources.msgDeleteFileExport, fi.Name, vbCrLf), "inoVBETools", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                fi.Delete()
                            End If
                        End If
                End Select
            Next
        End If

    End Sub

    Public Sub ImportModules(vbeProject As VBProject, strPath As String)
        If My.Settings.MakeBackup Then MakeBackup(vbeProject, strPath)
        Dim LFiles As New List(Of String)
        Dim strWorksheet As String = ""
        If MessageBox.Show(String.Format(inoVBETools.My.Resources.CMHOverwrite, vbeProject.Name) & vbCrLf & inoVBETools.My.Resources.msgContinue, inoVBETools.My.Resources.CHMTitleImport, MessageBoxButtons.YesNo) = vbYes Then
            Dim di As New IO.DirectoryInfo(strPath)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.*")
            Dim fi As IO.FileInfo

            For Each fi In aryFi
                Select Case fi.Extension
                    Case ".cls", ".frm", ".bas"
                        ImportCodeModule(vbeProject, fi.FullName)
                        LFiles.Add(fi.Name.Replace(fi.Extension, ""))
                    Case ".dcls"
                        If ImportCodeModuleSpecial(vbeProject, fi.FullName) Then
                            strWorksheet &= fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length) & vbCrLf
                        End If
                        LFiles.Add(fi.Name.Replace(fi.Extension, ""))
                End Select
            Next

            For Each vbmodule As VBComponent In vbeProject.VBComponents
                If vbmodule.Type <> vbext_ComponentType.vbext_ct_Document Then
                    If Not LFiles.Contains(vbmodule.Name) Then
                        If MessageBox.Show(String.Format(inoVBETools.My.Resources.msgDeleteModuleImport, vbmodule.Name, vbCrLf), "inoVBETools", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                            vbeProject.VBComponents.Remove(vbmodule)
                        End If
                    End If
                End If
            Next
        End If

        If strWorksheet <> "" Then
            MessageBox.Show(inoVBETools.My.Resources.msgCodeWorksIndetended & vbCrLf & strWorksheet)
        End If
    End Sub

    Function ComponentTypeToString(ComponentType As vbext_ComponentType) As String
        'ComponentTypeToString from http://www.cpearson.com/excel/vbe.aspx
        Select Case ComponentType


            Case vbext_ComponentType.vbext_ct_ActiveXDesigner
                Return "ActiveX Designer"

            Case vbext_ComponentType.vbext_ct_ClassModule
                Return "Class Module"

            Case vbext_ComponentType.vbext_ct_Document
                Return "Document Module"

            Case vbext_ComponentType.vbext_ct_MSForm
                Return "UserForm"

            Case vbext_ComponentType.vbext_ct_StdModule
                Return "Code Module"

            Case Else
                Return "Unknown Type: " & CStr(ComponentType)

        End Select

    End Function

    Private Function ImportCodeModuleSpecial(vbeProject As VBProject, strPath As String) As Boolean
        Dim ModuleName As String = getModuleNameFromPath(strPath)
        Dim reader As New StreamReader(strPath, Encoding.Default)
        Dim intLines As Int16 = 1
        Dim blnWorksheet As Boolean = False


        Dim CodeComponent As VBComponent = GetComponentByName(ModuleName, vbeProject)
        If IsNothing(CodeComponent) Then
            MessageBox.Show(String.Format(inoVBETools.My.Resources.CHM_ProblemImport, ModuleName))
        Else
            CodeComponent.CodeModule.DeleteLines(1, CodeComponent.CodeModule.CountOfLines)

            Dim strCode As String = ""
            Do Until reader.EndOfStream
                Dim codeline As String = reader.ReadLine
                If intLines > 9 Then
                    strCode &= codeline & vbCrLf
                End If
                intLines += 1
            Loop
            If intLines > 12 Then blnWorksheet = True
            reader.Close()

            CodeComponent.CodeModule.InsertLines(1, strCode)
        End If
        Return blnWorksheet
    End Function

    Public Function getModuleNameFromPath(strPath As String) As String
        Dim ModuleNameA() As String = strPath.Split("\")
        Dim ModuleName() As String = ModuleNameA.Last.Split(".")
        Return ModuleName.First
    End Function

    Public Sub MakeBackup(vbeProject As VBProject, strPath As String)
        strPath = Path.Combine(strPath, "backup_code")

        If My.Settings.KeepBackup = False Then
            Directory.Delete(strPath, True)
        End If

        If Not Directory.Exists(strPath) Then
            Directory.CreateDirectory(strPath)
            ClsGit.AppendToGitIgnoreFile(strPath, ".frm" & vbCrLf)
            ClsGit.AppendToGitIgnoreFile(strPath, ".frx" & vbCrLf)
            ClsGit.AppendToGitIgnoreFile(strPath, ".cls" & vbCrLf)
            ClsGit.AppendToGitIgnoreFile(strPath, ".bas" & vbCrLf)
            ClsGit.AppendToGitIgnoreFile(strPath, ".dcls" & vbCrLf)
        End If
        Dim strDate As String = Format(Now, "yyyyMMdd_hhmm")

        ExportModules(vbeProject, strPath & "\", strDate)
    End Sub

    Public Function CheckProjectHasName(vbeProject As VBProject) As Boolean
        ' Default project namne of Office Applications
        ' VBAProject - Excel, MS Project, Powerpoint
        ' Project - Word
        ' Access has no default project name
        ' TODO Check Visio, Outlook
        Select Case vbeProject.Name
            Case "VBAProject", "Project"
                MessageBox.Show(inoVBETools.My.Resources.msgProjectHasNoSpecificName & vbCrLf _
                                & inoVBETools.My.Resources.msgUseThisFunction & vbCrLf & inoVBETools.My.Resources.msgActionCanceled)
                Return False
        End Select
        Return True
    End Function

    Public Function UpdateVersion(vbeProject As VBProject) As Boolean
        Dim CM As VBComponent = GetComponentByName("mdl_Version", vbeProject)
        If IsNothing(CM) Then
            If MessageBox.Show(String.Format("No module 'mdl_Version found'.{0}Do you want to import it?", vbCrLf), My.Resources.Msg_Hint, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Dim tempFile As String = Path.Combine(Path.GetTempPath(), "mdl_Version.bas")
                FileCopy(Path.Combine(My.Application.Info.DirectoryPath, "ressources\vbafiles\mdl_Version.txt"), tempFile)
                vbeProject.VBComponents.Import(tempFile)
                File.Delete(tempFile)
                CM = GetComponentByName("mdl_Version", vbeProject)
            Else
                Return False
            End If

        End If

        With CM.CodeModule
            For intLine = 1 To .CountOfDeclarationLines
                Dim TestString As String = .Lines(intLine, 1)
                Dim strTest() As String
                Dim intMajor As Integer
                Dim intMinor As Integer
                Dim VersionDate As String
                If TestString.Contains("sub") Then Exit For
                If TestString.Contains("function") Then Exit For

                If TestString.Contains("MinorVersion") Then
                    strTest = TestString.Split("=")
                    intMinor = strTest(1).Trim + 1
                    .ReplaceLine(intLine, TestString.Replace(intMinor - 1, intMinor))
                End If

                If TestString.Contains("MajorVersion") Then
                    strTest = TestString.Split("=")
                    intMajor = strTest(1).Trim + 1
                End If

                If TestString.Contains("VersionDate") Then
                    strTest = TestString.Split("=")
                    VersionDate = Date.Today.Month & "/" & Date.Today.Day & "/" & Date.Today.Year
                    .ReplaceLine(intLine, TestString.Replace(strTest(1), "#" & VersionDate & "#"))
                End If
            Next
        End With
        Return True
    End Function
End Class
