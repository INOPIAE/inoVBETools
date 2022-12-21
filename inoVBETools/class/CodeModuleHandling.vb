Imports System.IO
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.AxHost
Imports Microsoft
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop
Public Class CodeModuleHandling

    Private ClsGit As New GitHandling
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
                        ImportCodeModuleSpecial(vbeProject, fi.FullName)
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

    Private Sub ImportCodeModuleSpecial(vbeProject As VBProject, strPath As String)
        Dim ModuleName As String = getModuleNameFromPath(strPath)
        Dim reader As New StreamReader(strPath, Encoding.Default)
        Dim intLines As Int16 = 1


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
            reader.Close()

            CodeComponent.CodeModule.InsertLines(1, strCode)
        End If
    End Sub

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
            ClsGit.AppendToGitIgnoreFile(strPath, ".frm")
            ClsGit.AppendToGitIgnoreFile(strPath, ".frx")
            ClsGit.AppendToGitIgnoreFile(strPath, ".cls")
            ClsGit.AppendToGitIgnoreFile(strPath, ".bas")
            ClsGit.AppendToGitIgnoreFile(strPath, ".dcls")
        End If
        Dim strDate As String = Format(Now, "yyyyMMdd_hhmm")

        ExportModules(vbeProject, strPath & "\", strDate)
    End Sub
End Class
