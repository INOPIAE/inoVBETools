Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Windows.Forms
Imports Microsoft
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop
Public Class CodeModuleHandling

    Public Sub ImportCodeModule(vbeProject As VBProject, ModuleFullPath As String, Optional blnMessage As Boolean = False)
        Dim ModuleNameA() As String = ModuleFullPath.Split("\")
        Dim ModuleName() As String = ModuleNameA.Last.Split(".")

        If ModuleExists(ModuleName.First, vbeProject) And blnMessage Then
            If MessageBox.Show(String.Format(inoVBETools.My.Resources.CMH_ModuleImported, ModuleName.First) & vbCrLf & inoVBETools.My.Resources.CMH_Replace, inoVBETools.My.Resources.Msg_Hint, MessageBoxButtons.YesNo) = vbYes Then
                vbeProject.VBComponents.Remove(GetModuleByName(ModuleName.First, vbeProject))
            Else
                Exit Sub
            End If

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

    Public Function GetModuleByName(Modulename As String, vbeProject As VBProject) As VBComponent
        For Each vbmodule As VBComponent In vbeProject.VBComponents
            If vbmodule.Name.ToLower = Modulename.ToLower Then
                Return vbmodule
            End If
        Next
        Return Nothing
    End Function

    Public Sub ExportModules(vbeProject As VBProject, strPath As String)

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
                vbmodule.Export(strPath & vbmodule.Name & strExtension)
            End If

        Next
    End Sub

    Public Sub ImportModules(vbeProject As VBProject, strPath As String)
        If MessageBox.Show("All existing code modules will be overwritten. Do you want to continue", "Import Modules", MessageBoxButtons.YesNo) = vbYes Then
            Dim di As New IO.DirectoryInfo(strPath)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.*")
            Dim fi As IO.FileInfo

            For Each fi In aryFi
                Select Case fi.Extension
                    Case ".cls", ".frm", ".bas"
                        ImportCodeModule(vbeProject, fi.FullName)
                End Select
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
End Class
